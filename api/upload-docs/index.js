// api/upload-docs/index.js
// Azure Function — POST /api/upload-docs
// Returning clients adding more documents. Verifies client exists,
// uploads to existing OneDrive folder, appends Excel note, notifies Teams.

const { getAccessToken, graphRequest, uploadFileToOneDrive } = require('../shared/_graph');
const { sendTeamsNotification } = require('../shared/_teams');
const { parseMultipartBuffer } = require('../shared/_multipart');

const ONEDRIVE_ROOT = process.env.ONEDRIVE_FOLDER || 'Mownur Clients';
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const ONEDRIVE_USER_ID = process.env.ONEDRIVE_USER_ID;

module.exports = async function (context, req) {
  if (req.method === 'OPTIONS') {
    context.res = { status: 200, headers: corsHeaders() };
    return;
  }

  let fields, files;
  try {
    ({ fields, files } = parseMultipartBuffer(req.rawBody, req.headers['content-type']));
  } catch (err) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'Failed to parse form data', detail: err.message } };
    return;
  }

  const { firstName, lastName, last4 } = fields;
  if (!firstName || !lastName || !last4) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'Missing required fields: firstName, lastName, last4' } };
    return;
  }
  if (!/^\d{4}$/.test(last4)) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'last4 must be exactly 4 digits' } };
    return;
  }
  if (!files?.length) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'No files provided' } };
    return;
  }

  let token;
  try {
    token = await getAccessToken();
  } catch (err) {
    context.log.error('Auth error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Authentication failed' } };
    return;
  }

  // Verify client exists
  let clientRow, rowIndex;
  try {
    ({ row: clientRow, index: rowIndex } = await findClientRow(token, firstName, lastName, last4));
  } catch (err) {
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Could not verify client' } };
    return;
  }
  if (!clientRow) {
    context.res = {
      status: 404,
      headers: corsHeaders(),
      body: { error: 'not_found', message: "We couldn't find a return with that information." },
    };
    return;
  }

  // Find existing OneDrive folder
  const folderName = `${firstName} ${lastName} - ${last4}`;
  const driveBase = ONEDRIVE_USER_ID
    ? `/users/${encodeURIComponent(ONEDRIVE_USER_ID)}/drive`
    : `/me/drive`;

  let folderId;
  try {
    const folder = await graphRequest(token, 'GET',
      `${driveBase}/root:/${encodeURIComponent(ONEDRIVE_ROOT)}/${encodeURIComponent(folderName)}`
    );
    folderId = folder.id;
  } catch (err) {
    context.log.error('OneDrive folder lookup error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Could not locate client OneDrive folder' } };
    return;
  }

  // Upload files
  const uploadResults = [];
  for (const file of files) {
    let attempt = 0, success = false;
    while (attempt < 3 && !success) {
      try {
        await uploadFileToOneDrive(token, folderId, file.category || 'Additional Documents', file);
        uploadResults.push({ name: file.filename, status: 'uploaded' });
        success = true;
      } catch (err) {
        attempt++;
        if (attempt < 3) await sleep(attempt * 1000);
        else uploadResults.push({ name: file.filename, status: 'failed', error: err.message });
      }
    }
  }

  // Append note to Excel
  const uploadedCount = uploadResults.filter(r => r.status === 'uploaded').length;
  try {
    await appendExcelNote(token, rowIndex, clientRow,
      `Additional documents received ${new Date().toLocaleDateString('en-US')} (${uploadedCount} file${uploadedCount !== 1 ? 's' : ''})`
    );
  } catch (err) {
    context.log.warn('Excel note update failed:', err.message);
  }

  // Teams notification
  try {
    await sendTeamsNotification({
      type: 'additional_docs',
      clientName: `${firstName} ${lastName}`,
      uploadedCount,
      submittedAt: new Date().toISOString(),
      adminUrl: `${process.env.SITE_URL || ''}/admin.html`,
    });
  } catch (err) {
    context.log.warn('Teams notification failed:', err.message);
  }

  const failedFiles = uploadResults.filter(r => r.status === 'failed');
  context.res = {
    status: 200,
    headers: corsHeaders(),
    body: {
      success: true,
      uploadedFiles: uploadedCount,
      failedFiles: failedFiles.length ? failedFiles : undefined,
      message: failedFiles.length
        ? `${uploadedCount} file(s) uploaded. ${failedFiles.length} failed — please retry.`
        : `${uploadedCount} file(s) uploaded successfully.`,
    },
  };
};

// ---------------------------------------------------------------------------

async function findClientRow(token, firstName, lastName, last4) {
  const response = await graphRequest(token, 'GET',
    `/me/drive/items/${EXCEL_FILE_ID}/workbook/tables/ClientDatabase/rows`
  );
  const rows = response.value || [];
  for (let i = 0; i < rows.length; i++) {
    const v = rows[i].values[0];
    if (
      String(v[0]).toLowerCase().trim() === firstName.toLowerCase().trim() &&
      String(v[1]).toLowerCase().trim() === lastName.toLowerCase().trim() &&
      String(v[16]).trim() === String(last4).trim()
    ) {
      return { row: rows[i], index: i };
    }
  }
  return { row: null, index: -1 };
}

async function appendExcelNote(token, rowIndex, existingRow, newNote) {
  const allValues = [...existingRow.values[0]];
  const current = String(allValues[15] || '');
  allValues[15] = current ? `${current}; ${newNote}` : newNote;
  await graphRequest(token, 'PATCH',
    `/me/drive/items/${EXCEL_FILE_ID}/workbook/tables/ClientDatabase/rows/itemAt(index=${rowIndex})`,
    { values: [allValues] }
  );
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
