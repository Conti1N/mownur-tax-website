// api/submit-intake/index.js
// Azure Function — POST /api/submit-intake
// Receives intake form + files, creates OneDrive folder, uploads files,
// writes Excel row, sends Teams notification + confirmation email.

const { getAccessToken, graphRequest, uploadFileToOneDrive } = require('../shared/_graph');
const { sendTeamsNotification } = require('../shared/_teams');
const { sendEmail } = require('../shared/_email');
const { parseMultipartBuffer } = require('../shared/_multipart');

const ONEDRIVE_ROOT = process.env.ONEDRIVE_FOLDER || 'Mownur Clients';
const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const ONEDRIVE_USER_ID = process.env.ONEDRIVE_USER_ID;

module.exports = async function (context, req) {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    context.res = { status: 200, headers: corsHeaders() };
    return;
  }

  // Parse multipart body
  let fields, files;
  try {
    const contentType = req.headers['content-type'] || '';
    ({ fields, files } = parseMultipartBuffer(req.rawBody, contentType));
  } catch (err) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'Failed to parse form data', detail: err.message } };
    return;
  }

  // Validate required fields
  const required = ['firstName', 'lastName', 'last4', 'email', 'phone', 'filingStatus'];
  const missing = required.filter(f => !fields[f]);
  if (missing.length) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'Missing required fields', fields: missing } };
    return;
  }

  const { firstName, lastName, last4, email, phone, filingStatus } = fields;
  const folderName = `${firstName} ${lastName} - ${last4}`;
  const submittedAt = new Date().toISOString();

  let token;
  try {
    token = await getAccessToken();
  } catch (err) {
    context.log.error('Auth error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Authentication failed' } };
    return;
  }

  // Check for duplicate submission
  try {
    const existing = await findExcelRow(token, firstName, lastName, last4);
    if (existing) {
      context.res = {
        status: 409,
        headers: corsHeaders(),
        body: { error: 'existing_client', message: 'A return already exists. Use the additional documents endpoint.' },
      };
      return;
    }
  } catch (err) {
    context.log.warn('Duplicate check failed:', err.message);
  }

  // Create OneDrive folder
  let oneDriveFolder;
  try {
    oneDriveFolder = await createOneDriveFolder(token, ONEDRIVE_ROOT, folderName);
  } catch (err) {
    context.log.error('OneDrive folder error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Failed to create OneDrive folder' } };
    return;
  }

  // Upload files (3 retries each)
  const uploadResults = [];
  for (const file of files) {
    let attempt = 0, success = false;
    while (attempt < 3 && !success) {
      try {
        await uploadFileToOneDrive(token, oneDriveFolder.id, file.category || 'General', file);
        uploadResults.push({ name: file.filename, status: 'uploaded' });
        success = true;
      } catch (err) {
        attempt++;
        if (attempt < 3) await sleep(attempt * 1000);
        else uploadResults.push({ name: file.filename, status: 'failed', error: err.message });
      }
    }
  }

  // Write Excel row
  const incomeTypes = toCSV(fields.incomeTypes);
  const lifeChanges = toCSV(fields.lifeChanges);
  const dependentsCount = fields.dependentsCount || '0';
  const uploadedCount = uploadResults.filter(r => r.status === 'uploaded').length;
  const folderWebUrl = oneDriveFolder.webUrl || '';

  try {
    await addExcelRow(token, {
      firstName, lastName, last4, email, phone, filingStatus,
      incomeTypes, lifeChanges, dependentsCount, submittedAt,
      status: 'Submitted', statusUpdatedAt: submittedAt,
      underReviewAt: '', filedAt: '', completedAt: '',
      oneDriveFolderUrl: folderWebUrl, notes: '',
    });
  } catch (err) {
    context.log.error('Excel write error:', err.message);
    // Non-fatal — folder and files already uploaded
  }

  // Teams notification
  const adminUrl = `${process.env.SITE_URL || ''}/admin.html`;
  try {
    await sendTeamsNotification({
      clientName: `${firstName} ${lastName}`,
      email, phone, filingStatus, incomeTypes, lifeChanges,
      dependentsCount, uploadedCount, folderUrl: folderWebUrl,
      adminUrl, submittedAt,
    });
  } catch (err) {
    context.log.warn('Teams notification failed:', err.message);
  }

  // Confirmation email
  try {
    await sendEmail({
      to: email,
      subject: 'We received your tax documents — Mownur Services',
      html: confirmationEmailHtml(firstName, folderName),
    });
  } catch (err) {
    context.log.warn('Confirmation email failed:', err.message);
  }

  const failedFiles = uploadResults.filter(r => r.status === 'failed');
  context.res = {
    status: 200,
    headers: corsHeaders(),
    body: {
      success: true,
      portalCode: `${firstName} ${lastName} + last 4: ${last4}`,
      uploadedFiles: uploadedCount,
      failedFiles: failedFiles.length ? failedFiles : undefined,
      message: failedFiles.length
        ? `Submission received. ${failedFiles.length} file(s) failed — please re-upload them.`
        : 'Submission received. Confirmation email on its way.',
    },
  };
};

// ---------------------------------------------------------------------------

async function createOneDriveFolder(token, rootFolder, clientFolder) {
  const userId = ONEDRIVE_USER_ID;
  const base = userId
    ? `/users/${encodeURIComponent(userId)}/drive`
    : `/me/drive`;

  try {
    return await graphRequest(token, 'POST',
      `${base}/root:/${encodeURIComponent(rootFolder)}:/children`,
      { name: clientFolder, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }
    );
  } catch (err) {
    if (err.status === 409) {
      return await graphRequest(token, 'GET',
        `${base}/root:/${encodeURIComponent(rootFolder)}/${encodeURIComponent(clientFolder)}`
      );
    }
    throw err;
  }
}

async function findExcelRow(token, firstName, lastName, last4) {
  const rows = await graphRequest(token, 'GET',
    `/me/drive/items/${EXCEL_FILE_ID}/workbook/tables/ClientDatabase/rows`
  );
  if (!rows?.value) return null;
  return rows.value.find(row => {
    const v = row.values[0];
    return (
      String(v[0]).toLowerCase().trim() === firstName.toLowerCase().trim() &&
      String(v[1]).toLowerCase().trim() === lastName.toLowerCase().trim() &&
      String(v[16]).trim() === String(last4).trim()
    );
  }) || null;
}

async function addExcelRow(token, d) {
  await graphRequest(token, 'POST',
    `/me/drive/items/${EXCEL_FILE_ID}/workbook/tables/ClientDatabase/rows/add`,
    { values: [[
      d.firstName, d.lastName, d.email, d.phone, d.filingStatus,
      d.incomeTypes, d.lifeChanges, d.dependentsCount, d.submittedAt,
      d.status, d.statusUpdatedAt, d.underReviewAt, d.filedAt, d.completedAt,
      d.oneDriveFolderUrl, d.notes, d.last4,
    ]] }
  );
}

function confirmationEmailHtml(firstName, folderName) {
  const portalKey = folderName.replace(/ - \d{4}$/, '');
  return `
    <div style="font-family:Inter,sans-serif;max-width:600px;margin:0 auto;color:#13263a;">
      <h2>Hi ${firstName}, we've got your documents.</h2>
      <p>Thank you for submitting your tax information to Mownur Services. Here's what happens next:</p>
      <ol>
        <li>Our team reviews your submission (usually within 1 business day)</li>
        <li>We prepare your return — you'll receive updates at each step</li>
        <li>Once filed, you'll get a confirmation email</li>
      </ol>
      <p><strong>Most returns are completed within 5 business days.</strong></p>
      <hr style="border:1px solid #e5e7eb;margin:24px 0;"/>
      <p style="font-size:14px;color:#6b7280;">
        Your portal access: <strong>${portalKey} + your last 4 SSN digits</strong><br/>
        Use this anytime to check your status on our website.
      </p>
      <p style="font-size:14px;color:#6b7280;">Questions? Reply here or call us. — Mownur Services, Minneapolis</p>
    </div>`;
}

function toCSV(val) {
  return Array.isArray(val) ? val.join(', ') : (val || '');
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
