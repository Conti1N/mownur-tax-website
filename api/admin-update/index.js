// api/admin-update/index.js
// Azure Function — POST /api/admin-update
// Admin-only. Verifies password, updates client status in Excel,
// auto-stamps date milestone, emails client about the change.

const { getAccessToken, graphRequest } = require('../shared/_graph');
const { sendEmail } = require('../shared/_email');

const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;
const VALID_STATUSES = ['Submitted', 'Under Review', 'Filed', 'Completed'];

const COL = {
  firstName: 0, lastName: 1, email: 2, phone: 3,
  filingStatus: 4, incomeTypes: 5, lifeChanges: 6, dependentsCount: 7,
  submittedAt: 8, status: 9, statusUpdatedAt: 10,
  underReviewAt: 11, filedAt: 12, completedAt: 13,
  oneDriveFolderUrl: 14, notes: 15, last4: 16,
};

const STATUS_DATE_COL = {
  'Under Review': COL.underReviewAt,
  'Filed': COL.filedAt,
  'Completed': COL.completedAt,
};

module.exports = async function (context, req) {
  if (req.method === 'OPTIONS') {
    context.res = { status: 200, headers: corsHeaders() };
    return;
  }

  const body = req.body || {};
  const { adminPassword, firstName, lastName, last4, newStatus, adminNote } = body;

  // Auth — compare password from env
  const storedPassword = process.env.ADMIN_PASSWORD;
  if (!adminPassword || !storedPassword || adminPassword !== storedPassword) {
    context.res = { status: 401, headers: corsHeaders(), body: { error: 'Unauthorized' } };
    return;
  }

  if (!firstName || !lastName || !last4 || !newStatus) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'Missing required fields: firstName, lastName, last4, newStatus' } };
    return;
  }
  if (!/^\d{4}$/.test(last4)) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'last4 must be exactly 4 digits' } };
    return;
  }
  if (!VALID_STATUSES.includes(newStatus)) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: `newStatus must be one of: ${VALID_STATUSES.join(', ')}` } };
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

  // Find client row
  let clientRow, rowIndex;
  try {
    ({ row: clientRow, index: rowIndex } = await findClientRow(token, firstName, lastName, last4));
  } catch (err) {
    context.log.error('Excel lookup error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Could not read client database' } };
    return;
  }
  if (!clientRow) {
    context.res = { status: 404, headers: corsHeaders(), body: { error: 'Client not found' } };
    return;
  }

  // Build updated row
  const now = new Date().toISOString();
  const updatedValues = [...clientRow.values[0]];
  updatedValues[COL.status] = newStatus;
  updatedValues[COL.statusUpdatedAt] = now;

  // Stamp milestone date (only first time)
  if (STATUS_DATE_COL[newStatus] !== undefined && !updatedValues[STATUS_DATE_COL[newStatus]]) {
    updatedValues[STATUS_DATE_COL[newStatus]] = now;
  }

  // Append admin note
  if (adminNote) {
    const existing = String(updatedValues[COL.notes] || '');
    const noteDate = new Date().toLocaleDateString('en-US');
    updatedValues[COL.notes] = existing
      ? `${existing}; [${noteDate}] ${adminNote}`
      : `[${noteDate}] ${adminNote}`;
  }

  // Update Excel
  try {
    await graphRequest(token, 'PATCH',
      `/me/drive/items/${EXCEL_FILE_ID}/workbook/tables/ClientDatabase/rows/itemAt(index=${rowIndex})`,
      { values: [updatedValues] }
    );
  } catch (err) {
    context.log.error('Excel update error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Failed to update client record' } };
    return;
  }

  // Email client
  const clientEmail = String(clientRow.values[0][COL.email]);
  try {
    await sendEmail({
      to: clientEmail,
      subject: statusEmailSubject(newStatus),
      html: statusEmailHtml(firstName, newStatus, adminNote),
    });
  } catch (err) {
    context.log.warn('Status email failed:', err.message);
  }

  context.res = {
    status: 200,
    headers: corsHeaders(),
    body: { success: true, client: `${firstName} ${lastName}`, newStatus, updatedAt: now },
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
      String(v[COL.firstName]).toLowerCase().trim() === firstName.toLowerCase().trim() &&
      String(v[COL.lastName]).toLowerCase().trim() === lastName.toLowerCase().trim() &&
      String(v[COL.last4]).trim() === String(last4).trim()
    ) {
      return { row: rows[i], index: i };
    }
  }
  return { row: null, index: -1 };
}

function statusEmailSubject(status) {
  return {
    'Under Review': 'Your tax return is under review — Mownur Services',
    'Filed': 'Your tax return has been filed! — Mownur Services',
    'Completed': 'Your tax return is complete — Mownur Services',
  }[status] || `Status update: ${status} — Mownur Services`;
}

function statusEmailHtml(firstName, status, note) {
  const content = {
    'Under Review': { headline: "We're reviewing your return.", body: "Our team has started working on your tax return. We'll notify you when it's filed." },
    'Filed': { headline: 'Your return has been filed!', body: 'Your tax return has been submitted to the IRS. Keep an eye out for your confirmation.' },
    'Completed': { headline: 'Your return is complete.', body: 'Everything is done and finalized. Thank you for choosing Mownur Services.' },
  }[status] || { headline: `Status: ${status}`, body: '' };

  return `
    <div style="font-family:Inter,sans-serif;max-width:600px;margin:0 auto;color:#13263a;">
      <h2>Hi ${firstName}, ${content.headline}</h2>
      <p>${content.body}</p>
      ${note ? `<p style="background:#f3f4f6;padding:12px 16px;border-radius:6px;font-size:14px;">${note}</p>` : ''}
      <hr style="border:1px solid #e5e7eb;margin:24px 0;"/>
      <p style="font-size:14px;color:#6b7280;">
        Check your status anytime at our website using your name and last 4 SSN digits.<br/>
        — Mownur Services, Minneapolis
      </p>
    </div>`;
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };
}
