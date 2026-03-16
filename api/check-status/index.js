// api/check-status/index.js
// Azure Function — GET /api/check-status?first=X&last=Y&last4=Z
// Returns status timeline for exact-match client. Never exposes other clients.

const { getAccessToken, graphRequest } = require('../shared/_graph');

const EXCEL_FILE_ID = process.env.EXCEL_FILE_ID;

module.exports = async function (context, req) {
  if (req.method === 'OPTIONS') {
    context.res = { status: 200, headers: corsHeaders() };
    return;
  }

  const { first, last, last4 } = req.query;

  if (!first || !last || !last4) {
    context.res = {
      status: 400,
      headers: corsHeaders(),
      body: { error: 'Missing required params: first, last, last4' },
    };
    return;
  }

  if (!/^\d{4}$/.test(last4)) {
    context.res = { status: 400, headers: corsHeaders(), body: { error: 'last4 must be exactly 4 digits' } };
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

  let rows;
  try {
    const response = await graphRequest(token, 'GET',
      `/me/drive/items/${EXCEL_FILE_ID}/workbook/tables/ClientDatabase/rows`
    );
    rows = response.value || [];
  } catch (err) {
    context.log.error('Excel read error:', err.message);
    context.res = { status: 500, headers: corsHeaders(), body: { error: 'Could not read client database' } };
    return;
  }

  // Column order: 0:firstName, 1:lastName, 2:email ... 16:last4
  const match = rows.find(row => {
    const v = row.values[0];
    return (
      String(v[0]).toLowerCase().trim() === first.toLowerCase().trim() &&
      String(v[1]).toLowerCase().trim() === last.toLowerCase().trim() &&
      String(v[16]).trim() === last4.trim()
    );
  });

  if (!match) {
    context.res = {
      status: 404,
      headers: corsHeaders(),
      body: {
        found: false,
        message: "We couldn't find a return with that information. Please check your details or contact us.",
      },
    };
    return;
  }

  const v = match.values[0];
  context.res = {
    status: 200,
    headers: corsHeaders(),
    body: {
      found: true,
      clientName: `${v[0]} ${v[1]}`,
      currentStatus: String(v[9] || 'Submitted'),
      timeline: [
        { step: 'Submitted',    date: v[8]  || null, done: true },
        { step: 'Under Review', date: v[11] || null, done: !!v[11] },
        { step: 'Filed',        date: v[12] || null, done: !!v[12] },
        { step: 'Completed',    date: v[13] || null, done: !!v[13] },
      ],
    },
  };
};

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };
}
