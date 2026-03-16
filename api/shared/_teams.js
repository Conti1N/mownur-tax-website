// api/shared/_teams.js
const https = require('https');

async function sendTeamsNotification(opts) {
  const webhookUrl = process.env.TEAMS_WEBHOOK_URL;
  if (!webhookUrl) throw new Error('TEAMS_WEBHOOK_URL not set');

  const type = opts.type || 'new_submission';
  const timestamp = opts.submittedAt
    ? new Date(opts.submittedAt).toLocaleString('en-US', { timeZone: 'America/Chicago' })
    : new Date().toLocaleString('en-US', { timeZone: 'America/Chicago' });

  let text;
  if (type === 'additional_docs') {
    text = [
      `📎 **Additional Documents Received — Mownur Services**`,
      ``,
      `**Client:** ${opts.clientName}`,
      `**Files Uploaded:** ${opts.uploadedCount || 0}`,
      ``,
      `🔗 [View in Admin Panel](${opts.adminUrl || ''})`,
      ``,
      `**Received:** ${timestamp} (CT)`,
    ].join('\n');
  } else {
    const lines = [
      `📋 **New Tax Client Submission — Mownur Services**`,
      ``,
      `**Client:** ${opts.clientName}`,
      `**Email:** ${opts.email || '—'} | **Phone:** ${opts.phone || '—'}`,
      `**Filing Status:** ${opts.filingStatus || '—'}`,
      `**Income Types:** ${opts.incomeTypes || '—'}`,
    ];
    if (opts.lifeChanges) lines.push(`**Life Changes:** ${opts.lifeChanges}`);
    if (opts.dependentsCount && opts.dependentsCount !== '0') {
      lines.push(`**Dependents:** ${opts.dependentsCount}`);
    }
    lines.push(
      `**Documents Uploaded:** ${opts.uploadedCount || 0} files`,
      ``,
      `📁 [OneDrive Folder](${opts.folderUrl || ''})`,
      `🔗 [View in Admin Panel](${opts.adminUrl || ''})`,
      ``,
      `**Submitted:** ${timestamp} (CT)`
    );
    text = lines.join('\n');
  }

  const payload = JSON.stringify({ text });

  return new Promise((resolve, reject) => {
    const urlObj = new URL(webhookUrl);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => (data += chunk));
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve();
        else reject(new Error(`Teams webhook ${res.statusCode}: ${data}`));
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

module.exports = { sendTeamsNotification };
