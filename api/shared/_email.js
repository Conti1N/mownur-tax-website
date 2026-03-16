// api/shared/_email.js
const { getAccessToken, graphRequest } = require('./_graph');

async function sendEmail({ to, subject, html }) {
  const from = process.env.EMAIL_FROM;
  if (!from) throw new Error('EMAIL_FROM not set');

  const token = await getAccessToken();
  const userId = process.env.ONEDRIVE_USER_ID || from;

  await graphRequest(token, 'POST',
    `/users/${encodeURIComponent(userId)}/sendMail`,
    {
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: to } }],
      },
      saveToSentItems: true,
    }
  );
}

module.exports = { sendEmail };
