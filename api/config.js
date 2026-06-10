export default function handler(req, res) {
  res.json({
    oauthEnabled: !!(process.env.PIPEDRIVE_CLIENT_ID && process.env.PIPEDRIVE_CLIENT_SECRET),
    fallbackToken: !!(process.env.PD_TOKEN),
    adminEmail: process.env.ADMIN_EMAIL || 'roye@contel.co.il',
  });
}
