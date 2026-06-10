import { generateState, signToken } from '../_utils.js';

export default function handler(req, res) {
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const { PIPEDRIVE_CLIENT_ID, PIPEDRIVE_REDIRECT_URI } = process.env;
  if (!PIPEDRIVE_CLIENT_ID) {
    res.status(500).json({ error: 'PIPEDRIVE_CLIENT_ID not configured' });
    return;
  }

  const state = generateState();
  // Store state in short-lived cookie for CSRF protection
  res.setHeader('Set-Cookie', `oauth_state=${state}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`);

  const scopes = [
    'leads:read','leads:write',
    'organizations:read','organizations:write',
    'persons:read','persons:write',
    'notes:read','notes:write',
    'users:read',
  ].join(' ');

  const params = new URLSearchParams({
    client_id:     PIPEDRIVE_CLIENT_ID,
    redirect_uri:  PIPEDRIVE_REDIRECT_URI,
    response_type: 'code',
    state,
    scope: scopes,
  });

  console.log('[OAuth] Login redirect started');
  res.redirect(`https://oauth.pipedrive.com/oauth/authorize?${params}`);
}
