import { parseCookies, setSessionCookie, pdApi } from '../_utils.js';

export default async function handler(req, res) {
  const { code, state, error } = req.query;

  if (error) {
    console.error('[OAuth] Callback error:', error);
    res.redirect('/?auth_error=' + encodeURIComponent(error));
    return;
  }

  // CSRF check
  const cookies = parseCookies(req.headers.cookie);
  if (!state || state !== cookies.oauth_state) {
    console.error('[OAuth] State mismatch — possible CSRF');
    res.redirect('/?auth_error=state_mismatch');
    return;
  }

  const { PIPEDRIVE_CLIENT_ID, PIPEDRIVE_CLIENT_SECRET, PIPEDRIVE_REDIRECT_URI } = process.env;

  try {
    console.log('[OAuth] Exchanging code for tokens...');
    const basic = Buffer.from(`${PIPEDRIVE_CLIENT_ID}:${PIPEDRIVE_CLIENT_SECRET}`).toString('base64');

    const tokenRes = await fetch('https://oauth.pipedrive.com/oauth/token', {
      method: 'POST',
      headers: {
        'Authorization': `Basic ${basic}`,
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        grant_type:   'authorization_code',
        code,
        redirect_uri: PIPEDRIVE_REDIRECT_URI,
      }),
    });

    const tokens = await tokenRes.json();
    if (!tokenRes.ok || !tokens.access_token) {
      console.error('[OAuth] Token exchange failed:', tokens);
      res.redirect('/?auth_error=token_exchange_failed');
      return;
    }

    console.log('[OAuth] Tokens received. Fetching user info...');
    const meRes = await pdApi(tokens.access_token, '/users/me');
    const user  = meRes.data;

    const session = {
      access_token:  tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_at:    Date.now() + (tokens.expires_in || 3600) * 1000,
      user: {
        id:        user.id,
        name:      user.name,
        email:     user.email,
        pic:       user.icon_url || '',
        role:      user.role_id,
      },
    };

    setSessionCookie(res, session);
    // Clear CSRF state cookie
    res.setHeader('Set-Cookie', [
      res.getHeader('Set-Cookie'),
      'oauth_state=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0',
    ].flat());

    console.log(`[OAuth] Login successful: ${user.name} <${user.email}> (id:${user.id})`);
    res.redirect('/');

  } catch (e) {
    console.error('[OAuth] Callback exception:', e.message);
    res.redirect('/?auth_error=' + encodeURIComponent(e.message));
  }
}
