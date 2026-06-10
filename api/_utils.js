// Shared utilities for all API routes
import { createHmac, randomBytes } from 'crypto';

/* ── Cookie helpers ── */
export function parseCookies(header = '') {
  return Object.fromEntries(
    (header || '').split(';')
      .map(c => c.trim()).filter(Boolean)
      .map(c => { const i = c.indexOf('='); return [c.slice(0,i).trim(), decodeURIComponent(c.slice(i+1).trim())]; })
  );
}

/* ── Simple JWT (HMAC-SHA256) without npm deps ── */
function b64url(s) { return Buffer.from(s).toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,''); }
function b64decode(s) { return Buffer.from(s.replace(/-/g,'+').replace(/_/g,'/'), 'base64').toString(); }

export function signToken(payload) {
  const secret = process.env.SESSION_SECRET || 'dev-secret-change-me';
  const header  = b64url(JSON.stringify({ alg:'HS256', typ:'JWT' }));
  const body    = b64url(JSON.stringify({ ...payload, iat: Math.floor(Date.now()/1000) }));
  const sig     = createHmac('sha256', secret).update(`${header}.${body}`).digest('base64url');
  return `${header}.${body}.${sig}`;
}

export function verifyToken(token) {
  if (!token) return null;
  const secret = process.env.SESSION_SECRET || 'dev-secret-change-me';
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const sig = createHmac('sha256', secret).update(`${parts[0]}.${parts[1]}`).digest('base64url');
  if (sig !== parts[2]) return null;
  try {
    const payload = JSON.parse(b64decode(parts[1]));
    // 7-day expiry
    if (payload.iat && Date.now()/1000 - payload.iat > 7*86400) return null;
    return payload;
  } catch { return null; }
}

/* ── Session ── */
export function getSession(req) {
  const cookies = parseCookies(req.headers.cookie);
  return verifyToken(cookies.pd_session);
}

export function setSessionCookie(res, data) {
  const token = signToken(data);
  const flags = 'HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=604800';
  res.setHeader('Set-Cookie', `pd_session=${token}; ${flags}`);
}

export function clearSession(res) {
  res.setHeader('Set-Cookie', 'pd_session=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0');
}

/* ── Pipedrive API helper ── */
export async function pdApi(accessToken, path, method = 'GET', body = null) {
  const url = `https://api.pipedrive.com/v1${path}`;
  console.log(`[PD] ${method} ${path}`);
  const res = await fetch(url, {
    method,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const json = await res.json();
  if (!res.ok) console.error(`[PD] Error ${res.status}:`, json.error || json);
  return json;
}

/* ── Token refresh ── */
export async function refreshAccessToken(refreshToken) {
  const { PIPEDRIVE_CLIENT_ID, PIPEDRIVE_CLIENT_SECRET } = process.env;
  const basic = Buffer.from(`${PIPEDRIVE_CLIENT_ID}:${PIPEDRIVE_CLIENT_SECRET}`).toString('base64');
  const res = await fetch('https://oauth.pipedrive.com/oauth/token', {
    method: 'POST',
    headers: {
      'Authorization': `Basic ${basic}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({ grant_type: 'refresh_token', refresh_token: refreshToken }),
  });
  if (!res.ok) throw new Error('Token refresh failed');
  return res.json();
}

/* ── Auth middleware ── */
export async function requireAuth(req, res) {
  let session = getSession(req);
  if (!session) { res.status(401).json({ error: 'Not authenticated. Please log in.' }); return null; }

  // Refresh token if access_token close to expiry (within 5 min)
  if (session.expires_at && Date.now() > (session.expires_at - 5*60*1000)) {
    try {
      console.log('[Auth] Refreshing access token...');
      const refreshed = await refreshAccessToken(session.refresh_token);
      session = {
        ...session,
        access_token:  refreshed.access_token,
        refresh_token: refreshed.refresh_token || session.refresh_token,
        expires_at:    Date.now() + refreshed.expires_in * 1000,
      };
      setSessionCookie(res, session);
      console.log('[Auth] Token refreshed ✓');
    } catch (e) {
      console.error('[Auth] Refresh failed:', e.message);
      clearSession(res);
      res.status(401).json({ error: 'Session expired. Please log in again.' });
      return null;
    }
  }
  return session;
}

/* ── CORS ── */
export function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', process.env.VERCEL_URL ? `https://${process.env.VERCEL_URL}` : '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
}

export function generateState() { return randomBytes(16).toString('hex'); }
