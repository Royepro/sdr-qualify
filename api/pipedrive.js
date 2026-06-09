// Vercel Serverless Function — Pipedrive Proxy
// Handles all PD API calls server-side so the token never leaks to the browser

export default async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-pd-token');

  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  // Get token: prefer env var (set in Vercel dashboard), fallback to header
  const token = process.env.PD_TOKEN || req.headers['x-pd-token'];

  if (!token) {
    console.error('[PD Proxy] No token — set PD_TOKEN in Vercel Environment Variables');
    res.status(401).json({ success: false, error: 'PD_TOKEN not configured on server' });
    return;
  }

  // Build Pipedrive URL
  const pdPath   = (req.query.pdpath || '').replace(/^\//, '');
  const { pdpath, ...rest } = req.query;
  const qs       = new URLSearchParams({ ...rest, api_token: token }).toString();
  const url      = `https://api.pipedrive.com/v1/${pdPath}?${qs}`;

  console.log(`[PD Proxy] ${req.method} /${pdPath}`);

  try {
    const pdRes = await fetch(url, {
      method:  req.method,
      headers: { 'Content-Type': 'application/json' },
      body:    ['POST','PUT','PATCH'].includes(req.method)
                 ? JSON.stringify(req.body)
                 : undefined,
    });

    const data = await pdRes.json();

    if (!pdRes.ok || !data.success) {
      console.error('[PD Proxy] Pipedrive error:', data.error || data);
    } else {
      console.log(`[PD Proxy] OK — ${pdPath} → id:`, data.data?.id ?? '—');
    }

    res.status(pdRes.status).json(data);

  } catch (err) {
    console.error('[PD Proxy] Fetch error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
}
