// Vercel Serverless Function — Pipedrive Proxy
// Solves CORS. Token stays on server, never exposed to browser.
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

  if (req.method === 'OPTIONS') { res.status(200).end(); return }

  const token = req.headers['x-pd-token'] || process.env.PD_TOKEN
  if (!token) { res.status(401).json({ success: false, error: 'Missing token' }); return }

  const { path = '', ...queryParams } = req.query
  const qs = new URLSearchParams({ ...queryParams, api_token: token }).toString()
  const url = `https://api.pipedrive.com/v1/${path}?${qs}`

  try {
    const response = await fetch(url, {
      method: req.method,
      headers: { 'Content-Type': 'application/json' },
      body: ['POST','PUT','PATCH'].includes(req.method) ? JSON.stringify(req.body) : undefined
    })
    const data = await response.json()
    res.status(response.status).json(data)
  } catch (err) {
    res.status(500).json({ success: false, error: err.message })
  }
}
