import { requireAuth, setCors } from '../_utils.js';

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const session = await requireAuth(req, res);
  if (!session) return;

  res.json({
    ok: true,
    user: session.user,
    expires_at: session.expires_at,
  });
}
