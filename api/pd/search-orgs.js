import { requireAuth, pdApi, setCors } from '../_utils.js';

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const session = await requireAuth(req, res);
  if (!session) return;

  const { term } = req.query;
  if (!term || term.trim().length < 2) { res.json({ ok: true, items: [] }); return; }

  console.log(`[Search Orgs] term="${term}"`);
  try {
    const result = await pdApi(session.access_token,
      `/organizations/search?term=${encodeURIComponent(term)}&fields=name&limit=10`);
    const items = (result.data?.items || []).map(i => ({
      id:        i.item.id,
      name:      i.item.name,
      address:   i.item.address || '',
      openLeads: i.item.open_deals_count || 0,
    }));
    console.log(`[Search Orgs] Found ${items.length} results`);
    res.json({ ok: true, items });
  } catch (e) {
    console.error('[Search Orgs] Error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
}
