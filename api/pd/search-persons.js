import { requireAuth, pdApi, setCors } from '../_utils.js';

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const session = await requireAuth(req, res);
  if (!session) return;

  const { term } = req.query;
  if (!term || term.trim().length < 2) { res.json({ ok: true, items: [] }); return; }

  console.log(`[Search Persons] term="${term}"`);
  try {
    const result = await pdApi(session.access_token,
      `/persons/search?term=${encodeURIComponent(term)}&fields=name,email,phone&limit=10`);
    const items = (result.data?.items || []).map(i => ({
      id:    i.item.id,
      name:  i.item.name,
      email: i.item.emails?.[0]?.value || '',
      phone: i.item.phones?.[0]?.value || '',
      org:   i.item.organization?.name || '',
      orgId: i.item.organization?.id   || null,
    }));
    console.log(`[Search Persons] Found ${items.length} results`);
    res.json({ ok: true, items });
  } catch (e) {
    console.error('[Search Persons] Error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
}
