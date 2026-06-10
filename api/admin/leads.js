import { requireAuth, pdApi, setCors } from '../_utils.js';

const ADMIN_EMAIL = process.env.ADMIN_EMAIL || 'roye@contel.co.il';

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const session = await requireAuth(req, res);
  if (!session) return;

  // Admin-only
  if (session.user.email !== ADMIN_EMAIL) {
    console.warn(`[Admin] Unauthorized access attempt by ${session.user.email}`);
    res.status(403).json({ error: 'Access denied. Admin only.' });
    return;
  }

  console.log('[Admin] Fetching leads from Pipedrive...');

  try {
    const { filter_user, search, start_date, end_date, limit = 100 } = req.query;

    // Fetch all leads with full details
    const leadsRes = await pdApi(session.access_token,
      `/leads?limit=${limit}&archived_status=not_archived`);

    if (!leadsRes.success) throw new Error(leadsRes.error || 'Failed to fetch leads');

    let leads = leadsRes.data || [];
    console.log(`[Admin] Fetched ${leads.length} leads`);

    // Fetch person + org details for each lead in parallel (batched)
    const enriched = await Promise.all(leads.map(async lead => {
      const [orgData, personData] = await Promise.all([
        lead.organization_id?.value
          ? pdApi(session.access_token, `/organizations/${lead.organization_id.value}`)
          : Promise.resolve(null),
        lead.person_id?.value
          ? pdApi(session.access_token, `/persons/${lead.person_id.value}`)
          : Promise.resolve(null),
      ]);

      // Fetch latest note for this lead
      const notesRes = await pdApi(session.access_token,
        `/leads/${lead.id}/notes?limit=1&sort=add_time+DESC`).catch(() => null);
      const latestNote = notesRes?.data?.[0]?.content || '';

      // Parse BANT scores from note
      const fitMatch   = latestNote.match(/Fit:\s*(\d+)\/60/);
      const bantMatch  = latestNote.match(/BANT:\s*(\d+)\/40/);
      const totalMatch = latestNote.match(/סה"כ:\s*(\d+)\/100/);
      const tierMatch  = latestNote.match(/HOT|WARM|COLD/);
      const sdrMatch   = latestNote.match(/SDR:\s*(.+?)\s*</);

      return {
        id:        lead.id,
        title:     lead.title,
        created:   lead.add_time,
        updated:   lead.update_time,
        owner:     lead.owner_id,
        orgId:     lead.organization_id?.value || null,
        orgName:   orgData?.data?.name || '—',
        personId:  lead.person_id?.value || null,
        personName:personData?.data?.name || '—',
        personEmail:personData?.data?.email?.[0]?.value || '—',
        fitScore:  fitMatch   ? parseInt(fitMatch[1])   : null,
        bantScore: bantMatch  ? parseInt(bantMatch[1])  : null,
        total:     totalMatch ? parseInt(totalMatch[1]) : null,
        tier:      tierMatch  ? tierMatch[0] : '—',
        sdrName:   sdrMatch   ? sdrMatch[1].trim() : '—',
        pdUrl:     `https://app.pipedrive.com/leads/inbox/${lead.id}`,
        note:      latestNote.slice(0,200),
      };
    }));

    // Filters
    let filtered = enriched;
    if (search) {
      const q = search.toLowerCase();
      filtered = filtered.filter(l =>
        l.title?.toLowerCase().includes(q) ||
        l.orgName?.toLowerCase().includes(q) ||
        l.personName?.toLowerCase().includes(q)
      );
    }
    if (start_date) filtered = filtered.filter(l => l.created >= start_date);
    if (end_date)   filtered = filtered.filter(l => l.created <= end_date + 'T23:59:59');
    if (filter_user) filtered = filtered.filter(l => l.sdrName?.toLowerCase().includes(filter_user.toLowerCase()));

    filtered.sort((a,b) => new Date(b.created) - new Date(a.created));

    console.log(`[Admin] Returning ${filtered.length} leads after filters`);
    res.json({ ok: true, leads: filtered, total: filtered.length });

  } catch (e) {
    console.error('[Admin] Error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
}
