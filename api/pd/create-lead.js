import { requireAuth, pdApi, setCors } from '../_utils.js';

export default async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST') { res.status(405).json({ error: 'Method not allowed' }); return; }

  const session = await requireAuth(req, res);
  if (!session) return;

  const {
    // Lead info
    leadName, company, contact, role, phone, email, notes, label,
    // Scores
    scores, fitScore, bantScore, totalScore, tier,
    // Existing PD IDs (if user selected from duplicate check)
    existingOrgId, existingPersonId,
    // Tasks (stored locally, not pushed to PD in this version)
  } = req.body;

  const log = [];
  const result = { orgId: null, personId: null, leadId: null, noteId: null };

  try {
    // ── 1. Organization ──────────────────────────────────────────────
    if (existingOrgId) {
      result.orgId = existingOrgId;
      console.log(`[Create Lead] Using existing org id:${existingOrgId}`);
      log.push(`✓ Org: existing (id:${existingOrgId})`);
    } else if (company) {
      console.log(`[Create Lead] Creating org: "${company}"`);
      const orgRes = await pdApi(session.access_token, '/organizations', 'POST', {
        name:     company,
        owner_id: session.user.id,
      });
      if (!orgRes.success) throw new Error('Create org failed: ' + (orgRes.error || JSON.stringify(orgRes)));
      result.orgId = orgRes.data.id;
      console.log(`[Create Lead] Org created, id:${result.orgId}`);
      log.push(`✓ Org created: "${company}" (id:${result.orgId})`);
    }

    // ── 2. Person ────────────────────────────────────────────────────
    if (existingPersonId) {
      result.personId = existingPersonId;
      console.log(`[Create Lead] Using existing person id:${existingPersonId}`);
      log.push(`✓ Person: existing (id:${existingPersonId})`);
      // Link to org if needed
      if (result.orgId) {
        await pdApi(session.access_token, `/persons/${existingPersonId}`, 'PUT', { org_id: result.orgId });
      }
    } else if (contact) {
      console.log(`[Create Lead] Creating person: "${contact}"`);
      const personBody = {
        name:     contact,
        owner_id: session.user.id,
      };
      if (result.orgId) personBody.org_id = result.orgId;
      if (phone)  personBody.phone  = [{ value: phone,  primary: true }];
      if (email)  personBody.email  = [{ value: email,  primary: true }];

      const personRes = await pdApi(session.access_token, '/persons', 'POST', personBody);
      if (!personRes.success) throw new Error('Create person failed: ' + (personRes.error || JSON.stringify(personRes)));
      result.personId = personRes.data.id;
      console.log(`[Create Lead] Person created, id:${result.personId}`);
      log.push(`✓ Person created: "${contact}" (id:${result.personId})`);
    }

    // ── 3. Lead ──────────────────────────────────────────────────────
    const title = leadName || `MES - ${company}`;
    console.log(`[Create Lead] Creating lead: "${title}"`);
    const leadBody = {
      title,
      owner_id: session.user.id,
    };
    if (result.orgId)    leadBody.organization_id = { value: result.orgId };
    if (result.personId) leadBody.person_id        = { value: result.personId };

    const leadRes = await pdApi(session.access_token, '/leads', 'POST', leadBody);
    if (!leadRes.success) throw new Error('Create lead failed: ' + (leadRes.error || JSON.stringify(leadRes)));
    result.leadId = leadRes.data.id;
    console.log(`[Create Lead] Lead created, id:${result.leadId}`);
    log.push(`✓ Lead created: "${title}" (id:${result.leadId})`);

    // ── 4. Note ──────────────────────────────────────────────────────
    const BANT = scores || {};
    const CRITERIA_LABELS = {
      budget:'Budget',authority:'Authority',need:'Need',timing:'Timing',
      industry:'תעשייה',lines:'קווי ייצור',complexity:'מורכבות',sites:'אתרים',employees:'עובדים',valuation:'שווי'
    };
    const bantDetail = ['budget','authority','need','timing']
      .map(k => `${CRITERIA_LABELS[k]}: ${BANT[k]||0}/9`).join(' | ');
    const fitDetail = ['industry','lines','complexity','sites','employees','valuation']
      .map(k => `${CRITERIA_LABELS[k]}: ${BANT[k]||0}/9`).join(' | ');

    const noteContent = [
      `🎯 SDR Qualification — ${title}`,
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,
      `📊 ציון: ${tier?.emoji || ''} ${tier?.label || ''} | Fit: ${fitScore||0}/60 | BANT: ${bantScore||0}/40 | סה"כ: ${totalScore||0}/100`,
      ``,
      `🏭 Fit:  ${fitDetail}`,
      `💰 BANT: ${bantDetail}`,
      ``,
      `📝 הערות SDR:`,
      notes || '—',
      ``,
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,
      `👤 SDR: ${session.user.name} <${session.user.email}>`,
      `📅 ${new Date().toLocaleString('he-IL', { timeZone: 'Asia/Jerusalem' })}`,
      `🔖 חברה: ${company} | איש קשר: ${contact||'—'} | תפקיד: ${role||'—'}`,
      `App: SDR.qualify v3`,
    ].join('\n');

    const noteBody = { content: noteContent };
    if (result.leadId)   noteBody.lead_id        = result.leadId;
    if (result.orgId)    noteBody.org_id          = result.orgId;
    if (result.personId) noteBody.person_id       = result.personId;

    const noteRes = await pdApi(session.access_token, '/notes', 'POST', noteBody);
    if (noteRes.success) {
      result.noteId = noteRes.data.id;
      console.log(`[Create Lead] Note created, id:${result.noteId}`);
      log.push(`✓ Note created (id:${result.noteId})`);
    } else {
      console.warn('[Create Lead] Note failed:', noteRes.error);
      log.push(`⚠ Note failed: ${noteRes.error}`);
    }

    res.json({ ok: true, result, log });

  } catch (e) {
    console.error('[Create Lead] Fatal error:', e.message);
    res.status(500).json({ ok: false, error: e.message, result, log });
  }
}
