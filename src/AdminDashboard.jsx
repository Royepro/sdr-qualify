
import { useState, useEffect } from "react";

const S = {
  input:{width:"100%",background:"#0a0f1a",border:"1px solid #1e293b",borderRadius:8,
    padding:"8px 12px",color:"#e2e8f0",fontSize:13,outline:"none",boxSizing:"border-box"},
  btn:{padding:"7px 14px",borderRadius:8,cursor:"pointer",border:"1px solid transparent",
    fontSize:12,fontFamily:"inherit",transition:"all .15s"},
};
const TIER_COLOR = {HOT:"#ef4444",WARM:"#f59e0b",COLD:"#60a5fa"};
const fmt = d => d ? new Date(d).toLocaleString('he-IL',{day:'2-digit',month:'2-digit',
  year:'2-digit',hour:'2-digit',minute:'2-digit'}) : '—';

export default function AdminDashboard({ user }) {
  const [leads, setLeads]       = useState([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState('');
  const [search, setSearch]     = useState('');
  const [filterUser, setFilterUser] = useState('');
  const [startDate, setStartDate]   = useState('');
  const [endDate, setEndDate]       = useState('');
  const [selected, setSelected]     = useState(null);

  const fetchLeads = async () => {
    setLoading(true); setError('');
    try {
      const params = new URLSearchParams();
      if (search)     params.set('search', search);
      if (filterUser) params.set('filter_user', filterUser);
      if (startDate)  params.set('start_date', startDate);
      if (endDate)    params.set('end_date', endDate);
      const res = await fetch('/api/admin/leads?' + params);
      const data = await res.json();
      if (!data.ok) throw new Error(data.error);
      setLeads(data.leads || []);
    } catch (e) { setError(e.message); }
    setLoading(false);
  };

  useEffect(() => { fetchLeads(); }, []);

  const sdrUsers = [...new Set(leads.map(l=>l.sdrName).filter(Boolean))];

  return (
    <div style={{minHeight:'100vh',background:'#070d14',color:'#e2e8f0',
      fontFamily:"'Segoe UI',Arial,sans-serif",direction:'rtl'}}>

      {/* Header */}
      <div style={{background:'#0a0f1a',borderBottom:'1px solid #1e293b',
        padding:'12px 24px',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div>
          <div style={{fontSize:18,fontWeight:900}}>
            SDR<span style={{color:'#f59e0b'}}>.</span>qualify
            <span style={{fontSize:12,color:'#475569',marginRight:10,fontWeight:400}}>Admin Dashboard</span>
          </div>
          <div style={{fontSize:11,color:'#475569',marginTop:2}}>{user?.name} · {user?.email}</div>
        </div>
        <div style={{display:'flex',gap:10,alignItems:'center'}}>
          <span style={{fontSize:13,color:'#f59e0b',fontWeight:700}}>
            {leads.length} לידים
          </span>
          <a href="/" style={{...S.btn,background:'#1e293b',color:'#94a3b8',textDecoration:'none'}}>
            ← חזור לאפליקציה
          </a>
          <a href="/api/auth/logout" style={{...S.btn,background:'#7f1d1d22',color:'#fca5a5',textDecoration:'none'}}>
            התנתק
          </a>
        </div>
      </div>

      {/* Filters */}
      <div style={{padding:'14px 24px',background:'#0a0f1a',borderBottom:'1px solid #1e293b',
        display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="🔍 חיפוש: חברה / ליד / איש קשר"
          style={{...S.input,maxWidth:280}}/>
        <select value={filterUser} onChange={e=>setFilterUser(e.target.value)} style={{...S.input,width:'auto'}}>
          <option value="">כל ה-SDR</option>
          {sdrUsers.map(u=><option key={u} value={u}>{u}</option>)}
        </select>
        <input type="date" value={startDate} onChange={e=>setStartDate(e.target.value)}
          style={{...S.input,width:'auto'}}/>
        <span style={{color:'#475569',fontSize:12}}>עד</span>
        <input type="date" value={endDate} onChange={e=>setEndDate(e.target.value)}
          style={{...S.input,width:'auto'}}/>
        <button onClick={fetchLeads} style={{...S.btn,background:'#f59e0b',color:'#000',fontWeight:700}}>
          {loading ? '⏳' : '🔄'} רענן
        </button>
        <button onClick={()=>{setSearch('');setFilterUser('');setStartDate('');setEndDate('');}}
          style={{...S.btn,background:'#1e293b',color:'#94a3b8'}}>נקה</button>
      </div>

      {/* Error */}
      {error && (
        <div style={{margin:'16px 24px',padding:'12px 16px',background:'#7f1d1d22',
          border:'1px solid #7f1d1d',borderRadius:10,color:'#fca5a5',fontSize:13}}>
          ❌ {error}
        </div>
      )}

      {/* Table */}
      <div style={{overflowX:'auto',padding:'0 24px 24px'}}>
        {loading ? (
          <div style={{textAlign:'center',padding:'60px 0',color:'#475569'}}>⏳ טוען נתונים מ-Pipedrive...</div>
        ) : leads.length === 0 ? (
          <div style={{textAlign:'center',padding:'60px 0',color:'#334155'}}>
            <div style={{fontSize:40,marginBottom:12}}>📋</div>
            אין לידים תואמים
          </div>
        ) : (
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:12,marginTop:16}}>
            <thead>
              <tr style={{background:'#0a0f1a',borderBottom:'2px solid #1e293b'}}>
                {['תאריך','SDR','שם ליד','חברה','איש קשר','Fit','BANT','סה״כ','סיווג','ID Org','ID Person','ID Lead','פתח'].map(h=>(
                  <th key={h} style={{padding:'10px 12px',textAlign:'right',color:'#64748b',fontWeight:600,whiteSpace:'nowrap'}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {leads.map((l,i)=>(
                <tr key={l.id} onClick={()=>setSelected(selected?.id===l.id?null:l)}
                  style={{borderBottom:'1px solid #0f1923',cursor:'pointer',
                    background:selected?.id===l.id?'#f59e0b08':i%2===0?'#0d1520':'#0a0f1a',
                    transition:'background .1s'}}>
                  <td style={{padding:'9px 12px',color:'#64748b',whiteSpace:'nowrap'}}>{fmt(l.created)}</td>
                  <td style={{padding:'9px 12px',color:'#94a3b8'}}>{l.sdrName}</td>
                  <td style={{padding:'9px 12px',color:'#e2e8f0',fontWeight:600,maxWidth:180,
                    overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.title}</td>
                  <td style={{padding:'9px 12px',color:'#94a3b8'}}>{l.orgName}</td>
                  <td style={{padding:'9px 12px',color:'#64748b'}}>{l.personName}</td>
                  <td style={{padding:'9px 12px',textAlign:'center',color:'#60a5fa'}}>{l.fitScore??'—'}</td>
                  <td style={{padding:'9px 12px',textAlign:'center',color:'#f59e0b'}}>{l.bantScore??'—'}</td>
                  <td style={{padding:'9px 12px',textAlign:'center',fontWeight:700,
                    color:l.total>=67?'#ef4444':l.total>=36?'#f59e0b':'#60a5fa'}}>{l.total??'—'}</td>
                  <td style={{padding:'9px 12px',textAlign:'center'}}>
                    {l.tier!=='—'&&<span style={{padding:'2px 8px',borderRadius:8,fontSize:11,fontWeight:700,
                      background:`${TIER_COLOR[l.tier]||'#334155'}20`,color:TIER_COLOR[l.tier]||'#94a3b8'}}>
                      {l.tier==='HOT'?'🔥':l.tier==='WARM'?'⚡':'❄️'} {l.tier}
                    </span>}
                  </td>
                  <td style={{padding:'9px 12px',color:'#334155',fontFamily:'monospace'}}>{l.orgId||'—'}</td>
                  <td style={{padding:'9px 12px',color:'#334155',fontFamily:'monospace'}}>{l.personId||'—'}</td>
                  <td style={{padding:'9px 12px',color:'#334155',fontFamily:'monospace',maxWidth:120,
                    overflow:'hidden',textOverflow:'ellipsis'}}>{l.id}</td>
                  <td style={{padding:'9px 12px'}}>
                    <a href={l.pdUrl} target="_blank" rel="noreferrer"
                      style={{color:'#f59e0b',fontSize:12,textDecoration:'none'}}>↗ PD</a>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Detail panel */}
      {selected && (
        <div style={{position:'fixed',bottom:0,right:0,left:0,background:'#0d1520',
          border:'1px solid #1e293b',borderRadius:'12px 12px 0 0',padding:'16px 24px',
          boxShadow:'0 -8px 32px #000a',maxHeight:'40vh',overflowY:'auto'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
            <div style={{fontSize:15,fontWeight:700,color:'#e2e8f0'}}>{selected.title}</div>
            <button onClick={()=>setSelected(null)}
              style={{background:'none',border:'none',color:'#475569',fontSize:20,cursor:'pointer'}}>×</button>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(200px,1fr))',gap:10,marginBottom:12}}>
            {[
              {l:'חברה',v:selected.orgName},{l:'איש קשר',v:selected.personName},
              {l:'אימייל',v:selected.personEmail},{l:'SDR',v:selected.sdrName},
              {l:'נוצר',v:fmt(selected.created)},{l:'עודכן',v:fmt(selected.updated)},
              {l:'Org ID',v:selected.orgId},{l:'Person ID',v:selected.personId},
            ].map(x=>(
              <div key={x.l} style={{background:'#070d14',padding:'8px 12px',borderRadius:8,border:'1px solid #1e293b'}}>
                <div style={{fontSize:10,color:'#475569',marginBottom:2}}>{x.l}</div>
                <div style={{fontSize:13,color:'#e2e8f0'}}>{x.v||'—'}</div>
              </div>
            ))}
          </div>
          {selected.note && (
            <div style={{background:'#070d14',padding:'10px 12px',borderRadius:8,border:'1px solid #1e293b',
              fontSize:11,color:'#64748b',whiteSpace:'pre-wrap',lineHeight:1.6}}>
              {selected.note}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
