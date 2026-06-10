import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import LoginPage from "./LoginPage";
import AdminDashboard from "./AdminDashboard";

/* ═══ CONSTANTS ═══ */
const CRITERIA = [
  { id:"industry",   group:"fit",  label:"סוג התעשייה",    low:"מיכלים/נגריות/הפצה",  mid:"CNC/אלקטרוניקה/יקבים", high:"מזון/משקאות/פלסטיקה/פארמה" },
  { id:"lines",      group:"fit",  label:"קווי ייצור",      low:"ידניים/מיושנים",       mid:"שילוב/חצי-אוטומטי",     high:"אוטומטיים מתקדמים" },
  { id:"complexity", group:"fit",  label:"מורכבות הייצור",  low:"מפעל קטן",             mid:"בינוני, סטנדרטי",        high:"מורכב, רב-שלבי" },
  { id:"sites",      group:"fit",  label:"אתרי ייצור",      low:"1 אתר",                mid:"1–2 אתרים",              high:"3+ אתרים" },
  { id:"employees",  group:"fit",  label:"עובדים",          low:"1–40",                 mid:"40–80",                  high:"80+" },
  { id:"valuation",  group:"fit",  label:"שווי / מחזור",    low:"עד 10M ₪",             mid:"10–100M ₪",              high:"100M+ ₪" },
  { id:"budget",     group:"bant", label:"Budget — תקציב",  low:"אין תקציב",            mid:"תקציב כללי",             high:"תקציב מאושר" },
  { id:"authority",  group:"bant", label:"Authority",       low:"משתמש קצה",            mid:"מנהל ביניים",            high:"C-Level" },
  { id:"need",       group:"bant", label:"Need — צורך",     low:"עניין כללי",           mid:"בעיה, Excel",            high:"כאב אקוטי" },
  { id:"timing",     group:"bant", label:"Timing — זמן",    low:"שנה הבאה",             mid:"תוך 6 חודשים",           high:"מיידי" },
];
const VALS = { low:2, mid:6, high:9 };
const TASK_TYPES = [
  {key:"call",label:"שיחה",icon:"📞",color:"#f59e0b"},
  {key:"email",label:"אימייל",icon:"📧",color:"#60a5fa"},
  {key:"meeting",label:"פגישה",icon:"🤝",color:"#8b5cf6"},
  {key:"demo",label:"דמו",icon:"💻",color:"#06b6d4"},
  {key:"follow",label:"Follow Up",icon:"🔁",color:"#22c55e"},
  {key:"proposal",label:"הצעת מחיר",icon:"📄",color:"#ec4899"},
  {key:"other",label:"אחר",icon:"📌",color:"#94a3b8"},
];
const HOURS = Array.from({length:14},(_,i)=>i+7);
const DAY_SH = ["א׳","ב׳","ג׳","ד׳","ה׳","ו׳","ש׳"];
const MON_HE = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני","יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"];

/* ═══ HELPERS ═══ */
const uid       = () => Date.now().toString(36) + Math.random().toString(36).slice(2,5);
const calcScore = (s={}) => CRITERIA.reduce((t,c)=>t+(s[c.id]||0),0);
const fitScore  = (s={}) => CRITERIA.filter(c=>c.group==="fit").reduce((t,c)=>t+(s[c.id]||0),0);
const bantScore = (s={}) => CRITERIA.filter(c=>c.group==="bant").reduce((t,c)=>t+(s[c.id]||0),0);
const getTier   = n => n>=67?{key:"hot",label:"HOT",emoji:"🔥",accent:"#ef4444"}:n>=36?{key:"warm",label:"WARM",emoji:"⚡",accent:"#f59e0b"}:{key:"cold",label:"COLD",emoji:"❄️",accent:"#60a5fa"};
const fmtDate   = d => !d?"—":new Date(d).toLocaleDateString("he-IL",{day:"2-digit",month:"2-digit"});
const isoDate   = d => d?new Date(d).toISOString().slice(0,10):"";
const isOverdue = d => d && new Date(d)<new Date(new Date().toDateString());
const taskType  = t => TASK_TYPES.find(x=>x.key===t)||TASK_TYPES[6];

/* ═══ SERVER API HELPER ═══ */
async function apiPost(path, body) {
  const res = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
    credentials: 'include',
  });
  return res.json();
}
async function apiGet(path) {
  const res = await fetch(path, { credentials: 'include' });
  return res.json();
}

/* ═══ TOAST ═══ */
function useToast() {
  const [toasts, setToasts] = useState([]);
  const add = (msg, type='info') => {
    const id = uid();
    setToasts(t=>[...t,{id,msg,type}]);
    setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)), 5000);
  };
  return { toasts, success:m=>add(m,'success'), error:m=>add(m,'error'), info:m=>add(m,'info') };
}
function ToastContainer({ toasts }) {
  return (
    <div style={{position:'fixed',bottom:20,left:'50%',transform:'translateX(-50%)',
      zIndex:9999,display:'flex',flexDirection:'column',gap:8,alignItems:'center',pointerEvents:'none'}}>
      {toasts.map(t=>(
        <div key={t.id} style={{padding:'10px 22px',borderRadius:10,fontSize:13,fontWeight:700,
          background:t.type==='success'?'#22c55e':t.type==='error'?'#ef4444':'#f59e0b',
          color:'#000',boxShadow:'0 4px 20px #0006',animation:'toastIn .2s ease',
          whiteSpace:'nowrap',maxWidth:'90vw',textAlign:'center'}}>
          {t.type==='success'?'✅ ':t.type==='error'?'❌ ':'ℹ️ '}{t.msg}
        </div>
      ))}
      <style>{`@keyframes toastIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}`}</style>
    </div>
  );
}

/* ═══ RESPONSIVE ═══ */
function useBreakpoint() {
  const [w, setW] = useState(window.innerWidth);
  useEffect(()=>{const fn=()=>setW(window.innerWidth);window.addEventListener("resize",fn);return()=>window.removeEventListener("resize",fn);},[]);
  return { isMobile:w<768, isTablet:w>=768&&w<1024, isDesktop:w>=1024, width:w };
}

/* ═══ STYLES ═══ */
const S = {
  input:{width:"100%",background:"#0a0f1a",border:"1px solid #1e293b",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,outline:"none",boxSizing:"border-box"},
  btn:{padding:"8px 16px",borderRadius:8,cursor:"pointer",border:"1px solid transparent",fontSize:13,transition:"all .15s",fontFamily:"inherit"},
  card:{background:"#0d1520",border:"1px solid #1e293b",borderRadius:12,padding:16},
};
const CSS = `
* { box-sizing:border-box; } body { margin:0; }
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:#070d14;}
::-webkit-scrollbar-thumb{background:#1e293b;border-radius:3px;}
input:focus,textarea:focus,select:focus{border-color:#f59e0b!important;outline:none;}
input::placeholder,textarea::placeholder{color:#334155;}
select option{background:#0a0f1a;color:#e2e8f0;}
.card-hover:hover{transform:translateY(-2px);box-shadow:0 6px 24px #0007;}
.ev:hover{filter:brightness(1.2);} .slot-drop{background:#f59e0b10!important;}
.task-drag{cursor:grab;} .task-drag:active{cursor:grabbing;}
@keyframes toastIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
@keyframes slideUp{from{transform:translateY(100%);opacity:0}to{transform:translateY(0);opacity:1}}
@media(max-width:767px){.hide-mobile{display:none!important;}.mobile-full{width:100%!important;max-width:100%!important;border-radius:0!important;height:100%!important;max-height:100%!important;}.cal-sidebar{display:none!important;}}
`;

/* ═══════════════════════════════════════════════
   DUPLICATE CHECK COMPONENTS
═══════════════════════════════════════════════ */
function DuplicateList({ items, type, onSelect, onCreateNew, onCreate }) {
  if (!items) return null;
  return (
    <div style={{position:"absolute",top:"100%",right:0,left:0,zIndex:60,
      background:"#0d1520",border:"1px solid #f59e0b44",borderRadius:8,
      overflow:"hidden",boxShadow:"0 8px 24px #0009",maxHeight:220,overflowY:"auto"}}>
      <div style={{padding:"7px 12px",fontSize:11,color:"#64748b",borderBottom:"1px solid #1e293b",
        background:"#0a0f1a",fontWeight:600}}>
        🔍 נמצא ב-Pipedrive — {type==="org"?"ארגונים":"אנשי קשר"} קיימים
      </div>
      {items.length === 0 && (
        <div style={{padding:"10px 14px",fontSize:12,color:"#334155"}}>לא נמצאו תוצאות</div>
      )}
      {items.map(item=>(
        <div key={item.id}
          onClick={()=>onSelect(item)}
          style={{padding:"10px 14px",cursor:"pointer",borderBottom:"1px solid #0f1923",
            transition:"background .1s"}}
          onMouseOver={e=>e.currentTarget.style.background="#f59e0b0a"}
          onMouseOut={e=>e.currentTarget.style.background=""}>
          <div style={{fontSize:13,color:"#e2e8f0",fontWeight:500}}>
            {type==="org"?"🏢":"👤"} {item.name}
            <span style={{fontSize:11,color:"#475569",marginRight:8}}>ID:{item.id}</span>
          </div>
          {type==="person" && (
            <div style={{fontSize:11,color:"#475569",marginTop:2}}>
              {item.email} {item.org&&`· ${item.org}`}
            </div>
          )}
          {type==="org" && item.openLeads>0 && (
            <div style={{fontSize:10,color:"#f59e0b",marginTop:2}}>{item.openLeads} עסקאות פתוחות</div>
          )}
        </div>
      ))}
      <div onClick={onCreateNew}
        style={{padding:"10px 14px",cursor:"pointer",fontSize:12,color:"#22c55e",
          background:"#22c55e08",borderTop:"1px solid #1e293b",fontWeight:600}}
        onMouseOver={e=>e.currentTarget.style.background="#22c55e12"}
        onMouseOut={e=>e.currentTarget.style.background="#22c55e08"}>
        + צור {type==="org"?"ארגון":"איש קשר"} חדש
      </div>
    </div>
  );
}

function LinkedBadge({ type, item, onClear }) {
  return (
    <div style={{display:"inline-flex",alignItems:"center",gap:6,padding:"4px 10px",
      background:"#22c55e18",border:"1px solid #22c55e44",borderRadius:20,fontSize:12}}>
      <span style={{color:"#22c55e"}}>{type==="org"?"🏢":"👤"} {item.name}</span>
      <span style={{color:"#334155",fontFamily:"monospace",fontSize:10}}>ID:{item.id}</span>
      <span onClick={onClear} style={{color:"#64748b",cursor:"pointer",marginRight:2}}>×</span>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   AI EMAIL MODAL
═══════════════════════════════════════════════ */
function AIEmailModal({ lead, user, onClose }) {
  const [loading, setLoading] = useState(true);
  const [draft, setDraft]   = useState("");
  const [subject, setSubject] = useState("");

  const buildFallback = () => {
    setSubject(`המשך שיחתנו — ${lead.company}`);
    setDraft(`שלום ${lead.contact||""},\n\nתודה על השיחה המעניינת שקיימנו.\n${lead.notes?`\nכפי שדיברנו: ${lead.notes}\n`:""}
אשמח לקבוע שיחת המשך קצרה לצורך:
• הצגת דמו מותאם לתהליכי הייצור שלכם  
• בחינת Quick Win ראשון תוך 4–6 שבועות

מתי נוח לך השבוע?\n\nבברכה,\n${user?.name||"Roy"}\nVP Business Development | Trunovate — PlantSharp MES`);
  };

  useEffect(()=>{
    (async()=>{
      try {
        const score = calcScore(lead.scores||{});
        const tier  = getTier(score);
        const prompt = `אתה מומחה מכירות B2B לתוכנת MES תעשייתי (PlantSharp MES של Trunovate).
כתוב מייל מקצועי בעברית לאיש הקשר הבא לאחר שיחת היכרות ראשונה.
חברה: ${lead.company} | איש קשר: ${lead.contact||"—"} | תפקיד: ${lead.role||"—"}
ציון: ${score}/100 (${tier.label}) | הערות: ${lead.notes||"אין"}
דרישות: פתיחה חמה, סיכום כאב, הצעה לדמו, CTA לקביעת זמן.
פורמט:\nנושא: [נושא]\n---\n[גוף המייל]`;
        const res  = await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:prompt}]})});
        if (!res.ok) throw new Error("API error");
        const data = await res.json();
        const text = (data.content||[]).find(b=>b.type==="text")?.text||"";
        if (!text) throw new Error("empty");
        const lines = text.split("\n");
        const subLine = lines.find(l=>l.startsWith("נושא:"));
        setSubject(subLine?subLine.replace("נושא:","").trim():`המשך שיחתנו — ${lead.company}`);
        const bodyStart = lines.findIndex(l=>l.trim()==="---");
        setDraft(bodyStart>=0?lines.slice(bodyStart+1).join("\n").trim():text);
      } catch { buildFallback(); }
      setLoading(false);
    })();
  },[]);

  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:"fixed",inset:0,background:"#000d",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#0d1520",border:"1px solid #8b5cf644",borderRadius:16,width:"100%",maxWidth:680,maxHeight:"92vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:"14px 18px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:16,fontWeight:700,color:"#e2e8f0"}}>🤖 סוכן AI — טיוטת מייל</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:16}}>
          {loading ? <div style={{textAlign:"center",padding:"40px 0",color:"#8b5cf6",fontSize:13}}>⏳ מנסח...</div> : (
            <>
              <div style={{marginBottom:12}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>נושא</label><input value={subject} onChange={e=>setSubject(e.target.value)} style={S.input}/></div>
              <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>גוף המייל</label><textarea value={draft} onChange={e=>setDraft(e.target.value)} rows={12} style={{...S.input,resize:"vertical",lineHeight:1.7,fontFamily:"inherit"}}/></div>
            </>
          )}
        </div>
        <div style={{padding:"12px 18px",borderTop:"1px solid #1e293b",background:"#0a0f1a",display:"flex",gap:8,flexWrap:"wrap"}}>
          {!loading&&<>
            <a href={`mailto:${lead.email||""}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(draft)}`} style={{...S.btn,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",fontSize:12,textDecoration:"none"}}>📧 פתח בדואר</a>
            <a href={`https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(lead.email||"")}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(draft)}`} target="_blank" rel="noreferrer" style={{...S.btn,background:"#0078d420",color:"#0078d4",border:"1px solid #0078d444",fontSize:12,textDecoration:"none"}}>📆 Outlook</a>
            <button onClick={()=>navigator.clipboard?.writeText(`נושא: ${subject}\n\n${draft}`)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12}}>📋 העתק</button>
          </>}
          <button onClick={onClose} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,marginRight:"auto"}}>סגור</button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   CRITERION ROW
═══════════════════════════════════════════════ */
function CriterionRow({c,value,onChange}){
  return(
    <div style={{marginBottom:14}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
        <span style={{fontSize:13,fontWeight:600,color:"#cbd5e1"}}>{c.label}</span>
        <span style={{fontSize:13,fontFamily:"monospace",fontWeight:700,color:value>=8?"#22c55e":value>=5?"#f59e0b":value>0?"#60a5fa":"#334155"}}>{value||"—"}</span>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
        {[{score:VALS.low,label:c.low,badge:"נמוך",color:"#60a5fa"},{score:VALS.mid,label:c.mid,badge:"בינוני",color:"#f59e0b"},{score:VALS.high,label:c.high,badge:"גבוה",color:"#22c55e"}].map(opt=>(
          <div key={opt.score} onClick={()=>onChange(c.id,value===opt.score?0:opt.score)}
            style={{padding:"7px 10px",borderRadius:8,cursor:"pointer",fontSize:11,lineHeight:1.4,
              border:`1.5px solid ${value===opt.score?opt.color:"#1e293b"}`,
              background:value===opt.score?`${opt.color}18`:"#0a0f1a",
              color:value===opt.score?opt.color:"#64748b",transition:"all .15s"}}>
            <div style={{fontSize:10,fontWeight:700,color:opt.color,marginBottom:2}}>{opt.badge}</div>{opt.label}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   TASK STEP
═══════════════════════════════════════════════ */
function TaskStep({ tasks, onChange }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState({type:"call",title:"",dueDate:"",time:"",notes:""});
  const setF=(k,v)=>setForm(p=>({...p,[k]:v}));
  const add=()=>{ if(!form.title.trim())return; onChange([...tasks,{id:uid(),...form,done:false,createdAt:new Date().toISOString()}]); setForm({type:"call",title:"",dueDate:"",time:"",notes:""}); setAdding(false); };
  return(
    <div>
      <div style={{...S.card,marginBottom:16,fontSize:12,color:"#475569",borderColor:"#1e3a5f"}}>✅ <strong style={{color:"#94a3b8"}}>פעולות המשך</strong></div>
      {!adding&&<button onClick={()=>setAdding(true)} style={{...S.btn,background:"#f59e0b20",color:"#f59e0b",border:"1px solid #f59e0b44",width:"100%",marginBottom:14,textAlign:"center"}}>+ הוסף משימה</button>}
      {adding&&(
        <div style={{...S.card,marginBottom:14,borderColor:"#f59e0b33"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>סוג</label><select value={form.type} onChange={e=>setF("type",e.target.value)} style={S.input}>{TASK_TYPES.map(t=><option key={t.key} value={t.key}>{t.icon} {t.label}</option>)}</select></div>
            <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>תאריך</label><input type="date" value={form.dueDate} onChange={e=>setF("dueDate",e.target.value)} style={S.input}/></div>
            <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>שעה</label><input type="time" value={form.time} onChange={e=>setF("time",e.target.value)} style={S.input}/></div>
          </div>
          <div style={{marginBottom:10}}><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>תיאור *</label><input value={form.title} onChange={e=>setF("title",e.target.value)} placeholder="מה צריך לעשות?" style={S.input}/></div>
          <div style={{marginBottom:10}}><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>הערות</label><textarea value={form.notes} onChange={e=>setF("notes",e.target.value)} rows={2} style={{...S.input,resize:"vertical"}}/></div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <button onClick={()=>setAdding(false)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12}}>ביטול</button>
            <button onClick={add} disabled={!form.title.trim()} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,fontSize:12,opacity:form.title.trim()?1:.4}}>הוסף</button>
          </div>
        </div>
      )}
      {tasks.map(t=>{
        const tt=taskType(t.type); const od=!t.done&&isOverdue(t.dueDate);
        return(<div key={t.id} className="task-drag" style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:8,border:`1px solid ${od?"#7f1d1d44":"#1e293b"}`,marginBottom:6,background:od?"#7f1d1d11":"#0d1520"}}>
          <input type="checkbox" checked={t.done} onChange={()=>onChange(tasks.map(x=>x.id===t.id?{...x,done:!x.done}:x))} style={{accentColor:"#22c55e",cursor:"pointer"}}/>
          <span>{tt.icon}</span>
          <div style={{flex:1}}><div style={{fontSize:13,color:t.done?"#475569":"#e2e8f0",textDecoration:t.done?"line-through":"none"}}>{t.title}</div>{(t.dueDate||t.time)&&<div style={{fontSize:11,color:od?"#ef4444":"#475569"}}>{t.dueDate&&`📅 ${fmtDate(t.dueDate)}`}{t.time&&` ⏰ ${t.time}`}</div>}</div>
          <button onClick={()=>onChange(tasks.filter(x=>x.id!==t.id))} style={{background:"none",border:"none",color:"#334155",cursor:"pointer",fontSize:16}}>×</button>
        </div>);
      })}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   LEAD MODAL — 4 steps with duplicate check
═══════════════════════════════════════════════ */
const STEPS=[{key:"info",label:"פרטים",icon:"📋"},{key:"fit",label:"Fit",icon:"🏭"},{key:"bant",label:"BANT",icon:"💰"},{key:"tasks",label:"משימות",icon:"✅"}];

function LeadModal({ lead:init, onSave, onDelete, onClose, user, toast }) {
  const isNew = !init.id;
  const [lead, setLead] = useState({leadName:"",role:"",email:"",...init,scores:{...init.scores},tasks:[...(init.tasks||[])]});
  const [step, setStep] = useState(0);
  const [confirmDel, setConfirmDel] = useState(false);
  const [showAI, setShowAI] = useState(false);
  const [saving, setSaving] = useState(false);
  const [pdResult, setPdResult] = useState(null);

  // Duplicate check state
  const [orgSearch, setOrgSearch]     = useState([]);
  const [personSearch, setPersonSearch] = useState([]);
  const [selectedOrg, setSelectedOrg]   = useState(init.pipedriveOrg || null);
  const [selectedPerson, setSelectedPerson] = useState(init.pipedrivePerson || null);
  const [orgLoading, setOrgLoading]   = useState(false);
  const [personLoading, setPersonLoading] = useState(false);
  const [showOrgDrop, setShowOrgDrop]   = useState(false);
  const [showPersonDrop, setShowPersonDrop] = useState(false);
  const debOrg   = useRef(); const debPerson = useRef();

  const set = (k,v) => setLead(p=>({...p,[k]:v}));
  const setScore = (id,v) => setLead(p=>({...p,scores:{...p.scores,[id]:v}}));
  const score = calcScore(lead.scores); const tier = getTier(score);
  const filled = CRITERIA.filter(c=>lead.scores[c.id]>0).length;
  const pendingTasks = (lead.tasks||[]).filter(t=>!t.done).length;

  // Debounced org search
  const searchOrg = useCallback((term) => {
    if (!term || term.length < 2) { setOrgSearch([]); setShowOrgDrop(false); return; }
    clearTimeout(debOrg.current);
    debOrg.current = setTimeout(async () => {
      setOrgLoading(true);
      console.log('[UI] Searching orgs:', term);
      const res = await apiGet(`/api/pd/search-orgs?term=${encodeURIComponent(term)}`);
      setOrgSearch(res.items || []);
      setShowOrgDrop(true);
      setOrgLoading(false);
    }, 600);
  }, []);

  const searchPerson = useCallback((term) => {
    if (!term || term.length < 2) { setPersonSearch([]); setShowPersonDrop(false); return; }
    clearTimeout(debPerson.current);
    debPerson.current = setTimeout(async () => {
      setPersonLoading(true);
      console.log('[UI] Searching persons:', term);
      const res = await apiGet(`/api/pd/search-persons?term=${encodeURIComponent(term)}`);
      setPersonSearch(res.items || []);
      setShowPersonDrop(true);
      setPersonLoading(false);
    }, 600);
  }, []);

  const handleSave = async () => {
    if (!lead.company.trim()) return;
    setSaving(true);
    console.log('[UI] Save started:', lead.company, '| selectedOrg:', selectedOrg?.id, '| selectedPerson:', selectedPerson?.id);

    const body = {
      leadName:         lead.leadName || `MES - ${lead.company}`,
      company:          lead.company,
      contact:          lead.contact,
      role:             lead.role,
      phone:            lead.phone,
      email:            lead.email,
      notes:            lead.notes,
      label:            lead.label,
      scores:           lead.scores,
      fitScore:         fitScore(lead.scores),
      bantScore:        bantScore(lead.scores),
      totalScore:       score,
      tier:             { label: tier.label, emoji: tier.emoji },
      existingOrgId:    selectedOrg?.id    || null,
      existingPersonId: selectedPerson?.id || null,
    };

    const result = await apiPost('/api/pd/create-lead', body);
    console.log('[UI] Create lead result:', result);

    if (result.ok) {
      setPdResult(result.result);
      toast.success(`ליד נוצר ב-Pipedrive! 🎉 Lead ID: ${result.result.leadId}`);
      result.log?.forEach(l => console.log('[PD Log]', l));

      const updated = {
        ...lead,
        pipedriveLeadId:  result.result.leadId,
        pipedriveOrgId:   result.result.orgId,
        pipedrivePersonId:result.result.personId,
        pipedriveOrg:     selectedOrg || (lead.company ? { id: result.result.orgId, name: lead.company } : null),
        pipedrivePerson:  selectedPerson || (lead.contact ? { id: result.result.personId, name: lead.contact } : null),
      };
      onSave(updated);
    } else {
      toast.error('שגיאה: ' + result.error);
      console.error('[UI] Create lead failed:', result.error);
    }
    setSaving(false);
  };

  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"#000c",zIndex:100,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:`1px solid ${tier.accent}44`,borderRadius:16,width:"100%",maxWidth:800,maxHeight:"92vh",display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:`0 0 40px ${tier.accent}12`}}>
        {/* Header */}
        <div style={{padding:"14px 22px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:17,fontWeight:700,color:"#e2e8f0"}}>{isNew?"➕ ליד חדש":`✏️ ${lead.company||"עריכת ליד"}`}</div>
            <div style={{fontSize:11,color:"#475569",marginTop:2}}>
              {filled}/{CRITERIA.length} קריטריונים · {pendingTasks} משימות ·
              {selectedOrg ? <span style={{color:"#22c55e"}}> 🏢 {selectedOrg.name} (ID:{selectedOrg.id})</span>
                : <span style={{color:"#334155"}}> יצור ב-Pipedrive</span>}
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{textAlign:"center"}}><div style={{fontSize:32,fontWeight:900,color:tier.accent,fontFamily:"monospace",lineHeight:1}}>{score}</div><div style={{fontSize:10,color:tier.accent}}>{tier.emoji} {tier.label}</div></div>
            <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button>
          </div>
        </div>
        <div style={{height:3,background:"#0a0f1a"}}><div style={{height:"100%",width:`${score}%`,background:tier.accent,transition:"width .4s"}}/></div>

        {/* Tabs */}
        <div style={{display:"flex",borderBottom:"1px solid #1e293b",background:"#0a0f1a"}}>
          {STEPS.map((s,i)=>(
            <button key={s.key} onClick={()=>setStep(i)} style={{flex:1,padding:"10px 0",background:"none",border:"none",cursor:"pointer",fontSize:13,fontWeight:step===i?700:400,fontFamily:"inherit",color:step===i?tier.accent:"#475569",borderBottom:`2px solid ${step===i?tier.accent:"transparent"}`,transition:"all .15s"}}>
              {s.icon} {s.label}{s.key==="tasks"&&pendingTasks>0&&<span style={{marginRight:4,background:"#f59e0b",color:"#000",borderRadius:8,padding:"1px 5px",fontSize:10,fontWeight:700}}>{pendingTasks}</span>}
            </button>
          ))}
        </div>

        {/* Body */}
        <div style={{flex:1,overflowY:"auto",padding:22}}>

          {/* ── STEP 0: Info + Duplicate Check ── */}
          {step===0&&(
            <div>
              {/* PD Result Banner */}
              {pdResult&&(
                <div style={{marginBottom:16,padding:"10px 14px",background:"#22c55e18",border:"1px solid #22c55e44",borderRadius:10,fontSize:12,color:"#22c55e"}}>
                  ✅ <strong>נוצר ב-Pipedrive!</strong>
                  {pdResult.orgId&&` · Org: ${pdResult.orgId}`}
                  {pdResult.personId&&` · Person: ${pdResult.personId}`}
                  {pdResult.leadId&&` · Lead: ${pdResult.leadId}`}
                </div>
              )}

              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                {/* Company + org search */}
                <div style={{gridColumn:"1/-1",position:"relative"}}>
                  <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>
                    שם חברה *
                    {orgLoading&&<span style={{color:"#f59e0b",marginRight:6,fontSize:11}}> ⏳ מחפש...</span>}
                  </label>
                  {selectedOrg ? (
                    <div style={{marginBottom:6}}>
                      <LinkedBadge type="org" item={selectedOrg} onClear={()=>{setSelectedOrg(null);setShowOrgDrop(false);}}/>
                    </div>
                  ) : (
                    <input value={lead.company||""} onChange={e=>{set("company",e.target.value);searchOrg(e.target.value);}}
                      placeholder="שם החברה" style={S.input}
                      onBlur={()=>setTimeout(()=>setShowOrgDrop(false),200)}/>
                  )}
                  {showOrgDrop&&!selectedOrg&&(
                    <DuplicateList items={orgSearch} type="org"
                      onSelect={org=>{setSelectedOrg(org);set("company",org.name);setShowOrgDrop(false);console.log('[UI] Org selected:',org.id,org.name);}}
                      onCreateNew={()=>setShowOrgDrop(false)}/>
                  )}
                </div>

                {/* Lead name */}
                <div style={{gridColumn:"1/-1"}}>
                  <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>שם ליד (Lead Name)</label>
                  <input value={lead.leadName||""} onChange={e=>set("leadName",e.target.value)}
                    placeholder={`MES - ${lead.company||"חברה"}`} style={S.input}/>
                </div>

                {/* Contact + person search */}
                <div style={{gridColumn:"1/-1",position:"relative"}}>
                  <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>
                    איש קשר
                    {personLoading&&<span style={{color:"#f59e0b",marginRight:6,fontSize:11}}> ⏳ מחפש...</span>}
                  </label>
                  {selectedPerson ? (
                    <div style={{marginBottom:6}}>
                      <LinkedBadge type="person" item={selectedPerson} onClear={()=>{setSelectedPerson(null);setShowPersonDrop(false);}}/>
                    </div>
                  ) : (
                    <input value={lead.contact||""} onChange={e=>{set("contact",e.target.value);searchPerson(e.target.value);}}
                      placeholder="שם איש הקשר" style={S.input}
                      onBlur={()=>setTimeout(()=>setShowPersonDrop(false),200)}/>
                  )}
                  {showPersonDrop&&!selectedPerson&&(
                    <DuplicateList items={personSearch} type="person"
                      onSelect={p=>{setSelectedPerson(p);set("contact",p.name);if(p.email)set("email",p.email);if(p.phone)set("phone",p.phone);setShowPersonDrop(false);console.log('[UI] Person selected:',p.id,p.name);}}
                      onCreateNew={()=>setShowPersonDrop(false)}/>
                  )}
                </div>

                {[{k:"role",label:"תפקיד",ph:"מנהל ייצור"},{k:"phone",label:"טלפון",ph:"050-1234567"},{k:"email",label:"אימייל",ph:"contact@company.com"},{k:"label",label:"תווית Pipedrive",ph:"Hot / Warm / Cold"}].map(f=>(
                  <div key={f.k}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>{f.label}</label><input value={lead[f.k]||""} onChange={e=>set(f.k,e.target.value)} placeholder={f.ph} style={S.input}/></div>
                ))}
                <div style={{gridColumn:"1/-1"}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>הערות שיחה</label><textarea value={lead.notes||""} onChange={e=>set("notes",e.target.value)} rows={4} style={{...S.input,resize:"vertical"}} placeholder="סיכום השיחה, כאבים שעלו..."/></div>
              </div>

              {/* User badge */}
              {user&&<div style={{marginTop:12,padding:"6px 12px",background:"#0a0f1a",borderRadius:8,border:"1px solid #1e293b",fontSize:11,color:"#475569"}}>
                👤 {user.name} · {user.email} · Pipedrive ID: {user.id}
              </div>}
            </div>
          )}

          {step===1&&(<div><div style={{...S.card,marginBottom:16,fontSize:12,color:"#475569",borderColor:"#1e3a5f"}}>🏭 <strong style={{color:"#94a3b8"}}>Basic Fit</strong> — {fitScore(lead.scores)}/60</div>{CRITERIA.filter(c=>c.group==="fit").map(c=><CriterionRow key={c.id} c={c} value={lead.scores[c.id]} onChange={setScore}/>)}</div>)}

          {step===2&&(<div>
            <div style={{...S.card,marginBottom:16,fontSize:12,color:"#475569",borderColor:"#1e3a5f"}}>💰 <strong style={{color:"#94a3b8"}}>BANT</strong> — {bantScore(lead.scores)}/40</div>
            {CRITERIA.filter(c=>c.group==="bant").map(c=><CriterionRow key={c.id} c={c} value={lead.scores[c.id]} onChange={setScore}/>)}
            <div style={{...S.card,marginTop:20}}><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,textAlign:"center"}}>
              {[{l:"Fit",v:fitScore(lead.scores),m:60},{l:"BANT",v:bantScore(lead.scores),m:40},{l:'סה"כ',v:score,m:100,a:tier.accent}].map(x=>(
                <div key={x.l} style={{padding:10,background:"#070d14",borderRadius:8,border:`1px solid ${x.a||"#1e293b"}`}}>
                  <div style={{fontSize:22,fontWeight:900,color:x.a||"#e2e8f0",fontFamily:"monospace"}}>{x.v}</div>
                  <div style={{fontSize:10,color:"#475569"}}>{x.l}/{x.m}</div>
                </div>
              ))}
            </div></div>
          </div>)}

          {step===3&&<TaskStep tasks={lead.tasks||[]} onChange={ts=>set("tasks",ts)}/>}
        </div>

        {/* Footer */}
        <div style={{padding:"12px 22px",borderTop:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <div style={{display:"flex",gap:8}}>
            {!isNew&&(confirmDel
              ?<><span style={{fontSize:12,color:"#fca5a5",lineHeight:"32px"}}>בטוח?</span><button onClick={()=>onDelete(lead.id)} style={{...S.btn,background:"#ef4444",color:"#fff",fontWeight:700,fontSize:12,padding:"5px 12px"}}>מחק ✓</button><button onClick={()=>setConfirmDel(false)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,padding:"5px 10px"}}>ביטול</button></>
              :<button onClick={()=>setConfirmDel(true)} style={{...S.btn,background:"#7f1d1d22",color:"#fca5a5",border:"1px solid #7f1d1d44",fontSize:12}}>🗑 מחק</button>)}
            <button onClick={()=>setShowAI(true)} style={{...S.btn,background:"#8b5cf620",color:"#8b5cf6",border:"1px solid #8b5cf644",fontSize:12}}>🤖 מייל AI</button>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <div style={{display:"flex",gap:5}}>{STEPS.map((_,i)=><div key={i} onClick={()=>setStep(i)} style={{width:7,height:7,borderRadius:"50%",cursor:"pointer",background:i===step?tier.accent:"#1e293b"}}/>)}</div>
            {step>0&&<button onClick={()=>setStep(s=>s-1)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:13}}>← קודם</button>}
            {step<STEPS.length-1
              ?<button onClick={()=>setStep(s=>s+1)} disabled={!lead.company?.trim()} style={{...S.btn,background:tier.accent,color:"#000",fontWeight:700,opacity:lead.company?.trim()?1:.4}}>הבא →</button>
              :<button onClick={handleSave} disabled={!lead.company?.trim()||saving} style={{...S.btn,background:tier.accent,color:"#000",fontWeight:700,opacity:lead.company?.trim()&&!saving?1:.4}}>
                {saving?"⏳ שומר...":isNew?"צור ב-Pipedrive ✓":"שמור שינויים ✓"}
              </button>}
          </div>
        </div>
      </div>
      {showAI&&<AIEmailModal lead={lead} user={user} onClose={()=>setShowAI(false)}/>}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   IMPORT MODAL
═══════════════════════════════════════════════ */
function ImportModal({ onImport, onClose }) {
  const [step,setStep]=useState("upload"); const [rows,setRows]=useState([]); const [selected,setSelected]=useState(new Set()); const [err,setErr]=useState(""); const [dragging,setDragging]=useState(false); const inputRef=useRef();
  const parseFile=file=>{setErr("");const reader=new FileReader();reader.onload=e=>{try{const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});const json=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});if(!json.length){setErr("הגיליון ריק");return;}setRows(json);setSelected(new Set(json.map((_,i)=>i)));setStep("preview");}catch(ex){setErr("שגיאה: "+ex.message);}};reader.readAsArrayBuffer(file);};
  const doImport=()=>{const TITLE=Object.keys(rows[0]).find(k=>/title/i.test(k));const LABEL=Object.keys(rows[0]).find(k=>/label/i.test(k));const DATE=Object.keys(rows[0]).find(k=>/created/i.test(k));const NACT=Object.keys(rows[0]).find(k=>/next.activity/i.test(k));const leads=[...selected].map(i=>{const r=rows[i];const title=r[TITLE]||"(ללא שם)";const tasks=r[NACT]?[{id:uid(),type:"follow",title:`פ.ה: ${r[NACT]}`,dueDate:"",time:"",notes:"מ-Pipedrive",done:false,createdAt:new Date().toISOString()}]:[];return{id:uid(),company:title.replace(/ lead$/i,"").replace(/ MES\b/i,"").trim(),name:title,leadName:"",contact:"",phone:"",email:"",role:"",label:r[LABEL]||"",tasks,notes:"",scores:{},source:"pipedrive",importedAt:r[DATE]||"",createdAt:new Date().toISOString()};});onImport(leads);onClose();};
  const TITLE_COL=rows[0]?Object.keys(rows[0]).find(k=>/title/i.test(k)):null;
  return(
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:"1px solid #1e293b",borderRadius:16,width:"100%",maxWidth:700,maxHeight:"88vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:"18px 24px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:17,fontWeight:700,color:"#e2e8f0"}}>⬆️ ייבוא מ-Pipedrive</div><button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button></div>
        <div style={{flex:1,overflowY:"auto",padding:24}}>
          {step==="upload"&&(<>
            <div style={{...S.card,marginBottom:20,borderColor:"#1e3a5f",fontSize:12,color:"#94a3b8",lineHeight:2}}><strong style={{color:"#60a5fa"}}>ייצוא מ-Pipedrive:</strong> Leads Inbox → ⋮ → Export leads → XLSX</div>
            <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)} onDrop={e=>{e.preventDefault();setDragging(false);parseFile(e.dataTransfer.files[0]);}} onClick={()=>inputRef.current.click()} style={{border:`2px dashed ${dragging?"#f59e0b":"#1e293b"}`,borderRadius:12,padding:40,textAlign:"center",cursor:"pointer",background:dragging?"#f59e0b08":"#0a0f1a",transition:"all .2s"}}>
              <div style={{fontSize:48,marginBottom:12}}>📂</div><div style={{fontSize:15,color:"#94a3b8",marginBottom:6,fontWeight:600}}>גרור קובץ XLSX / CSV</div><div style={{fontSize:12,color:"#475569"}}>או לחץ לבחירת קובץ</div>
              <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>parseFile(e.target.files[0])}/>
            </div>
            {err&&<div style={{marginTop:12,color:"#ef4444",fontSize:13,textAlign:"center"}}>{err}</div>}
          </>)}
          {step==="preview"&&(<>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div style={{fontSize:14,color:"#94a3b8"}}>נמצאו <strong style={{color:"#f59e0b"}}>{rows.length}</strong> · נבחרו <strong style={{color:"#22c55e"}}>{selected.size}</strong></div>
              <div style={{display:"flex",gap:8}}><button onClick={()=>setSelected(new Set(rows.map((_,i)=>i)))} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:11,padding:"4px 10px"}}>בחר הכל</button><button onClick={()=>setSelected(new Set())} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:11,padding:"4px 10px"}}>נקה</button></div>
            </div>
            <div style={{maxHeight:380,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
              {rows.map((r,i)=>{const title=TITLE_COL?r[TITLE_COL]:"(ללא שם)";const sel=selected.has(i);return(<div key={i} onClick={()=>setSelected(p=>{const n=new Set(p);sel?n.delete(i):n.add(i);return n;})} style={{display:"flex",alignItems:"center",gap:12,padding:"8px 12px",borderRadius:8,cursor:"pointer",background:sel?"#f59e0b08":"#0a0f1a",border:`1px solid ${sel?"#f59e0b44":"#1e293b"}`}}><input type="checkbox" checked={sel} onChange={()=>{}} style={{accentColor:"#f59e0b",cursor:"pointer",flexShrink:0}}/><div style={{flex:1,minWidth:0,fontSize:13,color:"#e2e8f0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{title}</div></div>);})}
            </div>
          </>)}
        </div>
        <div style={{padding:"14px 24px",borderTop:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"flex-end",gap:10}}>
          {step==="preview"&&<button onClick={()=>{setStep("upload");setRows([]);}} style={{...S.btn,background:"#1e293b",color:"#94a3b8"}}>← חזור</button>}
          <button onClick={onClose} style={{...S.btn,background:"#1e293b",color:"#94a3b8"}}>ביטול</button>
          {step==="preview"&&<button onClick={doImport} disabled={selected.size===0} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,opacity:selected.size?1:.4}}>ייבא {selected.size} →</button>}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   LEAD CARD
═══════════════════════════════════════════════ */
function LeadCard({lead,onClick}){
  const score=calcScore(lead.scores);const tier=getTier(score);
  const pTask=(lead.tasks||[]).filter(t=>!t.done);const overT=pTask.filter(t=>isOverdue(t.dueDate));const filled=CRITERIA.filter(c=>lead.scores[c.id]>0).length;
  return(
    <div onClick={onClick} className="card-hover" style={{background:"#0d1520",border:`1px solid ${tier.accent}33`,borderRadius:12,padding:16,cursor:"pointer",position:"relative",overflow:"hidden",transition:"all .2s"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:tier.accent}}/>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
        <div style={{flex:1,paddingLeft:8,minWidth:0}}>
          <div style={{fontWeight:700,fontSize:14,color:"#e2e8f0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{lead.company||"—"}</div>
          <div style={{fontSize:11,color:"#64748b",marginTop:1}}>{lead.contact}{lead.role?` · ${lead.role}`:""}</div>
          {lead.pipedriveLeadId&&<div style={{fontSize:10,color:"#22c55e",marginTop:2}}>✅ Pipedrive Lead: {lead.pipedriveLeadId}</div>}
          {lead.label&&<div style={{fontSize:10,color:"#94a3b8",marginTop:1}}>🏷 {lead.label}</div>}
        </div>
        <div style={{textAlign:"center",flexShrink:0}}><div style={{fontSize:26,fontWeight:900,color:tier.accent,fontFamily:"monospace",lineHeight:1}}>{score}</div><div style={{fontSize:9,color:tier.accent}}>{tier.emoji} {tier.label}</div></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"3px 10px",marginBottom:8}}>
        {[{l:"Fit",v:fitScore(lead.scores),m:60},{l:"BANT",v:bantScore(lead.scores),m:40}].map(x=>(
          <div key={x.l}><div style={{fontSize:9,color:"#475569",marginBottom:2}}>{x.l} {x.v}/{x.m}</div>
            <div style={{height:3,background:"#1e2a3a",borderRadius:2}}><div style={{width:`${(x.v/x.m)*100}%`,height:"100%",borderRadius:2,background:x.v/x.m>.66?"#22c55e":x.v/x.m>.33?"#f59e0b":"#60a5fa",transition:"width .3s"}}/></div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontSize:10,color:"#334155"}}>{filled}/{CRITERIA.length} ק׳</span>
        <div style={{display:"flex",gap:4}}>
          {pTask.length>0&&<span style={{fontSize:10,padding:"2px 6px",borderRadius:8,background:overT.length?"#7f1d1d":"#1e293b",color:overT.length?"#fca5a5":"#64748b"}}>{overT.length?"⚠️":"✅"} {pTask.length}</span>}
          {lead.pipedriveLeadId&&<span style={{fontSize:10,padding:"2px 6px",borderRadius:8,background:"#22c55e18",color:"#22c55e"}}>PD</span>}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   SETTINGS MODAL
═══════════════════════════════════════════════ */
function SettingsModal({ config, user, onClose }) {
  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:"1px solid #1e293b",borderRadius:16,width:"100%",maxWidth:520,overflow:"hidden"}}>
        <div style={{padding:"18px 24px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:17,fontWeight:700,color:"#e2e8f0"}}>⚙️ הגדרות</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button>
        </div>
        <div style={{padding:24,display:"flex",flexDirection:"column",gap:16}}>
          {/* User info */}
          {user&&(
            <div style={{...S.card,borderColor:"#22c55e33"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#22c55e",marginBottom:8}}>✅ מחובר ל-Pipedrive</div>
              <div style={{fontSize:13,color:"#e2e8f0",marginBottom:4}}>{user.name}</div>
              <div style={{fontSize:12,color:"#64748b"}}>{user.email}</div>
              <div style={{fontSize:11,color:"#334155",marginTop:4}}>Pipedrive User ID: {user.id}</div>
            </div>
          )}
          <div style={{...S.card,fontSize:12,color:"#475569",lineHeight:1.8}}>
            <strong style={{color:"#60a5fa"}}>OAuth Status:</strong> {config?.oauthEnabled?"✅ מופעל":"❌ לא מוגדר"}<br/>
            <strong style={{color:"#60a5fa"}}>PD_TOKEN fallback:</strong> {config?.fallbackToken?"✅":"—"}
          </div>
          <a href="/api/auth/logout" style={{...S.btn,background:"#7f1d1d22",color:"#fca5a5",border:"1px solid #7f1d1d44",textDecoration:"none",textAlign:"center"}}>
            🚪 התנתק
          </a>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   APP ROOT
═══════════════════════════════════════════════ */
const VIEWS=[{key:"leads",label:"לידים",icon:"📋"},{key:"calendar",label:"לוח שנה",icon:"📅"}];

export default function App() {
  const [leads, setLeads]         = useState([]);
  const [authUser, setAuthUser]   = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [config, setConfig]       = useState(null);
  const [view, setView]           = useState("leads");
  const [modal, setModal]         = useState(null);
  const [showImport, setShowImport] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [filter, setFilter]       = useState("all");
  const [sort, setSort]           = useState("score");
  const [ready, setReady]         = useState(false);
  const toast = useToast();
  const { isMobile } = useBreakpoint();

  // Check if admin route
  const isAdmin = window.location.pathname.startsWith('/admin');

  // Load auth state + config
  useEffect(()=>{
    Promise.all([
      apiGet('/api/auth/me').catch(()=>null),
      apiGet('/api/config').catch(()=>null),
    ]).then(([me, cfg])=>{
      if (me?.ok) { setAuthUser(me.user); console.log('[App] Logged in as', me.user.name); }
      else console.log('[App] Not logged in');
      setConfig(cfg);
      setAuthLoading(false);
    });
  },[]);

  // Load leads from storage
  useEffect(()=>{
    (async()=>{
      try {
        const r = await window.storage.get("sdr_v5");
        if (r) { const d=JSON.parse(r.value); setLeads(d.leads||[]); console.log('[App] Loaded',d.leads?.length,'leads'); }
      } catch(e) { console.error('[App] Storage load error:',e); }
      setReady(true);
    })();
  },[]);

  const persist = async (l) => {
    setLeads(l);
    try { await window.storage.set("sdr_v5", JSON.stringify({leads:l})); console.log('[App] Saved',l.length,'leads'); }
    catch(e) { console.error('[App] Storage save error:',e); toast.error('שגיאת שמירה: '+e.message); }
  };

  const handleSaveLead = async lead => {
    const updated = lead.id ? leads.map(l=>l.id===lead.id?lead:l) : [...leads,{...lead,id:uid(),createdAt:new Date().toISOString()}];
    await persist(updated);
    if (!lead.id) toast.success(`ליד חדש: ${lead.company}`);
    setModal(null);
  };

  const handleDelete = async id => { await persist(leads.filter(l=>l.id!==id)); toast.info('הליד נמחק'); setModal(null); };
  const handleImport = async imported => { const ex=new Set(leads.map(l=>l.company?.trim().toLowerCase())); const fresh=imported.filter(l=>!ex.has(l.company?.trim().toLowerCase())); await persist([...leads,...fresh]); toast.success(`יובאו ${fresh.length} לידים`); };

  const exportXLSX = () => {
    const rows=leads.map(l=>{const total=calcScore(l.scores);const tier=getTier(total);const openTasks=(l.tasks||[]).filter(t=>!t.done);const next=openTasks.sort((a,b)=>(a.dueDate||"9999").localeCompare(b.dueDate||"9999"))[0];
      return{"שם חברה":l.company,"שם ליד":l.leadName||"","איש קשר":l.contact,"תפקיד":l.role||"","טלפון":l.phone,"אימייל":l.email,"סיווג":tier.emoji+" "+tier.label,"Fit":fitScore(l.scores),"BANT":bantScore(l.scores),"ציון כולל":total,"תאריך פתיחה":String(l.importedAt||l.createdAt||"").slice(0,10),"PD Lead ID":l.pipedriveLeadId||"","PD Org ID":l.pipedriveOrgId||"","PD Person ID":l.pipedrivePersonId||"","משימות פתוחות":openTasks.length,"פעילות הבאה":next?.title||"","הערות":l.notes||""};});
    const ws=XLSX.utils.json_to_sheet(rows);ws["!cols"]=[{wch:20},{wch:18},{wch:16},{wch:14},{wch:14},{wch:24},{wch:10},{wch:6},{wch:6},{wch:10},{wch:14},{wch:16},{wch:12},{wch:14},{wch:12},{wch:24},{wch:40}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"לידים");
    XLSX.writeFile(wb,`SDR_Qualify_${isoDate(new Date())}.xlsx`);
  };

  const counts={hot:0,warm:0,cold:0}; leads.forEach(l=>{counts[getTier(calcScore(l.scores)).key]++;});
  const totalPending=leads.reduce((s,l)=>s+(l.tasks||[]).filter(t=>!t.done).length,0);
  const sorted=[...leads].filter(l=>filter==="all"||getTier(calcScore(l.scores)).key===filter).sort((a,b)=>{if(sort==="score")return calcScore(b.scores)-calcScore(a.scores);if(sort==="name")return(a.company||"").localeCompare(b.company||"","he");return new Date(b.createdAt||0)-new Date(a.createdAt||0);});

  // Loading
  if (authLoading) return (
    <div style={{minHeight:"100vh",background:"#070d14",display:"flex",alignItems:"center",justifyContent:"center",color:"#f59e0b",fontFamily:"monospace"}}>
      <div>⏳ בודק חיבור...</div>
    </div>
  );

  // Admin route
  if (isAdmin) {
    if (!authUser) return <LoginPage config={config}/>;
    return <AdminDashboard user={authUser}/>;
  }

  // Not logged in + OAuth required
  if (!authUser && config?.oauthEnabled) return <LoginPage config={config}/>;

  return (
    <div style={{height:"100vh",background:"#070d14",color:"#e2e8f0",fontFamily:"'Segoe UI',Arial,sans-serif",direction:"rtl",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <style>{CSS}</style>

      {/* Header */}
      <div style={{background:"#070d14",borderBottom:"1px solid #1e293b",padding:isMobile?"8px 12px":"10px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0,gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:isMobile?10:18}}>
          <div style={{fontSize:isMobile?15:18,fontWeight:900,color:"#e2e8f0",letterSpacing:-.5}}>
            SDR<span style={{color:"#f59e0b"}}>.</span>qualify
          </div>
          <div style={{display:"flex",gap:3}}>
            {VIEWS.map(v=><button key={v.key} onClick={()=>setView(v.key)} style={{...S.btn,padding:isMobile?"4px 8px":"5px 12px",fontSize:isMobile?11:12,background:view===v.key?"#f59e0b20":"none",color:view===v.key?"#f59e0b":"#475569",border:`1px solid ${view===v.key?"#f59e0b44":"transparent"}`}}>{v.icon}{!isMobile&&" "+v.label}</button>)}
          </div>
        </div>
        <div style={{display:"flex",gap:isMobile?8:14,alignItems:"center"}}>
          {!isMobile&&[{l:"סה״כ",v:leads.length,c:"#e2e8f0"},{l:"🔥",v:counts.hot,c:"#ef4444"},{l:"⚡",v:counts.warm,c:"#f59e0b"},{l:"❄️",v:counts.cold,c:"#60a5fa"},{l:"📋",v:totalPending,c:"#64748b"}].map(s=>(
            <div key={s.l} style={{textAlign:"center"}}><div style={{fontSize:17,fontWeight:900,color:s.c,fontFamily:"monospace"}}>{s.v}</div><div style={{fontSize:9,color:"#334155"}}>{s.l}</div></div>
          ))}
          {/* User avatar */}
          {authUser&&(
            <div style={{display:"flex",alignItems:"center",gap:6,padding:"4px 10px",background:"#1e293b",borderRadius:20,cursor:"pointer"}} onClick={()=>setShowSettings(true)}>
              <div style={{width:22,height:22,borderRadius:"50%",background:"#f59e0b",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:"#000"}}>{authUser.name?.[0]}</div>
              {!isMobile&&<span style={{fontSize:11,color:"#94a3b8"}}>{authUser.name?.split(" ")[0]}</span>}
            </div>
          )}
          {config?.adminEmail===authUser?.email&&(
            <a href="/admin" style={{...S.btn,background:"#8b5cf620",color:"#8b5cf6",border:"1px solid #8b5cf644",fontSize:11,textDecoration:"none",padding:"5px 10px"}}>
              👑 Admin
            </a>
          )}
        </div>
      </div>

      {/* Toolbar */}
      <div style={{padding:isMobile?"6px 10px":"8px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #1e293b",background:"#0a0f1a",gap:8,flexWrap:"wrap",flexShrink:0}}>
        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
          {[{k:"all",l:`הכל`,c:"#e2e8f0"},{k:"hot",l:"🔥",c:"#ef4444"},{k:"warm",l:"⚡",c:"#f59e0b"},{k:"cold",l:"❄️",c:"#60a5fa"}].map(f=>(
            <button key={f.k} onClick={()=>setFilter(f.k)} style={{...S.btn,fontSize:11,padding:"4px 9px",background:filter===f.k?`${f.c}18`:"#0d1520",color:filter===f.k?f.c:"#475569",border:`1px solid ${filter===f.k?f.c:"#1e293b"}`}}>{f.l}{f.k!=="all"&&counts[f.k]>0?` ${counts[f.k]}`:""}</button>
          ))}
          {!isMobile&&<select value={sort} onChange={e=>setSort(e.target.value)} style={{...S.input,padding:"4px 8px",width:"auto",fontSize:11}}>
            <option value="score">ציון</option><option value="name">שם</option><option value="date">תאריך</option>
          </select>}
        </div>
        <div style={{display:"flex",gap:6}}>
          {!isMobile&&<button onClick={()=>setShowImport(true)} style={{...S.btn,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",padding:"4px 10px",fontSize:11}}>⬆️ ייבא</button>}
          {!isMobile&&<button onClick={exportXLSX} disabled={!leads.length} style={{...S.btn,background:"#14532d22",color:"#22c55e",border:"1px solid #22c55e44",padding:"4px 10px",fontSize:11,opacity:leads.length?1:.5}}>📊 Excel</button>}
          <button onClick={()=>setModal({lead:{company:"",leadName:"",contact:"",role:"",phone:"",email:"",notes:"",label:"",scores:{},tasks:[],id:null}})}
            style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,padding:"4px 12px",fontSize:isMobile?13:12}}>
            + {!isMobile&&"ליד חדש"}
          </button>
        </div>
      </div>

      {/* Content */}
      <div style={{flex:1,overflowY:"auto",padding:isMobile?10:18}}>
        {sorted.length===0 ? (
          <div style={{textAlign:"center",color:"#334155",marginTop:60}}>
            <div style={{fontSize:52,marginBottom:12}}>📋</div>
            <div style={{fontSize:16,color:"#475569",marginBottom:8}}>{filter!=="all"?"אין לידים בקטגוריה":"אין לידים — התחל בייבוא או הוסף ידנית"}</div>
            {filter==="all"&&<button onClick={()=>setShowImport(true)} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700}}>⬆️ ייבא מ-Pipedrive</button>}
          </div>
        ) : (
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
            {sorted.map(l=><LeadCard key={l.id} lead={l} onClick={()=>setModal({lead:l})}/>)}
          </div>
        )}
        {isMobile&&leads.length>0&&(
          <div style={{display:"flex",gap:8,marginTop:14}}>
            <button onClick={()=>setShowImport(true)} style={{...S.btn,flex:1,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",fontSize:12}}>⬆️ ייבא</button>
            <button onClick={exportXLSX} style={{...S.btn,flex:1,background:"#14532d22",color:"#22c55e",border:"1px solid #22c55e44",fontSize:12}}>📊 Excel</button>
          </div>
        )}
      </div>

      {/* Footer */}
      {!isMobile&&<div style={{padding:"5px 20px",borderTop:"1px solid #1e293b",background:"#070d14",display:"flex",gap:20,alignItems:"center",fontSize:11,color:"#334155",flexShrink:0}}>
        <span>🔥≥67 · ⚡36-66 · ❄️0-35</span>
        <span style={{marginRight:"auto"}}>SDR.qualify v3 · מקסימום 100 · Fit×60 + BANT×40</span>
      </div>}

      {/* Modals */}
      {modal&&<LeadModal lead={modal.lead} onSave={handleSaveLead} onDelete={handleDelete} onClose={()=>setModal(null)} user={authUser} toast={toast}/>}
      {showImport&&<ImportModal onImport={handleImport} onClose={()=>setShowImport(false)}/>}
      {showSettings&&<SettingsModal config={config} user={authUser} onClose={()=>setShowSettings(false)}/>}
      <ToastContainer toasts={toast.toasts}/>
    </div>
  );
}
