import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

/* ══ CONSTANTS ══ */
const CRITERIA = [
  { id:"industry",   group:"fit",  label:"סוג התעשייה",      low:"מיכלים / נגריות / תוכנה / הפצה",      mid:"CNC / אלקטרוניקה / יקבים",           high:"מזון / משקאות / פלסטיקה / כימיקאלים / פארמה" },
  { id:"lines",      group:"fit",  label:"קווי ייצור",        low:"ידניים / מיושנים",                    mid:"שילוב / חצי-אוטומטי",                 high:"אוטומטיים מתקדמים / High Volume" },
  { id:"complexity", group:"fit",  label:"מורכבות הייצור",    low:"מפעל קטן / Garage",                   mid:"בינוני, סטנדרטי",                     high:"מורכב, רב-שלבי — כאב ברור" },
  { id:"sites",      group:"fit",  label:"אתרי ייצור",        low:"1 אתר", mid:"1–2 אתרים",               high:"3+ אתרים" },
  { id:"employees",  group:"fit",  label:"עובדים",            low:"1–40",  mid:"40–80",                   high:"80+" },
  { id:"valuation",  group:"fit",  label:"שווי / מחזור",      low:"עד 10M ₪", mid:"10–100M ₪",            high:"100M+ ₪" },
  { id:"budget",     group:"bant", label:"Budget — תקציב",    low:"אין תקציב", mid:"תקציב כללי",          high:"תקציב מאושר ומוגדר" },
  { id:"authority",  group:"bant", label:"Authority — סמכות", low:"משתמש קצה", mid:"מנהל ביניים",         high:"C-Level / מקבל החלטה" },
  { id:"need",       group:"bant", label:"Need — צורך",       low:"עניין כללי", mid:"בעיה, מסתדרים",      high:"כאב אקוטי — פתרון דחוף" },
  { id:"timing",     group:"bant", label:"Timing — זמן",      low:"שנה הבאה", mid:"תוך 6 חודשים",         high:"מיידי / תוך 3 חודשים" },
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

/* ══ HELPERS ══ */
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2,5);
const calcScore = (s={}) => CRITERIA.reduce((t,c)=>t+(s[c.id]||0),0);
const fitScore  = (s={}) => CRITERIA.filter(c=>c.group==="fit").reduce((t,c)=>t+(s[c.id]||0),0);
const bantScore = (s={}) => CRITERIA.filter(c=>c.group==="bant").reduce((t,c)=>t+(s[c.id]||0),0);
const getTier   = n => n>=67?{key:"hot",label:"HOT",emoji:"🔥",accent:"#ef4444"}:n>=36?{key:"warm",label:"WARM",emoji:"⚡",accent:"#f59e0b"}:{key:"cold",label:"COLD",emoji:"❄️",accent:"#60a5fa"};
const fmtDate   = d => !d?"—":new Date(d).toLocaleDateString("he-IL",{day:"2-digit",month:"2-digit"});
const fmtFull   = d => !d?"—":new Date(d).toLocaleDateString("he-IL",{day:"2-digit",month:"2-digit",year:"numeric"});
const isoDate   = d => d?new Date(d).toISOString().slice(0,10):"";
const isOverdue = d => d && new Date(d) < new Date(new Date().toDateString());
const taskType  = t => TASK_TYPES.find(x=>x.key===t)||TASK_TYPES[6];

const outlookUrl = (task, lead) => {
  const date = task.dueDate || isoDate(new Date());
  const time = task.time || "09:00";
  const start = new Date(`${date}T${time}:00`);
  const end   = new Date(start.getTime()+60*60*1000);
  const body  = `ליד: ${lead?.company||""}\nאיש קשר: ${lead?.contact||""}\n${task.notes||""}`;
  return `https://outlook.office.com/calendar/deeplink/compose?subject=${encodeURIComponent(task.title)}&startdt=${start.toISOString()}&enddt=${end.toISOString()}&body=${encodeURIComponent(body)}&path=%2Fcalendar%2Faction%2Fcompose`;
};

const downloadICS = (task, lead) => {
  const date  = (task.dueDate||isoDate(new Date())).replace(/-/g,"");
  const time  = (task.time||"09:00").replace(":","")+"00";
  const [h,m] = (task.time||"09:00").split(":").map(Number);
  const endH  = String(h+1).padStart(2,"0")+String(m).padStart(2,"0")+"00";
  const desc  = `ליד: ${lead?.company||""}\\nאיש קשר: ${lead?.contact||""}\\n${task.notes||""}`.replace(/\n/g,"\\n");
  const ics = `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//SDR Qualify//EN\r\nBEGIN:VEVENT\r\nUID:${task.id}@sdrqualify\r\nDTSTAMP:${date}T${time}Z\r\nDTSTART:${date}T${time}\r\nDTEND:${date}T${endH}\r\nSUMMARY:${task.title}\r\nDESCRIPTION:${desc}\r\nEND:VEVENT\r\nEND:VCALENDAR`;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([ics],{type:"text/calendar;charset=utf-8"}));
  a.download = `${task.title.replace(/\s+/g,"_")}.ics`; a.click();
};

/* ══ PIPEDRIVE ══ */
const IS_VERCEL = typeof window !== 'undefined' && !window.location.hostname.includes('claude');
const PD = {
  base: IS_VERCEL ? '/api/pipedrive' : 'https://api.pipedrive.com/v1',
  async req(token,path,method="GET",body=null){
    if(!token) throw new Error("אין Token");
    let url, headers = {};
    if(IS_VERCEL){
      // Use serverless proxy — token goes in header, never in URL
      const cleanPath = path.replace(/^\//, '').replace(/\?.*/,'');
      const qs = path.includes('?') ? '&'+path.split('?')[1] : '';
      url = `/api/pipedrive?path=${encodeURIComponent(cleanPath)}${qs}`;
      headers = { 'Content-Type':'application/json', 'x-pd-token': token };
    } else {
      const sep = path.includes("?")?"&":"?";
      url = `https://api.pipedrive.com/v1${path}${sep}api_token=${token}`;
      if(body) headers['Content-Type'] = 'application/json';
    }
    const r = await fetch(url,{method,headers,body:body?JSON.stringify(body):undefined});
    const j = await r.json();
    if(!j.success) throw new Error(j.error||"PD error");
    return j;
  },
  searchOrg:(t,q)=>PD.req(t,`/organizations/search?term=${encodeURIComponent(q)}&fields=name&limit=5`),
  createOrg:(t,name)=>PD.req(t,"/organizations","POST",{name}),
  updateOrg:(t,id,data)=>PD.req(t,`/organizations/${id}`,"PUT",data),
  createPerson:(t,d)=>PD.req(t,"/persons","POST",d),
  addNote:(t,d)=>PD.req(t,"/notes","POST",d),
  createLead:(t,d)=>PD.req(t,"/leads","POST",d),
};

/* ══ STYLES ══ */
const S = {
  input:{width:"100%",background:"#0a0f1a",border:"1px solid #1e293b",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,outline:"none",boxSizing:"border-box"},
  btn:{padding:"8px 16px",borderRadius:8,cursor:"pointer",border:"1px solid transparent",fontSize:13,transition:"all .15s",fontFamily:"inherit"},
  card:{background:"#0d1520",border:"1px solid #1e293b",borderRadius:12,padding:16},
};
const CSS = `
* { box-sizing:border-box; }
body { margin:0; }
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:#070d14;}
::-webkit-scrollbar-thumb{background:#1e293b;border-radius:3px;}
input:focus,textarea:focus,select:focus{border-color:#f59e0b!important;outline:none;}
input::placeholder,textarea::placeholder{color:#334155;}
select option{background:#0a0f1a;color:#e2e8f0;}
.card-hover:hover{transform:translateY(-2px);box-shadow:0 6px 24px #0007;}
.ev:hover{filter:brightness(1.2);}
.ev{ touch-action:none; }
.slot-drop{background:#f59e0b10!important;border-color:#f59e0b44!important;}
.task-drag{cursor:grab;} .task-drag:active{cursor:grabbing;opacity:.7;}
/* ── Responsive ── */
@media(max-width:767px){
  .hide-mobile{display:none!important;}
  .mobile-full{width:100%!important;max-width:100%!important;border-radius:0!important;height:100%!important;max-height:100%!important;}
  .mobile-col{grid-template-columns:1fr!important;}
  .mobile-row{flex-direction:column!important;align-items:stretch!important;}
  .mobile-wrap{flex-wrap:wrap!important;}
  .mobile-p{padding:12px!important;}
  .stat-num{font-size:14px!important;}
  .cal-sidebar{display:none!important;}
  .modal-slide{animation:slideUp .25s ease;}
}
@media(min-width:768px) and (max-width:1023px){
  .hide-tablet{}
  .tablet-col2{grid-template-columns:1fr 1fr!important;}
}
@keyframes slideUp{from{transform:translateY(100%);opacity:0;}to{transform:translateY(0);opacity:1;}}
`;

/* ══ RESPONSIVE HOOK ══ */
function useBreakpoint() {
  const [w, setW] = useState(window.innerWidth);
  useEffect(()=>{
    const fn = ()=>setW(window.innerWidth);
    window.addEventListener("resize",fn);
    return ()=>window.removeEventListener("resize",fn);
  },[]);
  return { isMobile: w<768, isTablet: w>=768&&w<1024, isDesktop: w>=1024, width:w };
}

/* ══════════════════════════════════
   AI EMAIL MODAL
══════════════════════════════════ */
function AIEmailModal({ lead, onClose }) {
  const [loading, setLoading] = useState(true);
  const [draft, setDraft] = useState("");
  const [subject, setSubject] = useState("");
  const [error, setError] = useState("");

  const tier = getTier(calcScore(lead.scores||{}));
  const scoreSum = `Fit: ${fitScore(lead.scores)}/60 | BANT: ${bantScore(lead.scores)}/40 | סה"כ: ${calcScore(lead.scores)}/100 (${tier.label})`;
  const bantLabels = CRITERIA.filter(c=>c.group==="bant").map(c=>`${c.label}: ${lead.scores?.[c.id]||0}/9`).join(", ");

  // Fallback template if API fails
  const buildFallback = () => {
    const subj = `המשך שיחתנו — ${lead.company}`;
    const body = `שלום ${lead.contact||""},

תודה על השיחה המעניינת שקיימנו היום.

כפי שדיברנו, ${lead.notes ? lead.notes : `אנחנו ב-Trunovate מתמחים בפתרונות MES מודולריים עבור מפעלי ייצור כמו ${lead.company}.`}

אשמח לקבוע שיחת המשך קצרה של 30 דקות לצורך:
• הצגת דמו מותאם לתהליכי הייצור שלכם
• בחינת Quick Win ראשון שניתן ליישם תוך 4–6 שבועות

מתי נוח לך השבוע?

בברכה,
Roy
VP Business Development | Trunovate — PlantSharp MES`;
    setSubject(subj);
    setDraft(body);
  };

  useEffect(()=>{
    (async()=>{
      try {
        const prompt = `אתה מומחה מכירות B2B לתוכנת MES תעשייתי (PlantSharp MES של Trunovate).
כתוב מייל מקצועי בעברית לאיש הקשר הבא לאחר שיחת היכרות ראשונה.

פרטי הליד:
- חברה: ${lead.company}
- איש קשר: ${lead.contact||"לא ידוע"}
- תפקיד: ${lead.role||"לא ידוע"}
- ציון Qualification: ${scoreSum}
- BANT: ${bantLabels}
- הערות מהשיחה: ${lead.notes||"אין הערות"}

דרישות המייל:
1. פתיחה חמה ותזכורת קצרה לשיחה
2. 2-3 שורות סיכום הכאב/הצורך שעלה בשיחה
3. הצעה לשיחת המשך / דמו קצר
4. CTA ברור לקביעת זמן

כתוב רק את תוכן המייל. פורמט:
נושא: [שורת נושא]
---
[גוף המייל]`;

        const res = await fetch("https://api.anthropic.com/v1/messages", {
          method:"POST",
          headers:{"Content-Type":"application/json"},
          body:JSON.stringify({
            model:"claude-sonnet-4-20250514",
            max_tokens:1000,
            messages:[{role:"user",content:prompt}]
          })
        });

        if(!res.ok) { buildFallback(); setLoading(false); return; }

        const data = await res.json();

        // Handle API error response
        if(data.error || data.type==="error") {
          buildFallback(); setLoading(false); return;
        }

        const text = (data.content||[]).find(b=>b.type==="text")?.text || "";
        if(!text) { buildFallback(); setLoading(false); return; }

        const lines = text.split("\n");
        const subjectLine = lines.find(l=>l.startsWith("נושא:"));
        setSubject(subjectLine ? subjectLine.replace("נושא:","").trim() : `המשך שיחתנו — ${lead.company}`);
        const bodyStart = lines.findIndex(l=>l.trim()==="---");
        setDraft(bodyStart>=0 ? lines.slice(bodyStart+1).join("\n").trim() : text.replace(/^נושא:.*\n?/,"").replace(/^---\n?/,"").trim());

      } catch(e) {
        // Network error or anything else — use fallback template
        buildFallback();
      }
      setLoading(false);
    })();
  },[]);

  const openMailto = () => {
    const mailto = `mailto:${lead.email||""}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(draft)}`;
    window.open(mailto,"_blank");
  };

  const openOutlookWeb = () => {
    const url = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(lead.email||"")}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(draft)}`;
    window.open(url,"_blank");
  };

  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:"fixed",inset:0,background:"#000d",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#0d1520",border:"1px solid #8b5cf644",borderRadius:16,width:"100%",maxWidth:680,maxHeight:"92vh",display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"0 0 40px #8b5cf622"}}>
        <div style={{padding:"14px 18px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:"#e2e8f0"}}>🤖 סוכן AI — טיוטת מייל</div>
            <div style={{fontSize:11,color:"#475569",marginTop:2}}>{lead.company}{lead.contact?` · ${lead.contact}`:""}</div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button>
        </div>

        <div style={{flex:1,overflowY:"auto",padding:16}}>
          {loading ? (
            <div style={{textAlign:"center",padding:"40px 0"}}>
              <div style={{fontSize:32,marginBottom:12}}>⏳</div>
              <div style={{color:"#8b5cf6",fontSize:13}}>מנסח מייל מותאם אישית...</div>
            </div>
          ) : (
            <>
              <div style={{marginBottom:12}}>
                <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>נושא המייל</label>
                <input value={subject} onChange={e=>setSubject(e.target.value)} style={S.input}/>
              </div>
              <div>
                <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>גוף המייל</label>
                <textarea value={draft} onChange={e=>setDraft(e.target.value)}
                  rows={13} style={{...S.input,resize:"vertical",lineHeight:1.7,fontFamily:"inherit"}}/>
              </div>
              <div style={{marginTop:8,padding:"7px 12px",background:"#8b5cf610",border:"1px solid #8b5cf633",borderRadius:8,fontSize:11,color:"#8b5cf6"}}>
                💡 ניתן לערוך את הטיוטה לפני השליחה
              </div>
            </>
          )}
        </div>

        <div style={{padding:"12px 18px",borderTop:"1px solid #1e293b",background:"#0a0f1a",display:"flex",gap:8,flexWrap:"wrap"}}>
          {!loading&&(
            <>
              <button onClick={openMailto}
                style={{...S.btn,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",fontSize:12,padding:"6px 12px"}}>
                📧 פתח בדואר
              </button>
              <button onClick={openOutlookWeb}
                style={{...S.btn,background:"#0078d420",color:"#0078d4",border:"1px solid #0078d444",fontSize:12,padding:"6px 12px"}}>
                📆 Outlook
              </button>
              <button onClick={()=>navigator.clipboard?.writeText(`נושא: ${subject}\n\n${draft}`)}
                style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,padding:"6px 12px"}}>
                📋 העתק
              </button>
            </>
          )}
          <button onClick={onClose} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,padding:"6px 12px",marginRight:"auto"}}>סגור</button>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   TASK STEP (inside wizard)
══════════════════════════════════ */
function TaskStep({ tasks, onChange }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm]     = useState({type:"call",title:"",dueDate:"",time:"",notes:""});
  const setF = (k,v) => setForm(p=>({...p,[k]:v}));

  const add = () => {
    if(!form.title.trim()) return;
    onChange([...tasks,{id:uid(),...form,done:false,createdAt:new Date().toISOString()}]);
    setForm({type:"call",title:"",dueDate:"",time:"",notes:""});
    setAdding(false);
  };

  const toggle = id => onChange(tasks.map(t=>t.id===id?{...t,done:!t.done}:t));
  const remove = id => onChange(tasks.filter(t=>t.id!==id));

  return (
    <div>
      <div style={{...S.card,marginBottom:16,fontSize:12,color:"#475569",borderColor:"#1e3a5f"}}>
        ✅ <strong style={{color:"#94a3b8"}}>פעולות המשך</strong> — הוסף משימות לניהול הליד
      </div>

      {!adding && (
        <button onClick={()=>setAdding(true)}
          style={{...S.btn,background:"#f59e0b20",color:"#f59e0b",border:"1px solid #f59e0b44",width:"100%",marginBottom:14,textAlign:"center"}}>
          + הוסף משימה
        </button>
      )}

      {adding && (
        <div style={{...S.card,marginBottom:14,borderColor:"#f59e0b33"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
            <div>
              <label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>סוג</label>
              <select value={form.type} onChange={e=>setF("type",e.target.value)} style={S.input}>
                {TASK_TYPES.map(t=><option key={t.key} value={t.key}>{t.icon} {t.label}</option>)}
              </select>
            </div>
            <div>
              <label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>תאריך</label>
              <input type="date" value={form.dueDate} onChange={e=>setF("dueDate",e.target.value)} style={S.input}/>
            </div>
            <div>
              <label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>שעה</label>
              <input type="time" value={form.time} onChange={e=>setF("time",e.target.value)} style={S.input}/>
            </div>
          </div>
          <div style={{marginBottom:10}}>
            <label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>תיאור *</label>
            <input value={form.title} onChange={e=>setF("title",e.target.value)} placeholder="מה צריך לעשות?" style={S.input}/>
          </div>
          <div style={{marginBottom:10}}>
            <label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>הערות</label>
            <textarea value={form.notes} onChange={e=>setF("notes",e.target.value)} rows={2} style={{...S.input,resize:"vertical"}} placeholder="פרטים נוספים..."/>
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <button onClick={()=>setAdding(false)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12}}>ביטול</button>
            <button onClick={add} disabled={!form.title.trim()}
              style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,fontSize:12,opacity:form.title.trim()?1:.4}}>
              הוסף
            </button>
          </div>
        </div>
      )}

      {tasks.length===0 && !adding && (
        <div style={{textAlign:"center",color:"#334155",padding:"20px 0",fontSize:13}}>אין משימות — לחץ להוסיף</div>
      )}

      {tasks.map(t=>{
        const tt = taskType(t.type);
        const od = !t.done && isOverdue(t.dueDate);
        return (
          <div key={t.id} className="task-drag" draggable
            style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:8,
              border:`1px solid ${od?"#7f1d1d44":"#1e293b"}`,marginBottom:6,background:od?"#7f1d1d11":"#0d1520"}}>
            <input type="checkbox" checked={t.done} onChange={()=>toggle(t.id)} style={{accentColor:"#22c55e",cursor:"pointer"}}/>
            <span style={{fontSize:14}}>{tt.icon}</span>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:13,color:t.done?"#475569":"#e2e8f0",textDecoration:t.done?"line-through":"none"}}>{t.title}</div>
              {(t.dueDate||t.time) && (
                <div style={{fontSize:11,color:od?"#ef4444":"#475569"}}>
                  {t.dueDate&&<span>📅 {fmtDate(t.dueDate)}</span>}
                  {t.time&&<span style={{marginRight:6}}>⏰ {t.time}</span>}
                </div>
              )}
            </div>
            <button onClick={()=>remove(t.id)} style={{background:"none",border:"none",color:"#334155",cursor:"pointer",fontSize:16}}>×</button>
          </div>
        );
      })}
    </div>
  );
}

/* ══════════════════════════════════
   CRITERION ROW
══════════════════════════════════ */
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

/* ══════════════════════════════════
   TASK DETAIL MODAL
══════════════════════════════════ */
function TaskDetailModal({ task, lead, onClose, onToggle, onSave, onOpenLead }) {
  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({...task});
  const tt = taskType(task.type);
  const tier = getTier(calcScore(lead?.scores||{}));
  const set = (k,v) => setForm(p=>({...p,[k]:v}));

  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:"fixed",inset:0,background:"#000c",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:`1px solid ${tt.color}44`,borderRadius:16,width:"100%",maxWidth:500,overflow:"hidden",boxShadow:`0 0 32px ${tt.color}18`}}>
        <div style={{padding:"16px 20px",background:"#0a0f1a",borderBottom:"1px solid #1e293b",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{fontSize:28,lineHeight:1}}>{tt.icon}</div>
            <div>
              <div style={{fontSize:11,color:tt.color,fontWeight:700,letterSpacing:1,marginBottom:2}}>{tt.label.toUpperCase()}</div>
              <div style={{fontSize:16,fontWeight:700,color:"#e2e8f0"}}>{task.title}</div>
            </div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:20,cursor:"pointer"}}>×</button>
        </div>
        <div style={{padding:20,display:"flex",flexDirection:"column",gap:14}}>
          {lead && (
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:"#070d14",borderRadius:10,border:`1px solid ${tier.accent}33`}}>
              <div>
                <div style={{fontSize:12,color:"#475569",marginBottom:2}}>ליד</div>
                <div style={{fontSize:15,fontWeight:700,color:"#e2e8f0"}}>{lead.company}</div>
                {lead.contact&&<div style={{fontSize:11,color:"#64748b"}}>{lead.contact}</div>}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <div style={{textAlign:"center"}}>
                  <div style={{fontSize:22,fontWeight:900,color:tier.accent,fontFamily:"monospace"}}>{calcScore(lead.scores)}</div>
                  <div style={{fontSize:10,color:tier.accent}}>{tier.emoji} {tier.label}</div>
                </div>
                <button onClick={()=>{onOpenLead(lead.id);onClose();}}
                  style={{...S.btn,background:`${tier.accent}18`,color:tier.accent,border:`1px solid ${tier.accent}44`,fontSize:12,padding:"6px 12px"}}>
                  פתח ליד →
                </button>
              </div>
            </div>
          )}
          {!editing ? (
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <div style={{display:"flex",gap:12}}>
                {task.dueDate&&<div style={{flex:1,padding:"8px 12px",background:"#070d14",borderRadius:8,border:"1px solid #1e293b"}}><div style={{fontSize:10,color:"#475569",marginBottom:2}}>📅 תאריך</div><div style={{fontSize:14,color:"#e2e8f0"}}>{fmtDate(task.dueDate)}</div></div>}
                {task.time&&<div style={{flex:1,padding:"8px 12px",background:"#070d14",borderRadius:8,border:"1px solid #1e293b"}}><div style={{fontSize:10,color:"#475569",marginBottom:2}}>⏰ שעה</div><div style={{fontSize:14,color:"#e2e8f0"}}>{task.time}</div></div>}
              </div>
              {task.notes&&<div style={{padding:"10px 12px",background:"#070d14",borderRadius:8,border:"1px solid #1e293b",fontSize:13,color:"#94a3b8",lineHeight:1.6}}>{task.notes}</div>}
              <div style={{padding:"6px 12px",background:task.done?"#22c55e18":"#f59e0b18",borderRadius:8,border:`1px solid ${task.done?"#22c55e44":"#f59e0b44"}`,fontSize:12,color:task.done?"#22c55e":"#f59e0b",textAlign:"center"}}>{task.done?"✅ הושלמה":"⏳ פתוחה"}</div>
            </div>
          ) : (
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>תיאור</label><input value={form.title} onChange={e=>set("title",e.target.value)} style={S.input}/></div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>סוג</label><select value={form.type} onChange={e=>set("type",e.target.value)} style={S.input}>{TASK_TYPES.map(t=><option key={t.key} value={t.key}>{t.icon} {t.label}</option>)}</select></div>
                <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>תאריך</label><input type="date" value={form.dueDate||""} onChange={e=>set("dueDate",e.target.value)} style={S.input}/></div>
                <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>שעה</label><input type="time" value={form.time||""} onChange={e=>set("time",e.target.value)} style={S.input}/></div>
              </div>
              <div><label style={{fontSize:11,color:"#64748b",display:"block",marginBottom:4}}>הערות</label><textarea value={form.notes||""} onChange={e=>set("notes",e.target.value)} rows={2} style={{...S.input,resize:"vertical"}}/></div>
            </div>
          )}
        </div>
        <div style={{padding:"12px 20px",borderTop:"1px solid #1e293b",background:"#0a0f1a",display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={onToggle} style={{...S.btn,background:task.done?"#1e293b":"#22c55e18",color:task.done?"#94a3b8":"#22c55e",border:`1px solid ${task.done?"#1e293b":"#22c55e44"}`,fontSize:12,padding:"6px 10px"}}>{task.done?"↩ פתח מחדש":"✓ סמן הושלם"}</button>
          <button onClick={()=>downloadICS(task,lead)} style={{...S.btn,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",fontSize:12,padding:"6px 10px"}}>📅 ICS</button>
          {task.dueDate&&<a href={outlookUrl(task,lead)} target="_blank" rel="noreferrer" style={{...S.btn,background:"#0078d420",color:"#0078d4",border:"1px solid #0078d444",fontSize:12,padding:"6px 10px",textDecoration:"none",display:"inline-flex",alignItems:"center"}}>📆 Outlook</a>}
          <div style={{marginRight:"auto",display:"flex",gap:8}}>
            {editing
              ? <><button onClick={()=>setEditing(false)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,padding:"6px 10px"}}>ביטול</button><button onClick={()=>{onSave(form);setEditing(false);}} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,fontSize:12,padding:"6px 12px"}}>שמור</button></>
              : <button onClick={()=>setEditing(true)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,padding:"6px 10px"}}>✏️ ערוך</button>}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   EVENT CHIP — desktop drag + mobile long-press
══════════════════════════════════ */
function EventChip({ task, onClick, onDragStart, compact, onMoveRequest }) {
  const tt = taskType(task.type);
  const overdue = !task.done && isOverdue(task.dueDate);
  const color = task.done?"#334155":overdue?"#ef4444":tt.color;
  const longPressRef = useRef(null);

  // Mobile: long-press fires onMoveRequest
  const startLongPress = (e) => {
    longPressRef.current = setTimeout(()=>{ onMoveRequest&&onMoveRequest(task); }, 600);
  };
  const cancelLongPress = () => clearTimeout(longPressRef.current);

  return (
    <div className="ev" draggable
      onDragStart={e=>{
        e.dataTransfer.effectAllowed="move";
        e.dataTransfer.setData("taskId", task.id);
        e.dataTransfer.setData("leadId", task.leadId||"");
        onDragStart&&onDragStart(e);
      }}
      onTouchStart={startLongPress}
      onTouchEnd={cancelLongPress}
      onTouchMove={cancelLongPress}
      onClick={e=>{e.stopPropagation();onClick&&onClick();}}
      style={{display:"flex",alignItems:"center",gap:4,padding:compact?"3px 6px":"5px 8px",
        borderRadius:6,background:`${color}20`,border:`1px solid ${color}44`,
        cursor:"pointer",userSelect:"none",transition:"filter .1s",
        opacity:task.done?.6:1,overflow:"hidden",marginBottom:2,minHeight:compact?22:28}}>
      <span style={{fontSize:compact?11:12,flexShrink:0}}>{tt.icon}</span>
      <span style={{fontSize:compact?10:11,color:task.done?"#475569":"#cbd5e1",
        textDecoration:task.done?"line-through":"none",
        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>
        {task.time&&!compact&&<span style={{color,fontFamily:"monospace",marginLeft:4,fontSize:10}}>{task.time} </span>}
        {task.title}
      </span>
      {task.leadCompany&&!compact&&<span style={{fontSize:9,color:color,flexShrink:0,marginRight:2}}>·{task.leadCompany}</span>}
    </div>
  );
}

/* ══════════════════════════════════
   WEEKLY GRID — responsive (7/3/1 days)
══════════════════════════════════ */
function WeeklyGrid({ days, allTasks, onEventClick, onDropTask, onSlotClick, onMoveRequest }) {
  const [dragOver, setDragOver] = useState(null);
  const { isMobile, isTablet } = useBreakpoint();
  const today = new Date();
  const isToday = d => d.toDateString()===today.toDateString();

  // Responsive: mobile=3 days around today, tablet=5, desktop=7
  const visibleDays = isMobile
    ? days.filter((_,i)=>i>=Math.max(0,days.findIndex(d=>isToday(d))-1)&&i<=Math.min(6,days.findIndex(d=>isToday(d))+1)).slice(0,3)
    : isTablet ? days.slice(0,5) : days;

  const HOUR_H = isMobile?48:56;
  const TIME_W = isMobile?32:44;

  const tasksForSlot = (dayIso, hour) => allTasks.filter(t=>{
    if(!t.dueDate) return false;
    return isoDate(new Date(t.dueDate))===dayIso && (t.time?parseInt(t.time):null)===hour;
  }).sort((a,b)=>(a.time||"").localeCompare(b.time||""));

  const untimedForDay = dayIso => allTasks.filter(t=>isoDate(new Date(t.dueDate||"1970"))===dayIso&&!t.time&&t.dueDate);

  const handleDrop = (e, dayIso, hour) => {
    e.preventDefault();
    const taskId = e.dataTransfer.getData("taskId");
    const leadId = e.dataTransfer.getData("leadId");
    if(taskId) onDropTask({leadId,taskId},dayIso,hour);
    setDragOver(null);
  };

  const dayIdx = d => days.indexOf(d);

  return (
    <div style={{display:"flex",flexDirection:"column",flex:1,overflow:"hidden"}}>
      {/* Day headers */}
      <div style={{display:"flex",background:"#0a0f1a",borderBottom:"2px solid #1e293b",flexShrink:0}}>
        <div style={{width:TIME_W,flexShrink:0}}/>
        {visibleDays.map((day,i)=>(
          <div key={i} style={{flex:1,textAlign:"center",padding:isMobile?"6px 2px":"8px 0",borderRight:"1px solid #1e293b"}}>
            <div style={{fontSize:isMobile?10:11,color:isToday(day)?"#f59e0b":"#475569",fontWeight:isToday(day)?700:400}}>
              {DAY_SH[dayIdx(day)]}
            </div>
            <div style={{fontSize:isMobile?16:18,fontWeight:900,fontFamily:"monospace",color:isToday(day)?"#f59e0b":"#94a3b8"}}>
              {day.getDate()}
            </div>
          </div>
        ))}
      </div>
      {/* All-day row */}
      <div style={{display:"flex",background:"#070d14",borderBottom:"1px solid #1e293b",flexShrink:0,minHeight:28}}>
        <div style={{width:TIME_W,fontSize:8,color:"#334155",padding:"4px 0",textAlign:"center",flexShrink:0}}>כל<br/>היום</div>
        {visibleDays.map((day,i)=>{
          const dayIso=isoDate(day);
          return (
            <div key={i} style={{flex:1,borderRight:"1px solid #1e293b",padding:"2px 2px"}}
              onDragOver={e=>e.preventDefault()} onDrop={e=>handleDrop(e,dayIso,null)}>
              {untimedForDay(dayIso).map(t=>(
                <EventChip key={t.id} task={t} compact onClick={()=>onEventClick(t)} onMoveRequest={onMoveRequest}/>
              ))}
            </div>
          );
        })}
      </div>
      {/* Time slots */}
      <div style={{flex:1,overflowY:"auto"}}>
        {HOURS.map(hour=>(
          <div key={hour} style={{display:"flex",height:HOUR_H,borderBottom:"1px solid #0f1923"}}>
            <div style={{width:TIME_W,flexShrink:0,padding:"2px 4px 0 0",textAlign:"left",fontSize:9,color:"#334155",fontFamily:"monospace",userSelect:"none",lineHeight:1.2}}>
              {String(hour).padStart(2,"0")}
              {!isMobile&&":00"}
            </div>
            {visibleDays.map((day,i)=>{
              const dayIso=isoDate(day);
              const isOver=dragOver?.dayIso===dayIso&&dragOver?.hour===hour;
              return (
                <div key={i} className={isOver?"slot-drop":""}
                  style={{flex:1,borderRight:"1px solid #1e293b",padding:"2px 2px",
                    background:isToday(day)?"#f59e0b04":"transparent",transition:"background .1s",cursor:"pointer"}}
                  onClick={()=>onSlotClick(dayIso,hour)}
                  onDragOver={e=>{e.preventDefault();setDragOver({dayIso,hour});}}
                  onDragLeave={()=>setDragOver(null)}
                  onDrop={e=>handleDrop(e,dayIso,hour)}>
                  {tasksForSlot(dayIso,hour).map(t=>(
                    <EventChip key={t.id} task={t}
                      onClick={()=>onEventClick(t)}
                      onMoveRequest={onMoveRequest}
                      compact={isMobile}/>
                  ))}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   MONTHLY GRID
══════════════════════════════════ */
function MonthlyGrid({ year, month, allTasks, onEventClick, onDropTask, onSlotClick }) {
  const [dragOver, setDragOver] = useState(null);
  const today = new Date();
  const firstDay = new Date(year,month,1);
  const startSun = new Date(firstDay); startSun.setDate(1-firstDay.getDay());
  const cells = Array.from({length:42},(_,i)=>{const d=new Date(startSun);d.setDate(startSun.getDate()+i);return d;});

  const tasksForDay = dayIso => allTasks.filter(t=>t.dueDate&&isoDate(new Date(t.dueDate))===dayIso).sort((a,b)=>(a.time||"25:00").localeCompare(b.time||"25:00"));

  const handleDrop = (e, dayIso) => {
    e.preventDefault();
    const taskId = e.dataTransfer.getData("taskId");
    const leadId = e.dataTransfer.getData("leadId");
    if(taskId) onDropTask({leadId,taskId},dayIso,null);
    setDragOver(null);
  };

  return (
    <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",background:"#0a0f1a",borderBottom:"2px solid #1e293b",flexShrink:0}}>
        {DAY_SH.map(d=><div key={d} style={{textAlign:"center",padding:"8px 0",fontSize:11,color:"#475569",fontWeight:600}}>{d}</div>)}
      </div>
      <div style={{flex:1,display:"grid",gridTemplateColumns:"repeat(7,1fr)",gridTemplateRows:"repeat(6,1fr)",overflow:"hidden"}}>
        {cells.map((day,i)=>{
          const dayIso=isoDate(day);
          const inMonth=day.getMonth()===month;
          const isTd=day.toDateString()===today.toDateString();
          const isOver=dragOver===dayIso;
          const tasks=tasksForDay(dayIso);
          return (
            <div key={i} className={isOver?"slot-drop":""}
              style={{border:"1px solid #1e293b",padding:"4px 5px",overflow:"hidden",background:isTd?"#f59e0b08":"transparent",opacity:inMonth?1:.4,cursor:"pointer",transition:"background .1s"}}
              onClick={()=>inMonth&&onSlotClick(dayIso,9)}
              onDragOver={e=>e.preventDefault()} onDragLeave={()=>setDragOver(null)}
              onDrop={e=>handleDrop(e,dayIso)}>
              <div style={{fontSize:12,fontWeight:isTd?900:inMonth?500:400,color:isTd?"#f59e0b":inMonth?"#94a3b8":"#334155",marginBottom:3,lineHeight:1}}>{day.getDate()}</div>
              {tasks.slice(0,3).map(t=><EventChip key={t.id} task={t} compact onClick={()=>onEventClick(t)}/>)}
              {tasks.length>3&&<div style={{fontSize:9,color:"#475569",marginTop:2}}>+{tasks.length-3} עוד</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   MOBILE RESCHEDULE MODAL
══════════════════════════════════ */
function RescheduleModal({ task, lead, onReschedule, onClose }) {
  const [date, setDate] = useState(task.dueDate||isoDate(new Date()));
  const [time, setTime] = useState(task.time||"09:00");
  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:"fixed",inset:0,background:"#000c",zIndex:400,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
      <div style={{background:"#0d1520",border:"1px solid #f59e0b44",borderRadius:"16px 16px 0 0",
        width:"100%",maxWidth:500,padding:20,animation:"slideUp .25s ease"}}>
        <div style={{width:40,height:4,background:"#1e293b",borderRadius:2,margin:"0 auto 16px"}}/>
        <div style={{fontSize:15,fontWeight:700,color:"#e2e8f0",marginBottom:4}}>⏱ שינוי זמן</div>
        <div style={{fontSize:12,color:"#475569",marginBottom:16}}>{task.title}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>תאריך</label>
            <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={S.input}/></div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>שעה</label>
            <input type="time" value={time} onChange={e=>setTime(e.target.value)} style={S.input}/></div>
        </div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={onClose} style={{...S.btn,flex:1,background:"#1e293b",color:"#94a3b8"}}>ביטול</button>
          <button onClick={()=>{onReschedule(task.leadId,task.id,{dueDate:date,time});onClose();}}
            style={{...S.btn,flex:2,background:"#f59e0b",color:"#000",fontWeight:700}}>עדכן זמן</button>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   CALENDAR VIEW
══════════════════════════════════ */
function CalendarView({ leads, onTaskUpdate, onOpenLead, onAddTask }) {
  const [calView, setCalView] = useState("week");
  const [weekOffset, setWeekOffset] = useState(0);
  const [monthOffset, setMonthOffset] = useState(0);
  const [taskDetail, setTaskDetail] = useState(null);
  const [reschedule, setReschedule] = useState(null); // mobile reschedule
  const [showUndated, setShowUndated] = useState(false); // mobile undated drawer
  const { isMobile } = useBreakpoint();
  const today = new Date();

  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate()-today.getDay()+weekOffset*7);
  const days = Array.from({length:7},(_,i)=>{const d=new Date(startOfWeek);d.setDate(startOfWeek.getDate()+i);return d;});
  const mDate = new Date(today.getFullYear(),today.getMonth()+monthOffset,1);

  const allTasks = leads.flatMap(l=>(l.tasks||[]).map(t=>({...t,leadId:l.id,leadCompany:l.company,leadTier:getTier(calcScore(l.scores))})));
  const overdueCount = allTasks.filter(t=>!t.done&&isOverdue(t.dueDate)).length;
  const openCount = allTasks.filter(t=>!t.done).length;
  const undated = allTasks.filter(t=>!t.done&&!t.dueDate);

  const handleDropTask = ({leadId,taskId},dayIso,hour) => {
    const newTime = hour!=null?`${String(hour).padStart(2,"0")}:00`:undefined;
    onTaskUpdate(leadId,taskId,{dueDate:dayIso,...(newTime?{time:newTime}:{time:undefined})});
  };

  const handleMoveRequest = task => {
    const lead = leads.find(l=>l.id===task.leadId);
    setReschedule({task, lead});
  };

  return (
    <div style={{display:"flex",height:"100%",overflow:"hidden"}}>
      {/* Desktop sidebar — undated tasks */}
      {!isMobile && (
        <div className="cal-sidebar" style={{width:200,flexShrink:0,borderLeft:"1px solid #1e293b",background:"#0a0f1a",display:"flex",flexDirection:"column",overflow:"hidden"}}>
          <div style={{padding:"10px 12px",borderBottom:"1px solid #1e293b",fontSize:12,fontWeight:700,color:"#64748b"}}>
            📌 ללא תאריך {undated.length>0&&<span style={{background:"#f59e0b",color:"#000",borderRadius:8,padding:"1px 6px",fontSize:10,marginRight:4}}>{undated.length}</span>}
          </div>
          <div style={{flex:1,overflowY:"auto",padding:"6px 5px"}}>
            <div style={{fontSize:10,color:"#334155",padding:"4px 6px 8px",lineHeight:1.4}}>💡 גרור לתוך הלוח לתזמון</div>
            {undated.length===0&&<div style={{fontSize:11,color:"#334155",textAlign:"center",padding:"12px 0"}}>הכל מתוזמן ✓</div>}
            {undated.map(t=>(
              <div key={t.id} className="task-drag" draggable
                onDragStart={e=>{e.dataTransfer.effectAllowed="move";e.dataTransfer.setData("taskId",t.id);e.dataTransfer.setData("leadId",t.leadId||"");}}
                onClick={()=>setTaskDetail({task:t,lead:leads.find(l=>l.id===t.leadId)})}
                style={{marginBottom:4,cursor:"grab"}}>
                <EventChip task={t} compact onMoveRequest={handleMoveRequest}/>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Main calendar */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        {/* Toolbar */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
          padding:isMobile?"6px 10px":"8px 16px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",flexShrink:0,gap:6,flexWrap:"wrap"}}>
          <div style={{display:"flex",gap:4,alignItems:"center"}}>
            {[{k:"week",l:"שבועי"},{k:"month",l:"חודשי"}].map(v=>(
              <button key={v.k} onClick={()=>setCalView(v.k)}
                style={{...S.btn,padding:"4px 8px",fontSize:11,background:calView===v.k?"#f59e0b20":"#1e293b",color:calView===v.k?"#f59e0b":"#64748b",border:`1px solid ${calView===v.k?"#f59e0b44":"transparent"}`}}>{v.l}</button>
            ))}
            <button onClick={()=>calView==="week"?setWeekOffset(w=>w-1):setMonthOffset(m=>m-1)} style={{...S.btn,background:"#1e293b",color:"#64748b",padding:"4px 7px",fontSize:12}}>←</button>
            <button onClick={()=>calView==="week"?setWeekOffset(0):setMonthOffset(0)} style={{...S.btn,background:"#1e293b",color:(calView==="week"?weekOffset:monthOffset)===0?"#f59e0b":"#64748b",padding:"4px 7px",fontSize:11}}>היום</button>
            <button onClick={()=>calView==="week"?setWeekOffset(w=>w+1):setMonthOffset(m=>m+1)} style={{...S.btn,background:"#1e293b",color:"#64748b",padding:"4px 7px",fontSize:12}}>→</button>
          </div>
          <div style={{fontSize:isMobile?11:13,fontWeight:700,color:"#94a3b8"}}>
            {calView==="week"?`${fmtDate(days[0])} – ${fmtDate(days[6])}`:`${MON_HE[mDate.getMonth()]} ${mDate.getFullYear()}`}
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {isMobile && undated.length>0 && (
              <button onClick={()=>setShowUndated(v=>!v)}
                style={{...S.btn,background:"#f59e0b20",color:"#f59e0b",border:"1px solid #f59e0b44",padding:"4px 8px",fontSize:11}}>
                📌 {undated.length}
              </button>
            )}
            {overdueCount>0&&<span style={{fontSize:11,color:"#ef4444"}}>⚠️{overdueCount}</span>}
            <span style={{fontSize:11,color:"#64748b"}}>{openCount} פ׳</span>
            <button onClick={()=>onAddTask(isoDate(today),9)} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,padding:"4px 10px",fontSize:12}}>+</button>
          </div>
        </div>

        {/* Mobile undated drawer */}
        {isMobile && showUndated && (
          <div style={{background:"#0a0f1a",borderBottom:"1px solid #1e293b",padding:"8px 10px",maxHeight:120,overflowX:"auto"}}>
            <div style={{display:"flex",gap:6,minWidth:"max-content"}}>
              {undated.map(t=>(
                <div key={t.id} style={{flexShrink:0,maxWidth:140}}>
                  <EventChip task={t} compact
                    onClick={()=>setTaskDetail({task:t,lead:leads.find(l=>l.id===t.leadId)})}
                    onMoveRequest={handleMoveRequest}/>
                </div>
              ))}
            </div>
            <div style={{fontSize:10,color:"#334155",marginTop:4}}>לחץ על משימה לשינוי זמן</div>
          </div>
        )}

        {calView==="week"
          ? <WeeklyGrid days={days} allTasks={allTasks}
              onEventClick={t=>setTaskDetail({task:t,lead:leads.find(l=>l.id===t.leadId)})}
              onDropTask={handleDropTask}
              onSlotClick={(d,h)=>onAddTask(d,h)}
              onMoveRequest={handleMoveRequest}/>
          : <MonthlyGrid year={mDate.getFullYear()} month={mDate.getMonth()} allTasks={allTasks}
              onEventClick={t=>setTaskDetail({task:t,lead:leads.find(l=>l.id===t.leadId)})}
              onDropTask={handleDropTask}
              onSlotClick={(d,h)=>onAddTask(d,h)}/>}
      </div>

      {taskDetail&&(
        <TaskDetailModal
          task={taskDetail.task} lead={taskDetail.lead}
          onClose={()=>setTaskDetail(null)}
          onToggle={()=>{onTaskUpdate(taskDetail.task.leadId,taskDetail.task.id,{done:!taskDetail.task.done});setTaskDetail(null);}}
          onSave={form=>{onTaskUpdate(taskDetail.task.leadId,taskDetail.task.id,form);setTaskDetail(null);}}
          onOpenLead={id=>{onOpenLead(id);}}
        />
      )}
      {reschedule&&(
        <RescheduleModal
          task={reschedule.task} lead={reschedule.lead}
          onReschedule={onTaskUpdate}
          onClose={()=>setReschedule(null)}
        />
      )}
    </div>
  );
}

/* ══════════════════════════════════
   QUICK ADD TASK MODAL
══════════════════════════════════ */
function QuickTaskModal({ defaultDate, defaultHour, leads, onAdd, onClose }) {
  const [leadId, setLeadId] = useState(leads[0]?.id||"");
  const [form, setForm] = useState({type:"call",title:"",dueDate:defaultDate||"",time:defaultHour!=null?`${String(defaultHour).padStart(2,"0")}:00`:"",notes:""});
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const submit=()=>{if(!form.title||!leadId)return;onAdd(leadId,{...form,id:uid(),done:false,createdAt:new Date().toISOString()});onClose();};
  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:"fixed",inset:0,background:"#000c",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:"1px solid #f59e0b44",borderRadius:14,width:"100%",maxWidth:460,overflow:"hidden"}}>
        <div style={{padding:"14px 20px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:15,fontWeight:700,color:"#e2e8f0"}}>+ משימה חדשה</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:20,cursor:"pointer"}}>×</button>
        </div>
        <div style={{padding:20,display:"flex",flexDirection:"column",gap:12}}>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>ליד *</label><select value={leadId} onChange={e=>setLeadId(e.target.value)} style={S.input}>{leads.map(l=><option key={l.id} value={l.id}>{l.company}</option>)}</select></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
            <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>סוג</label><select value={form.type} onChange={e=>set("type",e.target.value)} style={S.input}>{TASK_TYPES.map(t=><option key={t.key} value={t.key}>{t.icon} {t.label}</option>)}</select></div>
            <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>תאריך</label><input type="date" value={form.dueDate} onChange={e=>set("dueDate",e.target.value)} style={S.input}/></div>
            <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>שעה</label><input type="time" value={form.time} onChange={e=>set("time",e.target.value)} style={S.input}/></div>
          </div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>תיאור *</label><input value={form.title} onChange={e=>set("title",e.target.value)} placeholder="מה צריך לעשות?" style={S.input}/></div>
          <div><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>הערות</label><textarea value={form.notes} onChange={e=>set("notes",e.target.value)} rows={2} style={{...S.input,resize:"vertical"}}/></div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <button onClick={onClose} style={{...S.btn,background:"#1e293b",color:"#94a3b8"}}>ביטול</button>
            <button onClick={submit} disabled={!form.title||!leadId} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,opacity:form.title&&leadId?1:.4}}>הוסף</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   IMPORT MODAL
══════════════════════════════════ */
function ImportModal({ onImport, onClose }) {
  const [step,setStep]=useState("upload");
  const [rows,setRows]=useState([]);
  const [selected,setSelected]=useState(new Set());
  const [err,setErr]=useState("");
  const [dragging,setDragging]=useState(false);
  const inputRef=useRef();

  const parseFile=file=>{
    setErr("");
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const json=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
        if(!json.length){setErr("הגיליון ריק");return;}
        setRows(json);setSelected(new Set(json.map((_,i)=>i)));setStep("preview");
      }catch(ex){setErr("שגיאה: "+ex.message);}
    };
    reader.readAsArrayBuffer(file);
  };

  const doImport=()=>{
    const TITLE=Object.keys(rows[0]).find(k=>/title/i.test(k));
    const LABEL=Object.keys(rows[0]).find(k=>/label/i.test(k));
    const DATE=Object.keys(rows[0]).find(k=>/created/i.test(k));
    const NACT=Object.keys(rows[0]).find(k=>/next.activity/i.test(k));
    const leads=[...selected].map(i=>{
      const r=rows[i];
      const title=r[TITLE]||"(ללא שם)";
      const tasks=r[NACT]?[{id:uid(),type:"follow",title:`פ.ה: ${r[NACT]}`,dueDate:"",time:"",notes:"מ-Pipedrive",done:false,createdAt:new Date().toISOString()}]:[];
      return{id:uid(),company:title.replace(/ lead$/i,"").replace(/ MES\b/i,"").trim(),name:title,contact:"",phone:"",email:"",role:"",label:r[LABEL]||"",tasks,notes:"",scores:{},source:"pipedrive",importedAt:r[DATE]||"",createdAt:new Date().toISOString()};
    });
    onImport(leads);onClose();
  };

  const TITLE_COL=rows[0]?Object.keys(rows[0]).find(k=>/title/i.test(k)):null;
  const LABEL_COL=rows[0]?Object.keys(rows[0]).find(k=>/label/i.test(k)):null;
  const DATE_COL=rows[0]?Object.keys(rows[0]).find(k=>/created/i.test(k)):null;

  return(
    <div onClick={e=>e.target===e.currentTarget&&onClose()}
      style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:"1px solid #1e293b",borderRadius:16,width:"100%",maxWidth:700,maxHeight:"88vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:"18px 24px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontSize:17,fontWeight:700,color:"#e2e8f0"}}>⬆️ ייבוא מ-Pipedrive</div><div style={{fontSize:11,color:"#475569",marginTop:2}}>XLSX / CSV</div></div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:24}}>
          {step==="upload"&&(
            <>
              <div style={{...S.card,marginBottom:20,borderColor:"#1e3a5f",fontSize:12,color:"#94a3b8",lineHeight:2}}>
                <strong style={{color:"#60a5fa"}}>ייצוא מ-Pipedrive:</strong> Leads Inbox → ⋮ → Export leads → XLSX
              </div>
              <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)}
                onDrop={e=>{e.preventDefault();setDragging(false);parseFile(e.dataTransfer.files[0]);}}
                onClick={()=>inputRef.current.click()}
                style={{border:`2px dashed ${dragging?"#f59e0b":"#1e293b"}`,borderRadius:12,padding:40,textAlign:"center",cursor:"pointer",background:dragging?"#f59e0b08":"#0a0f1a",transition:"all .2s"}}>
                <div style={{fontSize:48,marginBottom:12}}>📂</div>
                <div style={{fontSize:15,color:"#94a3b8",marginBottom:6,fontWeight:600}}>גרור קובץ XLSX / CSV</div>
                <div style={{fontSize:12,color:"#475569"}}>או לחץ לבחירת קובץ</div>
                <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>parseFile(e.target.files[0])}/>
              </div>
              {err&&<div style={{marginTop:12,color:"#ef4444",fontSize:13,textAlign:"center"}}>{err}</div>}
            </>
          )}
          {step==="preview"&&(
            <>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div style={{fontSize:14,color:"#94a3b8"}}>נמצאו <strong style={{color:"#f59e0b"}}>{rows.length}</strong> · נבחרו <strong style={{color:"#22c55e"}}>{selected.size}</strong></div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>setSelected(new Set(rows.map((_,i)=>i)))} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:11,padding:"4px 10px"}}>בחר הכל</button>
                  <button onClick={()=>setSelected(new Set())} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:11,padding:"4px 10px"}}>נקה</button>
                </div>
              </div>
              {TITLE_COL&&<div style={{...S.card,marginBottom:12,fontSize:11,color:"#475569"}}>
                זיהוי: {TITLE_COL&&<span style={{color:"#f59e0b"}}>{TITLE_COL} </span>}{LABEL_COL&&<span style={{color:"#f59e0b"}}>· {LABEL_COL} </span>}{DATE_COL&&<span style={{color:"#f59e0b"}}>· {DATE_COL}</span>}
              </div>}
              <div style={{maxHeight:380,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
                {rows.map((r,i)=>{
                  const title=TITLE_COL?r[TITLE_COL]:"(ללא שם)";const label=LABEL_COL?r[LABEL_COL]:"";const date=DATE_COL?r[DATE_COL]:"";const sel=selected.has(i);
                  return(<div key={i} onClick={()=>setSelected(p=>{const n=new Set(p);sel?n.delete(i):n.add(i);return n;})}
                    style={{display:"flex",alignItems:"center",gap:12,padding:"8px 12px",borderRadius:8,cursor:"pointer",background:sel?"#f59e0b08":"#0a0f1a",border:`1px solid ${sel?"#f59e0b44":"#1e293b"}`}}>
                    <input type="checkbox" checked={sel} onChange={()=>{}} style={{accentColor:"#f59e0b",cursor:"pointer",flexShrink:0}}/>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:13,color:"#e2e8f0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{title}</div>
                      <div style={{fontSize:11,color:"#475569"}}>{date&&String(date).slice(0,10)}{label&&` · 🏷 ${label}`}</div>
                    </div>
                  </div>);
                })}
              </div>
            </>
          )}
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

/* ══════════════════════════════════
   SETTINGS MODAL
══════════════════════════════════ */
function SettingsModal({token,onSave,onClose}){
  const [val,setVal]=useState(token||"");
  const [testing,setTesting]=useState(false);
  const [res,setRes]=useState(null);
  const test=async()=>{setTesting(true);setRes(null);try{const r=await PD.req(val,"/users/me");setRes({ok:true,name:r.data?.name});}catch(e){setRes({ok:false,err:e.message});}setTesting(false);};
  return(
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:"1px solid #1e293b",borderRadius:16,width:"100%",maxWidth:520,overflow:"hidden"}}>
        <div style={{padding:"18px 24px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:17,fontWeight:700,color:"#e2e8f0"}}>⚙️ חיבור ל-Pipedrive</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#475569",fontSize:22,cursor:"pointer"}}>×</button>
        </div>
        <div style={{padding:24}}>
          <div style={{...S.card,marginBottom:20,borderColor:"#1e3a5f",fontSize:12,color:"#94a3b8",lineHeight:1.8}}><strong style={{color:"#60a5fa"}}>איך מוצאים?</strong><br/>Pipedrive → שמך → Personal preferences → API</div>
          <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:6}}>API Token</label>
          <input value={val} onChange={e=>setVal(e.target.value)} type="password" placeholder="הדבק token..." style={{...S.input,marginBottom:12,letterSpacing:2}}/>
          {res&&<div style={{padding:"8px 12px",borderRadius:8,fontSize:12,marginBottom:12,background:res.ok?"#22c55e10":"#ef444410",color:res.ok?"#22c55e":"#ef4444",border:`1px solid ${res.ok?"#22c55e44":"#ef444444"}`}}>{res.ok?`✅ שלום ${res.name}`:`❌ ${res.err}`}</div>}
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <button onClick={test} disabled={!val||testing} style={{...S.btn,background:"#1e293b",color:"#94a3b8",opacity:val?1:.5}}>{testing?"בודק...":"🔌 בדוק חיבור"}</button>
            <button onClick={()=>onSave(val)} disabled={!val} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,opacity:val?1:.4}}>שמור Token</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   LEAD MODAL (4-step wizard)
══════════════════════════════════ */
const STEPS=[{key:"info",label:"פרטים",icon:"📋"},{key:"fit",label:"Fit",icon:"🏭"},{key:"bant",label:"BANT",icon:"💰"},{key:"tasks",label:"משימות",icon:"✅"}];

function LeadModal({lead:init,onSave,onDelete,onClose,apiToken}){
  const isNew=!init.id;
  const [lead,setLead]=useState({role:"",email:"",...init,scores:{...init.scores},tasks:[...(init.tasks||[])]});
  const [step,setStep]=useState(0);
  const [searching,setSearching]=useState(false);
  const [pdMatches,setPdMatches]=useState(null);
  const [pdStatus,setPdStatus]=useState("");
  const [confirmDel,setConfirmDel]=useState(false);
  const [showAI,setShowAI]=useState(false);
  const debRef=useRef();

  const set=(k,v)=>setLead(p=>({...p,[k]:v}));
  const setScore=(id,v)=>setLead(p=>({...p,scores:{...p.scores,[id]:v}}));
  const score=calcScore(lead.scores);
  const tier=getTier(score);
  const filled=CRITERIA.filter(c=>lead.scores[c.id]>0).length;
  const pendingTasks=(lead.tasks||[]).filter(t=>!t.done).length;

  const searchPD=useCallback((term)=>{
    if(!apiToken||!term||term.length<3){setPdMatches(null);return;}
    clearTimeout(debRef.current);
    debRef.current=setTimeout(async()=>{
      setSearching(true);
      try{const res=await PD.searchOrg(apiToken,term);setPdMatches((res.data?.items||[]).map(i=>i.item));}
      catch{setPdMatches([]);}
      setSearching(false);
    },600);
  },[apiToken]);

  const handleSave=async()=>{
    let updated={...lead};
    // Create or link org in Pipedrive
    if(apiToken&&lead.company){
      setPdStatus("creating");
      try{
        let orgId=lead.pipedriveOrgId;
        if(!orgId){
          const orgRes=await PD.createOrg(apiToken,lead.company);
          orgId=orgRes.data.id;
          updated.pipedriveOrgId=orgId;
          if(lead.contact){
            const pr=await PD.createPerson(apiToken,{name:lead.contact,org_id:orgId,phone:[{value:lead.phone||"",primary:true}],email:[{value:lead.email||"",primary:true}]});
            updated.pipedrivePersonId=pr.data.id;
          }
        }
        // Push qualification note to Pipedrive
        const noteContent=`🎯 SDR Qualification\nסיווג: ${tier.emoji} ${tier.label}\nFit: ${fitScore(lead.scores)}/60 | BANT: ${bantScore(lead.scores)}/40 | סה"כ: ${score}/100\n\n${CRITERIA.filter(c=>c.group==="bant").map(c=>`${c.label}: ${lead.scores?.[c.id]||0}`).join(" | ")}\n\nהערות SDR:\n${lead.notes||"—"}`;
        await PD.addNote(apiToken,{content:noteContent,org_id:orgId,...(updated.pipedrivePersonId?{person_id:updated.pipedrivePersonId}:{})});
        // Update org label
        await PD.updateOrg(apiToken,orgId,{label:tier.label});
        setPdStatus("done");
      }catch(ex){setPdStatus("warn: "+ex.message);}
    }
    onSave(updated);
  };

  const isLastStep = step===STEPS.length-1;

  return(
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"#000c",zIndex:100,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0d1520",border:`1px solid ${tier.accent}44`,borderRadius:16,width:"100%",maxWidth:800,maxHeight:"92vh",display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:`0 0 40px ${tier.accent}12`}}>
        {/* Header */}
        <div style={{padding:"14px 22px",borderBottom:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:17,fontWeight:700,color:"#e2e8f0"}}>{isNew?"➕ ליד חדש":`✏️ ${lead.company||"עריכת ליד"}`}</div>
            <div style={{fontSize:11,color:"#475569",marginTop:2}}>
              {filled}/{CRITERIA.length} קריטריונים · {pendingTasks} משימות ·
              {lead.pipedriveOrgId?<span style={{color:"#60a5fa"}}> 🔗 מקושר</span>:apiToken?<span style={{color:"#334155"}}> יצור ב-Pipe</span>:<span style={{color:"#7f1d1d"}}> ⚙️ חסר Token</span>}
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
            <button key={s.key} onClick={()=>setStep(i)}
              style={{flex:1,padding:"10px 0",background:"none",border:"none",cursor:"pointer",fontSize:13,fontWeight:step===i?700:400,fontFamily:"inherit",color:step===i?tier.accent:"#475569",borderBottom:`2px solid ${step===i?tier.accent:"transparent"}`,transition:"all .15s"}}>
              {s.icon} {s.label}{s.key==="tasks"&&pendingTasks>0&&<span style={{marginRight:4,background:"#f59e0b",color:"#000",borderRadius:8,padding:"1px 5px",fontSize:10,fontWeight:700}}>{pendingTasks}</span>}
            </button>
          ))}
        </div>
        {/* Body */}
        <div style={{flex:1,overflowY:"auto",padding:22}}>
          {/* STEP 0 — Info */}
          {step===0&&(
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                <div style={{gridColumn:"1/-1",position:"relative"}}>
                  <label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>שם חברה *{searching&&<span style={{color:"#f59e0b",marginRight:6,fontSize:11}}> ⏳ מחפש...</span>}{lead.pipedriveOrgId&&<span style={{color:"#60a5fa",marginRight:6,fontSize:11}}> 🔗 ID:{lead.pipedriveOrgId}</span>}</label>
                  <input value={lead.company||""} onChange={e=>{set("company",e.target.value);if(!lead.pipedriveOrgId)searchPD(e.target.value);}} placeholder="שם החברה" style={S.input}/>
                  {pdMatches&&pdMatches.length>0&&!lead.pipedriveOrgId&&(
                    <div style={{position:"absolute",top:"100%",right:0,left:0,zIndex:50,background:"#0d1520",border:"1px solid #f59e0b44",borderRadius:8,overflow:"hidden",boxShadow:"0 8px 24px #0008"}}>
                      <div style={{padding:"7px 12px",fontSize:11,color:"#64748b",borderBottom:"1px solid #1e293b"}}>🔍 נמצא ב-Pipedrive</div>
                      {pdMatches.map(m=>(
                        <div key={m.id} onClick={()=>{set("company",m.name);set("pipedriveOrgId",m.id);setPdMatches(null);}}
                          style={{padding:"10px 14px",cursor:"pointer",fontSize:13,color:"#e2e8f0",borderBottom:"1px solid #0f1923",display:"flex",justifyContent:"space-between"}}
                          onMouseOver={e=>e.currentTarget.style.background="#f59e0b0a"} onMouseOut={e=>e.currentTarget.style.background=""}>
                          🏢 {m.name}<span style={{fontSize:11,color:"#475569"}}>ID {m.id}</span>
                        </div>
                      ))}
                      <div onClick={()=>setPdMatches(null)} style={{padding:"8px 14px",cursor:"pointer",fontSize:12,color:"#64748b",textAlign:"center",background:"#0a0f1a"}}>× צור כחדש</div>
                    </div>
                  )}
                  {pdMatches&&pdMatches.length===0&&<div style={{fontSize:11,color:"#334155",marginTop:3}}>✓ לא קיים — ייווצר עם שמירה</div>}
                </div>
                {[{k:"contact",label:"איש קשר",ph:"ישראל ישראלי"},{k:"role",label:"תפקיד",ph:"מנהל ייצור"},{k:"phone",label:"טלפון",ph:"050-1234567"},{k:"email",label:"אימייל",ph:"contact@company.com"},{k:"label",label:"תווית Pipedrive",ph:"Hot / Warm / Cold"}].map(f=>(
                  <div key={f.k}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>{f.label}</label><input value={lead[f.k]||""} onChange={e=>set(f.k,e.target.value)} placeholder={f.ph} style={S.input}/></div>
                ))}
                <div style={{gridColumn:"1/-1"}}><label style={{fontSize:12,color:"#64748b",display:"block",marginBottom:4}}>הערות שיחה</label><textarea value={lead.notes||""} onChange={e=>set("notes",e.target.value)} rows={4} style={{...S.input,resize:"vertical"}} placeholder="סיכום השיחה, כאבים שעלו, שלב בתהליך..."/></div>
              </div>
              {pdStatus&&<div style={{marginTop:12,padding:"8px 12px",borderRadius:8,fontSize:12,background:pdStatus.includes("done")?"#22c55e10":pdStatus.includes("creating")?"#f59e0b10":"#ef444410",color:pdStatus.includes("done")?"#22c55e":pdStatus.includes("creating")?"#f59e0b":"#ef4444"}}>{pdStatus.includes("done")?"✅ מסונכרן ל-Pipedrive!":pdStatus.includes("creating")?"⏳ מסנכרן...":"⚠️ "+pdStatus}</div>}
            </div>
          )}
          {/* STEP 1 — Fit */}
          {step===1&&(<div><div style={{...S.card,marginBottom:16,fontSize:12,color:"#475569",borderColor:"#1e3a5f"}}>🏭 <strong style={{color:"#94a3b8"}}>Basic Fit</strong> — עד 60 נקודות · {fitScore(lead.scores)}/60</div>{CRITERIA.filter(c=>c.group==="fit").map(c=><CriterionRow key={c.id} c={c} value={lead.scores[c.id]} onChange={setScore}/>)}</div>)}
          {/* STEP 2 — BANT */}
          {step===2&&(<div>
            <div style={{...S.card,marginBottom:16,fontSize:12,color:"#475569",borderColor:"#1e3a5f"}}>💰 <strong style={{color:"#94a3b8"}}>BANT</strong> — עד 40 נקודות · {bantScore(lead.scores)}/40</div>
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
          {/* STEP 3 — Tasks */}
          {step===3&&(<TaskStep tasks={lead.tasks||[]} onChange={ts=>set("tasks",ts)}/>)}
        </div>
        {/* Footer */}
        <div style={{padding:"12px 22px",borderTop:"1px solid #1e293b",background:"#0a0f1a",display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <div style={{display:"flex",gap:8}}>
            {!isNew&&(
              confirmDel
                ? <><span style={{fontSize:12,color:"#fca5a5",lineHeight:"32px"}}>בטוח?</span><button onClick={()=>onDelete(lead.id)} style={{...S.btn,background:"#ef4444",color:"#fff",fontWeight:700,fontSize:12,padding:"5px 12px"}}>מחק ✓</button><button onClick={()=>setConfirmDel(false)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:12,padding:"5px 10px"}}>ביטול</button></>
                : <button onClick={()=>setConfirmDel(true)} style={{...S.btn,background:"#7f1d1d22",color:"#fca5a5",border:"1px solid #7f1d1d44",fontSize:12}}>🗑 מחק</button>
            )}
            {/* AI Email button — show on last step or always */}
            <button onClick={()=>setShowAI(true)}
              style={{...S.btn,background:"#8b5cf620",color:"#8b5cf6",border:"1px solid #8b5cf644",fontSize:12}}>
              🤖 מייל AI
            </button>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <div style={{display:"flex",gap:5}}>{STEPS.map((_,i)=><div key={i} onClick={()=>setStep(i)} style={{width:7,height:7,borderRadius:"50%",cursor:"pointer",background:i===step?tier.accent:"#1e293b",transition:"background .2s"}}/>)}</div>
            {step>0&&<button onClick={()=>setStep(s=>s-1)} style={{...S.btn,background:"#1e293b",color:"#94a3b8",fontSize:13}}>← קודם</button>}
            {!isLastStep
              ? <button onClick={()=>setStep(s=>s+1)} disabled={!lead.company?.trim()} style={{...S.btn,background:tier.accent,color:"#000",fontWeight:700,opacity:lead.company?.trim()?1:.4}}>הבא →</button>
              : <button onClick={handleSave} disabled={!lead.company?.trim()} style={{...S.btn,background:tier.accent,color:"#000",fontWeight:700,opacity:lead.company?.trim()?1:.4}}>
                  {isNew?"הוסף ✓":"שמור ✓"}{apiToken?" + Pipe":""}
                </button>}
          </div>
        </div>
      </div>
      {showAI&&<AIEmailModal lead={lead} onClose={()=>setShowAI(false)}/>}
    </div>
  );
}

/* ══════════════════════════════════
   LEAD CARD
══════════════════════════════════ */
function LeadCard({lead,onClick}){
  const score=calcScore(lead.scores);const tier=getTier(score);
  const pTask=(lead.tasks||[]).filter(t=>!t.done);const overT=pTask.filter(t=>isOverdue(t.dueDate));
  const filled=CRITERIA.filter(c=>lead.scores[c.id]>0).length;
  const nextTask=pTask.sort((a,b)=>(a.dueDate||"9999").localeCompare(b.dueDate||"9999"))[0];
  return(
    <div onClick={onClick} className="card-hover" style={{background:"#0d1520",border:`1px solid ${tier.accent}33`,borderRadius:12,padding:16,cursor:"pointer",position:"relative",overflow:"hidden",transition:"all .2s"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:tier.accent}}/>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
        <div style={{flex:1,paddingLeft:8,minWidth:0}}>
          <div style={{fontWeight:700,fontSize:14,color:"#e2e8f0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{lead.company||"—"}</div>
          <div style={{fontSize:11,color:"#64748b",marginTop:1}}>{lead.contact}{lead.role?` · ${lead.role}`:""}</div>
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
        <div style={{fontSize:10,color:"#334155"}}>{filled}/{CRITERIA.length} ק׳</div>
        <div style={{display:"flex",gap:4,alignItems:"center"}}>
          {nextTask&&<span style={{fontSize:10,color:overT.length?"#ef4444":"#64748b",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{overT.length?"⚠️":""}{nextTask.title}</span>}
          {pTask.length>0&&<span style={{fontSize:10,padding:"2px 6px",borderRadius:8,background:overT.length?"#7f1d1d":"#1e293b",color:overT.length?"#fca5a5":"#64748b"}}>{pTask.length}</span>}
          {lead.pipedriveOrgId&&<span style={{fontSize:10,padding:"2px 6px",borderRadius:8,background:"#1e3a5f",color:"#60a5fa"}}>🔗</span>}
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════
   APP ROOT
══════════════════════════════════ */
const VIEWS=[{key:"leads",label:"לידים",icon:"📋"},{key:"calendar",label:"לוח שנה",icon:"📅"}];

export default function App(){
  const [leads,setLeads]=useState([]);
  const [view,setView]=useState("leads");
  const [modal,setModal]=useState(null);
  const [showImport,setShowImport]=useState(false);
  const [showSettings,setShowSettings]=useState(false);
  const [quickTask,setQuickTask]=useState(null);
  const [apiToken,setApiToken]=useState("");
  const [filter,setFilter]=useState("all");
  const [sort,setSort]=useState("score");
  const [ready,setReady]=useState(false);

  useEffect(()=>{(async()=>{try{const r=await window.storage.get("sdr_v5");if(r){const d=JSON.parse(r.value);setLeads(d.leads||[]);setApiToken(d.token||"");}}catch{}setReady(true);})();},[]);

  const persist=async(l,t)=>{const l2=l??leads;const t2=t??apiToken;setLeads(l2);if(t!==undefined)setApiToken(t);try{await window.storage.set("sdr_v5",JSON.stringify({leads:l2,token:t2}));}catch{}};

  const handleSaveLead=async lead=>{const updated=lead.id?leads.map(l=>l.id===lead.id?lead:l):[...leads,{...lead,id:uid(),createdAt:new Date().toISOString()}];await persist(updated);setModal(null);};
  const handleDelete=async id=>{await persist(leads.filter(l=>l.id!==id));setModal(null);};
  const handleImport=async imported=>{const ex=new Set(leads.map(l=>l.company?.trim().toLowerCase()));const fresh=imported.filter(l=>!ex.has(l.company?.trim().toLowerCase()));await persist([...leads,...fresh]);if(imported.length-fresh.length>0)alert(`יובאו ${fresh.length}. ${imported.length-fresh.length} כפילויות דולגו.`);};
  const handleTaskUpdate=async(leadId,taskId,updates)=>{const updated=leads.map(l=>l.id!==leadId?l:{...l,tasks:(l.tasks||[]).map(t=>t.id===taskId?{...t,...updates}:t)});await persist(updated);};
  const handleAddTask=async(leadId,task)=>{const updated=leads.map(l=>l.id!==leadId?l:{...l,tasks:[...(l.tasks||[]),task]});await persist(updated);};

  /* ── Excel Export ── */
  const exportXLSX = () => {
    const rows = leads.map(l=>{
      const total=calcScore(l.scores);
      const tier=getTier(total);
      const openTasks=(l.tasks||[]).filter(t=>!t.done);
      const nextTask=openTasks.sort((a,b)=>(a.dueDate||"9999").localeCompare(b.dueDate||"9999"))[0];
      return {
        "שם חברה": l.company||"",
        "איש קשר": l.contact||"",
        "תפקיד": l.role||"",
        "טלפון": l.phone||"",
        "אימייל": l.email||"",
        "סיווג": tier.emoji+" "+tier.label,
        "Fit": fitScore(l.scores),
        "BANT": bantScore(l.scores),
        "ציון כולל": total,
        "תאריך פתיחה": l.importedAt?String(l.importedAt).slice(0,10):l.createdAt?String(l.createdAt).slice(0,10):"",
        "תווית Pipedrive": l.label||"",
        "מקושר ל-Pipedrive": l.pipedriveOrgId?"כן":"לא",
        "משימות פתוחות": openTasks.length,
        "פעילות הבאה": nextTask?nextTask.title:"",
        "תאריך פעילות הבאה": nextTask?.dueDate||"",
        "שעה": nextTask?.time||"",
        "הערות": l.notes||"",
      };
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    // Column widths
    ws["!cols"]=[{wch:20},{wch:18},{wch:16},{wch:14},{wch:24},{wch:10},{wch:6},{wch:6},{wch:10},{wch:14},{wch:12},{wch:14},{wch:12},{wch:24},{wch:14},{wch:8},{wch:40}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"לידים");

    // Summary sheet
    const hot=leads.filter(l=>getTier(calcScore(l.scores)).key==="hot");
    const warm=leads.filter(l=>getTier(calcScore(l.scores)).key==="warm");
    const cold=leads.filter(l=>getTier(calcScore(l.scores)).key==="cold");
    const summary=[{"סיווג":"🔥 HOT","כמות":hot.length},{"סיווג":"⚡ WARM","כמות":warm.length},{"סיווג":"❄️ COLD","כמות":cold.length},{"סיווג":"סה\"כ","כמות":leads.length}];
    const ws2=XLSX.utils.json_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb,ws2,"סיכום");

    XLSX.writeFile(wb,`SDR_Qualify_${isoDate(new Date())}.xlsx`);
  };

  const exportCSV=()=>{
    const cols=["שם חברה","איש קשר","תפקיד","סיווג","Fit","BANT","ציון כולל","תאריך פתיחה","פעילות הבאה","הערות"];
    const rows=leads.map(l=>{const total=calcScore(l.scores);const tier=getTier(total);const next=(l.tasks||[]).filter(t=>!t.done).sort((a,b)=>(a.dueDate||"9999").localeCompare(b.dueDate||"9999"))[0];return[l.company,l.contact,l.role,tier.label,fitScore(l.scores),bantScore(l.scores),total,String(l.importedAt||l.createdAt||"").slice(0,10),next?.title||"",l.notes];});
    const csv=[cols,...rows].map(r=>r.map(v=>`"${(v||"").toString().replace(/"/g,'""')}"`).join(",")).join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8"}));a.download=`SDR_${isoDate(new Date())}.csv`;a.click();
  };

  const counts={hot:0,warm:0,cold:0};leads.forEach(l=>{counts[getTier(calcScore(l.scores)).key]++;});
  const totalPending=leads.reduce((s,l)=>s+(l.tasks||[]).filter(t=>!t.done).length,0);
  const totalOverdue=leads.reduce((s,l)=>s+(l.tasks||[]).filter(t=>!t.done&&isOverdue(t.dueDate)).length,0);
  const sorted=[...leads].filter(l=>filter==="all"||getTier(calcScore(l.scores)).key===filter).sort((a,b)=>{if(sort==="score")return calcScore(b.scores)-calcScore(a.scores);if(sort==="name")return(a.company||"").localeCompare(b.company||"","he");if(sort==="tasks")return(b.tasks||[]).filter(t=>!t.done).length-(a.tasks||[]).filter(t=>!t.done).length;return new Date(b.createdAt||0)-new Date(a.createdAt||0);});

  const { isMobile, isTablet } = useBreakpoint();

  if(!ready) return (
    <div style={{minHeight:"100vh",background:"#070d14",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{color:"#f59e0b",fontFamily:"monospace"}}>טוען...</div>
    </div>
  );

  return(
    <div style={{height:"100vh",background:"#070d14",color:"#e2e8f0",fontFamily:"'Segoe UI',Arial,sans-serif",direction:"rtl",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <style>{CSS}</style>
      {/* Header */}
      <div style={{background:"#070d14",borderBottom:"1px solid #1e293b",padding:isMobile?"8px 12px":"10px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0,gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:isMobile?10:18}}>
          <div style={{fontSize:isMobile?15:18,fontWeight:900,color:"#e2e8f0",letterSpacing:-.5,whiteSpace:"nowrap"}}>SDR<span style={{color:"#f59e0b"}}>.</span>qualify</div>
          <div style={{display:"flex",gap:3}}>
            {VIEWS.map(v=><button key={v.key} onClick={()=>setView(v.key)} style={{...S.btn,padding:isMobile?"4px 8px":"5px 12px",fontSize:isMobile?11:12,background:view===v.key?"#f59e0b20":"none",color:view===v.key?"#f59e0b":"#475569",border:`1px solid ${view===v.key?"#f59e0b44":"transparent"}`}}>{v.icon}{!isMobile&&" "+v.label}</button>)}
          </div>
        </div>
        <div style={{display:"flex",gap:isMobile?8:14,alignItems:"center"}}>
          {/* Stats — hide some on mobile */}
          {[
            {l:"סה״כ",v:leads.length,c:"#e2e8f0",always:true},
            {l:"🔥",v:counts.hot,c:"#ef4444",always:true},
            {l:"⚡",v:counts.warm,c:"#f59e0b",always:false},
            {l:"❄️",v:counts.cold,c:"#60a5fa",always:false},
            {l:"📋",v:totalPending,c:totalOverdue?"#ef4444":"#64748b",always:true},
          ].filter(s=>!isMobile||s.always).map(s=>(
            <div key={s.l} style={{textAlign:"center"}}><div className="stat-num" style={{fontSize:17,fontWeight:900,color:s.c,fontFamily:"monospace"}}>{s.v}</div><div style={{fontSize:9,color:"#334155"}}>{s.l}</div></div>
          ))}
          <button onClick={()=>setShowSettings(true)} style={{...S.btn,background:apiToken?"#1e3a5f22":"#7f1d1d22",color:apiToken?"#60a5fa":"#fca5a5",border:`1px solid ${apiToken?"#1e3a5f":"#7f1d1d44"}`,padding:"5px 8px",fontSize:isMobile?10:12}}>
            {apiToken?"🔗":"⚙️"}{!isMobile&&(apiToken?" Pipe":" חבר")}
          </button>
        </div>
      </div>

      {/* Toolbar */}
      {view==="leads"&&<div style={{padding:isMobile?"6px 10px":"8px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #1e293b",background:"#0a0f1a",gap:8,flexWrap:"wrap",flexShrink:0}}>
        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
          {[{k:"all",l:`הכל`,c:"#e2e8f0"},{k:"hot",l:"🔥",c:"#ef4444"},{k:"warm",l:"⚡",c:"#f59e0b"},{k:"cold",l:"❄️",c:"#60a5fa"}].map(f=>(
            <button key={f.k} onClick={()=>setFilter(f.k)} style={{...S.btn,fontSize:11,padding:"4px 9px",background:filter===f.k?`${f.c}18`:"#0d1520",color:filter===f.k?f.c:"#475569",border:`1px solid ${filter===f.k?f.c:"#1e293b"}`}}>{f.l}{f.k!=="all"&&counts[f.k]>0?` ${counts[f.k]}`:""}</button>
          ))}
          {!isMobile&&<select value={sort} onChange={e=>setSort(e.target.value)} style={{...S.input,padding:"4px 8px",width:"auto",fontSize:11}}>
            <option value="score">ציון</option><option value="name">שם</option><option value="tasks">משימות</option><option value="date">תאריך</option>
          </select>}
        </div>
        <div style={{display:"flex",gap:6}}>
          {!isMobile&&<button onClick={()=>setShowImport(true)} style={{...S.btn,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",padding:"4px 10px",fontSize:11}}>⬆️ ייבא</button>}
          {!isMobile&&<button onClick={exportXLSX} disabled={!leads.length} style={{...S.btn,background:"#14532d22",color:"#22c55e",border:"1px solid #22c55e44",padding:"4px 10px",fontSize:11,opacity:leads.length?1:.5}}>📊 Excel</button>}
          {!isMobile&&<button onClick={exportCSV} disabled={!leads.length} style={{...S.btn,background:"#1e293b",color:"#94a3b8",padding:"4px 10px",fontSize:11,opacity:leads.length?1:.5}}>⬇️ CSV</button>}
          <button onClick={()=>setModal({lead:{company:"",contact:"",role:"",phone:"",email:"",notes:"",label:"",scores:{},tasks:[],id:null}})} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700,padding:"4px 12px",fontSize:isMobile?13:12}}>+ {!isMobile&&"ליד חדש"}</button>
        </div>
      </div>}

      {/* Content */}
      <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
        {view==="leads"
          ? <div style={{flex:1,overflowY:"auto",padding:isMobile?10:18}}>
              {sorted.length===0
                ? <div style={{textAlign:"center",color:"#334155",marginTop:60}}><div style={{fontSize:52,marginBottom:12}}>📋</div><div style={{fontSize:16,color:"#475569",marginBottom:8}}>{filter!=="all"?"אין לידים":"התחל בייבוא מ-Pipedrive"}</div>{filter==="all"&&<button onClick={()=>setShowImport(true)} style={{...S.btn,background:"#f59e0b",color:"#000",fontWeight:700}}>⬆️ ייבא</button>}</div>
                : <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":isTablet?"1fr 1fr":"repeat(auto-fill,minmax(280px,1fr))",gap:isMobile?10:12}}>
                    {sorted.map(l=><LeadCard key={l.id} lead={l} onClick={()=>setModal({lead:l})}/>)}
                  </div>}
              {/* Mobile action buttons */}
              {isMobile&&leads.length>0&&(
                <div style={{display:"flex",gap:8,marginTop:14,flexWrap:"wrap"}}>
                  <button onClick={()=>setShowImport(true)} style={{...S.btn,flex:1,background:"#1e3a5f",color:"#60a5fa",border:"1px solid #1e4a7f",fontSize:12,textAlign:"center"}}>⬆️ ייבא</button>
                  <button onClick={exportXLSX} style={{...S.btn,flex:1,background:"#14532d22",color:"#22c55e",border:"1px solid #22c55e44",fontSize:12}}>📊 Excel</button>
                </div>
              )}
            </div>
          : <CalendarView
              leads={leads}
              onTaskUpdate={handleTaskUpdate}
              onOpenLead={id=>{const l=leads.find(x=>x.id===id);if(l){setModal({lead:l});setView("leads");}}}
              onAddTask={(date,hour)=>setQuickTask({date,hour})}
            />}
      </div>

      {/* Footer */}
      {!isMobile&&<div style={{padding:"5px 20px",borderTop:"1px solid #1e293b",background:"#070d14",display:"flex",gap:20,alignItems:"center",fontSize:11,color:"#334155",flexShrink:0}}>
        <span>סף: <span style={{color:"#ef4444"}}>🔥≥67</span> <span style={{color:"#f59e0b"}}>⚡36-66</span> <span style={{color:"#60a5fa"}}>❄️0-35</span></span>
        <span style={{marginRight:"auto"}}>מקסימום 100 · Fit×60 + BANT×40</span>
      </div>}

      {/* Modals */}
      {modal&&<LeadModal lead={modal.lead} onSave={handleSaveLead} onDelete={handleDelete} onClose={()=>setModal(null)} apiToken={apiToken}/>}
      {showImport&&<ImportModal onImport={handleImport} onClose={()=>setShowImport(false)}/>}
      {showSettings&&<SettingsModal token={apiToken} onSave={t=>{persist(null,t);setShowSettings(false);}} onClose={()=>setShowSettings(false)}/>}
      {quickTask&&<QuickTaskModal defaultDate={quickTask.date} defaultHour={quickTask.hour} leads={leads} onAdd={handleAddTask} onClose={()=>setQuickTask(null)}/>}
    </div>
  );
}
