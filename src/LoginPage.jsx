
export default function LoginPage({ authError, config }) {
  const err = authError || new URLSearchParams(window.location.search).get('auth_error');
  const errMsg = {
    state_mismatch:        'אימות אבטחה נכשל — נסה שנית',
    token_exchange_failed: 'שגיאה בקבלת הרשאה מ-Pipedrive',
    'access_denied':       'הגישה נדחתה ב-Pipedrive',
  }[err] || (err ? `שגיאה: ${err}` : null);

  return (
    <div style={{minHeight:'100vh',background:'#070d14',display:'flex',alignItems:'center',
      justifyContent:'center',fontFamily:"'Segoe UI',Arial,sans-serif",direction:'rtl',padding:20}}>
      <div style={{background:'#0d1520',border:'1px solid #1e293b',borderRadius:20,padding:'40px 48px',
        maxWidth:440,width:'100%',textAlign:'center',boxShadow:'0 0 60px #f59e0b18'}}>

        {/* Logo */}
        <div style={{fontSize:32,fontWeight:900,color:'#e2e8f0',marginBottom:8,letterSpacing:-1}}>
          SDR<span style={{color:'#f59e0b'}}>.</span>qualify
        </div>
        <div style={{fontSize:13,color:'#475569',marginBottom:40}}>
          מערכת ניהול ובחינת לידים — PlantSharp MES
        </div>

        {/* Error */}
        {errMsg && (
          <div style={{background:'#7f1d1d22',border:'1px solid #7f1d1d',borderRadius:10,
            padding:'10px 16px',color:'#fca5a5',fontSize:13,marginBottom:24}}>
            ❌ {errMsg}
          </div>
        )}

        {/* Login button */}
        {config?.oauthEnabled ? (
          <>
            <a href="/api/auth/login"
              style={{display:'flex',alignItems:'center',justifyContent:'center',gap:12,
                background:'#f59e0b',color:'#000',fontWeight:700,fontSize:16,
                padding:'14px 28px',borderRadius:12,textDecoration:'none',
                transition:'all .2s',boxShadow:'0 4px 20px #f59e0b44',cursor:'pointer'}}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 14.5v-9l6 4.5-6 4.5z"/>
              </svg>
              התחבר באמצעות Pipedrive
            </a>
            <div style={{fontSize:12,color:'#334155',marginTop:16}}>
              הכניסה מאובטחת דרך OAuth 2.0 של Pipedrive
            </div>
          </>
        ) : config?.fallbackToken ? (
          <div style={{background:'#f59e0b18',border:'1px solid #f59e0b44',borderRadius:10,padding:16,color:'#f59e0b',fontSize:13}}>
            ⚠️ OAuth לא מוגדר. המערכת פועלת עם PD_TOKEN קבוע.
            <br/>
            <a href="/" style={{color:'#f59e0b',marginTop:8,display:'block'}}>→ כניסה ישירה</a>
          </div>
        ) : (
          <div style={{background:'#7f1d1d22',border:'1px solid #7f1d1d44',borderRadius:10,padding:16,color:'#fca5a5',fontSize:13}}>
            ⚙️ המערכת לא מוגדרת עדיין.<br/>
            יש להגדיר Environment Variables ב-Vercel.
          </div>
        )}

        <div style={{marginTop:40,padding:'16px',background:'#0a0f1a',borderRadius:10,
          fontSize:12,color:'#475569',textAlign:'right',lineHeight:2}}>
          <strong style={{color:'#64748b',display:'block',marginBottom:4}}>מידע על האפליקציה</strong>
          כל המשתמשים מתחברים עם חשבון Pipedrive האישי שלהם.<br/>
          הפעולות יירשמו תחת המשתמש המחובר.
        </div>
      </div>
    </div>
  );
}
