import { useEffect, useState } from 'react'
import { LockKeyhole, Mail, ShieldCheck, AlertCircle, Cloud, Eye, EyeOff, Settings2 } from 'lucide-react'
import DashboardApp from './DashboardApp'
import {
  cloudConfigured,
  configurationError,
  supabase,
  supabaseUrl,
  testSupabaseConnection,
  urlWasNormalized,
  usingRuntimeUrl,
  saveRuntimeSupabaseUrl,
  clearRuntimeSupabaseUrl,
} from './supabase'
import './styles.css'

export default function App() {
  const [session, setSession] = useState(null)
  const [profile, setProfile] = useState(null)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState('')
  const [connectionStatus, setConnectionStatus] = useState(null)
  const [showConnectionEditor, setShowConnectionEditor] = useState(false)
  const [connectionUrl, setConnectionUrl] = useState(supabaseUrl)

  useEffect(() => {
    if (!cloudConfigured || configurationError || !supabase) return
    testSupabaseConnection().then(setConnectionStatus)
    supabase.auth.getSession().then(({ data }) => setSession(data.session || null))
    const { data } = supabase.auth.onAuthStateChange((_event, nextSession) => setSession(nextSession))
    return () => data.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (!session?.user || !supabase) { setProfile(null); return }
    supabase.from('profiles').select('id,email,full_name,role,is_active').eq('id', session.user.id).maybeSingle()
      .then(({ data, error }) => {
        // Authentication remains usable even before the optional profiles table is installed.
        if (error) {
          console.info('Profiles table is not available yet; using safe viewer fallback.', error.message)
        }
        setProfile(data || {
          id: session.user.id,
          email: session.user.email,
          full_name: session.user.user_metadata?.full_name || '',
          role: 'viewer',
          is_active: true,
        })
      })
  }, [session])

  const signIn = async (event) => {
    event.preventDefault()
    setBusy(true)
    setMessage('')
    try {
      const check = await testSupabaseConnection()
      setConnectionStatus(check)
      if (!check.ok) {
        setShowConnectionEditor(true)
        return
      }
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password })
      if (error) {
        setMessage(error.message === 'Invalid login credentials'
          ? 'כתובת הדוא״ל או הסיסמה אינם נכונים.'
          : error.message)
      }
    } catch (error) {
      setMessage(error?.message || 'לא ניתן להשלים את ההתחברות.')
    } finally {
      setBusy(false)
    }
  }

  const saveConnection = () => {
    try {
      saveRuntimeSupabaseUrl(connectionUrl)
      window.location.reload()
    } catch (error) {
      setMessage(error.message)
    }
  }

  const restoreNetlifyConnection = () => {
    clearRuntimeSupabaseUrl()
    window.location.reload()
  }

  const signOut = async () => { if (supabase) await supabase.auth.signOut() }

  if (!cloudConfigured || configurationError) {
    return <SetupRequired error={configurationError} connectionUrl={connectionUrl} setConnectionUrl={setConnectionUrl} saveConnection={saveConnection} />
  }

  if (!session) return <div className="auth-page" dir="rtl"><form className="auth-card" onSubmit={signIn}>
    <div className="auth-logo">IML<span>CONTROL</span></div>
    <div className="auth-icon"><LockKeyhole/></div><h1>כניסה למערכת</h1><p>התחבר כדי לצפות בנתוני המתקנים המשותפים.</p>
    <label><span>דוא״ל</span><div><Mail size={18}/><input type="email" value={email} onChange={e=>setEmail(e.target.value)} required autoComplete="email"/></div></label>
    <label><span>סיסמה</span><div><LockKeyhole size={18}/><input type={showPassword?'text':'password'} value={password} onChange={e=>setPassword(e.target.value)} required autoComplete="current-password"/><button type="button" onClick={()=>setShowPassword(v=>!v)}>{showPassword?<EyeOff size={18}/>:<Eye size={18}/>}</button></div></label>
    {urlWasNormalized && <div className="auth-info">כתובת Supabase נוקתה אוטומטית מתוספת /rest/v1/.</div>}
    {usingRuntimeUrl && <div className="auth-info">המערכת משתמשת בכתובת חיבור שתוקנה מקומית.</div>}
    {connectionStatus && !connectionStatus.ok && <div className="auth-error"><AlertCircle size={18}/>{connectionStatus.message}</div>}
    {message && <div className="auth-error"><AlertCircle size={18}/>{message}</div>}

    {(showConnectionEditor || (connectionStatus && !connectionStatus.ok)) && <div className="connection-editor">
      <div className="connection-editor-title"><Settings2 size={17}/> תיקון כתובת Supabase</div>
      <input dir="ltr" value={connectionUrl} onChange={e=>setConnectionUrl(e.target.value)} placeholder="https://PROJECT-REF.supabase.co" />
      <button type="button" className="connection-save" onClick={saveConnection}>שמור ובדוק מחדש</button>
      {usingRuntimeUrl && <button type="button" className="connection-reset" onClick={restoreNetlifyConnection}>חזור לכתובת Netlify</button>}
      <small>העתק את VITE_SUPABASE_URL מחלון Connect של Supabase. אין להוסיף ‎/rest/v1/‎.</small>
    </div>}

    <button className="auth-submit" disabled={busy}>{busy?'מתחבר...':'כניסה'}</button>
    {!showConnectionEditor && <button type="button" className="connection-link" onClick={()=>setShowConnectionEditor(true)}><Settings2 size={15}/> הגדרות חיבור</button>}
    <small><ShieldCheck size={15}/> הגישה נשלטת באמצעות Supabase Auth והרשאות תפקיד.</small>
  </form></div>

  if (profile && profile.is_active === false) return <div className="auth-page" dir="rtl"><div className="auth-card"><AlertCircle className="blocked-icon"/><h1>החשבון חסום</h1><p>פנה למנהל המערכת להפעלת המשתמש.</p><button className="auth-submit" onClick={signOut}>יציאה</button></div></div>

  return <DashboardApp currentUser={session.user} userRole={profile?.role || 'viewer'} onSignOut={signOut}/>
}

function SetupRequired({ error, connectionUrl, setConnectionUrl, saveConnection }) {
  return <div className="auth-page" dir="rtl"><div className="auth-card setup-card">
    <div className="auth-icon"><Cloud/></div><h1>נדרש חיבור לענן</h1><p>{error || 'חסרים משתני החיבור של Supabase.'}</p>
    <code>VITE_SUPABASE_URL</code><code>VITE_SUPABASE_ANON_KEY</code><small>נתמך גם: VITE_SUPABASE_PUBLISHABLE_KEY</small>
    <div className="connection-editor">
      <input dir="ltr" value={connectionUrl} onChange={e=>setConnectionUrl(e.target.value)} placeholder="https://PROJECT-REF.supabase.co" />
      <button type="button" className="connection-save" onClick={saveConnection}>שמור כתובת חיבור</button>
    </div>
    <div className="setup-note">ניתן להזין כאן את Project URL המדויק. המפתח עדיין נקרא באופן מאובטח ממשתני Netlify.</div>
  </div></div>
}
