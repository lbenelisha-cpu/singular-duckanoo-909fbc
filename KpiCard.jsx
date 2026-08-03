import { useEffect, useState } from 'react'
import { LockKeyhole, Mail, ShieldCheck, AlertCircle, Cloud, Eye, EyeOff, UserRoundCheck } from 'lucide-react'
import DashboardApp from './DashboardApp'
import { cloudConfigured, configurationError, supabase, supabaseUrl, testSupabaseConnection, urlWasNormalized } from './supabase'
import './styles.css'

export default function App() {
  const [session, setSession] = useState(null)
  const [sessionReady, setSessionReady] = useState(false)
  const [profile, setProfile] = useState(null)
  const [profileReady, setProfileReady] = useState(false)
  const [startupStage, setStartupStage] = useState('בודק חיבור לשרת...')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState('')
  const [connectionStatus, setConnectionStatus] = useState(null)
  const [publicViewer, setPublicViewer] = useState(false)

  useEffect(() => {
    if (!cloudConfigured || configurationError) return
    testSupabaseConnection().then(setConnectionStatus)
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session || null)
      setSessionReady(true)
      setStartupStage(data.session ? 'בודק הרשאות משתמש...' : 'מוכן להתחברות')
    }).catch(() => {
      setSessionReady(true)
      setStartupStage('מוכן להתחברות')
    })
    const { data } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession)
      setSessionReady(true)
      setProfile(null)
      setProfileReady(!nextSession)
      setStartupStage(nextSession ? 'בודק הרשאות משתמש...' : 'מוכן להתחברות')
    })
    return () => data.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (!session?.user || !cloudConfigured) { setProfile(null); setProfileReady(true); return }
    setProfileReady(false)
    setStartupStage('בודק הרשאות משתמש...')

    // Anonymous Supabase users are always treated as read-only guests.
    if (session.user.is_anonymous) {
      setProfile({ email: '', full_name: 'אורח', role: 'viewer', is_active: true, is_guest: true })
      setProfileReady(true)
      setStartupStage('ההרשאות אומתו')
      setMessage('')
      return
    }

    const cacheKey = `iml-profile-${session.user.id}`
    try {
      const cached = JSON.parse(localStorage.getItem(cacheKey) || 'null')
      if (cached?.role) setStartupStage(`מאמת הרשאת ${cached.role === 'admin' ? 'מנהל מערכת' : cached.role === 'manager' ? 'מנהל מתקן' : 'צפייה'}...`)
    } catch {}

    supabase.from('profiles').select('id,email,full_name,role,is_active').eq('id', session.user.id).maybeSingle()
      .then(({ data, error }) => {
        if (error) {
          setMessage('המשתמש התחבר, אך לא ניתן היה לאמת את פרופיל ההרשאה.')
          setProfile(null)
          return
        }
        const verifiedProfile = data || { email: session.user.email, role: 'viewer', is_active: true }
        setProfile(verifiedProfile)
        try { localStorage.setItem(cacheKey, JSON.stringify({ ...verifiedProfile, verifiedAt: Date.now() })) } catch {}
        setStartupStage('ההרשאות אומתו')
      })
      .finally(() => setProfileReady(true))
  }, [session])

  const signIn = async (event) => {
    event.preventDefault(); setBusy(true); setMessage('')
    try {
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password })
      if (error) setMessage(error.message === 'Invalid login credentials' ? 'כתובת הדוא״ל או הסיסמה אינם נכונים.' : error.message)
    } catch (error) {
      setMessage(connectionStatus?.message || `לא ניתן להגיע לשרת Supabase. כתובת החיבור הפעילה: ${supabaseUrl}`)
    } finally {
      setBusy(false)
    }
  }

  const enterPublicViewer = () => {
    setMessage('')
    setPublicViewer(true)
  }


  const signOut = async () => {
    if (publicViewer) {
      setPublicViewer(false)
      return
    }
    if (supabase) await supabase.auth.signOut({ scope: 'local' })
  }

  if (!cloudConfigured || configurationError) return <SetupRequired error={configurationError} />
  if (!sessionReady) return <StartupScreen stage={startupStage} />
  if (publicViewer) return <DashboardApp currentUser={null} userRole="viewer" isGuest={true} onSignOut={signOut}/>
  if (!session) return <div className="auth-page" dir="rtl"><form className="auth-card" onSubmit={signIn}>
    <div className="auth-brand"><img src="/icons/icon-192.png" alt="IML Control"/><div className="auth-logo">IML<span>CONTROL</span></div></div>
    <div className="auth-icon"><LockKeyhole/></div><h1>כניסה למערכת</h1><p>התחבר כדי לצפות בנתוני המתקנים המשותפים.</p>
    <label><span>דוא״ל</span><div><Mail size={18}/><input type="email" value={email} onChange={e=>setEmail(e.target.value)} required autoComplete="email"/></div></label>
    <label><span>סיסמה</span><div><LockKeyhole size={18}/><input type={showPassword?'text':'password'} value={password} onChange={e=>setPassword(e.target.value)} required autoComplete="current-password"/><button type="button" onClick={()=>setShowPassword(v=>!v)}>{showPassword?<EyeOff size={18}/>:<Eye size={18}/>}</button></div></label>
    {urlWasNormalized && <div className="auth-info">כתובת Supabase נוקתה אוטומטית מתוספת /rest/v1/.</div>}
    {connectionStatus && !connectionStatus.ok && <div className="auth-error"><AlertCircle size={18}/>{connectionStatus.message}</div>}
    {message && <div className="auth-error"><AlertCircle size={18}/>{message}</div>}
    <button className="auth-submit" disabled={busy}>{busy?'מתחבר...':'כניסת מנהל'}</button>
    <div className="auth-divider"><span>או</span></div>
    <button className="auth-guest" type="button" disabled={busy || (connectionStatus && !connectionStatus.ok)} onClick={enterPublicViewer}>
      <UserRoundCheck size={19}/>{busy?'מתחבר...':'צפייה בלבד'}
    </button>
    <div className="guest-note"><ShieldCheck size={16}/> משתמש צפייה יכול לצפות, לסנן, לחפש ולייצא בלבד. טעינה, מחיקה ושינוי יעדים חסומים.</div>
    <small><ShieldCheck size={15}/> מנהלים מתחברים באמצעות Supabase Auth. מצב צפייה פועל ללא חשבון ובקריאה בלבד. · Sprint 11.4.3 Build 3</small>
  </form></div>
  if (session && !profileReady) return <StartupScreen stage={startupStage} />
  if (session && profileReady && !profile) return <div className="auth-page" dir="rtl"><div className="auth-card"><AlertCircle className="blocked-icon"/><h1>לא ניתן לאמת הרשאה</h1><p>{message || 'נסה לרענן את הדף או להתחבר מחדש.'}</p><button className="auth-submit" onClick={signOut}>חזרה לכניסה</button></div></div>
  if (profile && profile.is_active === false) return <div className="auth-page" dir="rtl"><div className="auth-card"><AlertCircle className="blocked-icon"/><h1>החשבון חסום</h1><p>פנה למנהל המערכת להפעלת המשתמש.</p><button className="auth-submit" onClick={signOut}>יציאה</button></div></div>
  return <DashboardApp currentUser={session.user} userRole={profile.role} isGuest={Boolean(profile?.is_guest || session.user.is_anonymous)} onSignOut={signOut}/>
}

function SetupRequired({ error }){ return <div className="auth-page" dir="rtl"><div className="auth-card setup-card"><div className="auth-icon"><Cloud/></div><h1>נדרש חיבור לענן</h1><p>{error || 'חסרים משתני החיבור של Supabase.'}</p><code>VITE_SUPABASE_URL</code><code>VITE_SUPABASE_ANON_KEY</code><small>נתמך גם: VITE_SUPABASE_PUBLISHABLE_KEY</small><div className="setup-note">העתק את Project URL ישירות מ־Supabase, ללא /rest/v1/, ושמור ב־Netlify. לאחר מכן בצע Clear cache and deploy site.</div></div></div> }

function StartupScreen({ stage }) { return <div className="auth-page startup-page" dir="rtl"><div className="auth-card startup-card"><div className="auth-brand"><img src="/icons/icon-192.png" alt="IML Control"/><div className="auth-logo">IML<span>CONTROL</span></div></div><div className="startup-spinner" aria-hidden="true"/><h1>מפעיל את המערכת</h1><p>{stage || 'בודק הרשאות...'}</p><div className="startup-progress"><i/></div><small><ShieldCheck size={15}/> הממשק יוצג רק לאחר אימות ההרשאה — ללא מעבר זמני למצב צפייה.</small></div></div> }
