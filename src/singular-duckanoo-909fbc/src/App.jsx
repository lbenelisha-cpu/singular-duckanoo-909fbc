import { useEffect, useRef, useState } from 'react'
import { LockKeyhole, Mail, ShieldCheck, AlertCircle, Cloud, Eye, EyeOff, UserRoundCheck } from 'lucide-react'
import DashboardApp from './DashboardApp'
import { cloudConfigured, configurationError, supabase, supabaseUrl, testSupabaseConnection, urlWasNormalized } from './supabase'
import './styles.css'

export default function App() {
  const [session, setSession] = useState(null)
  const [profile, setProfile] = useState(null)
  const [profileLoading, setProfileLoading] = useState(true)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState('')
  const [connectionStatus, setConnectionStatus] = useState(null)
  const [adminLoginOpen, setAdminLoginOpen] = useState(false)

  const sessionRef = useRef(null)

  useEffect(() => {
    if (!cloudConfigured || configurationError) return
    let active = true

    testSupabaseConnection().then(result => { if (active) setConnectionStatus(result) })

    const applySession = (nextSession, event = '') => {
      if (!active) return

      if (!nextSession) {
        sessionRef.current = null
        setSession(null)
        setProfile(null)
        setProfileLoading(false)
        return
      }

      const previousUserId = sessionRef.current?.user?.id || ''
      const nextUserId = nextSession?.user?.id || ''
      sessionRef.current = nextSession

      // Supabase commonly emits TOKEN_REFRESHED / duplicate SIGNED_IN events when the
      // browser tab becomes active again.  Updating React state for the same user
      // would remount DashboardApp, re-run the large cloud/cache restore and flash
      // the permissions screen.  The Supabase client keeps the refreshed token
      // internally, so React only needs a new session when the actual user changes.
      if (!previousUserId || previousUserId !== nextUserId || event === 'USER_UPDATED') {
        setSession(nextSession)
      }
    }

    supabase.auth.getSession().then(({ data }) => applySession(data.session || null, 'INITIAL_SESSION'))
    const { data } = supabase.auth.onAuthStateChange((event, nextSession) => {
      if (event === 'SIGNED_OUT') {
        applySession(null, event)
        return
      }
      applySession(nextSession, event)
    })

    return () => {
      active = false
      data.subscription.unsubscribe()
    }
  }, [])

  const sessionUserId = session?.user?.id || ''
  const isAnonymousUser = Boolean(session?.user?.is_anonymous)

  useEffect(() => {
    let active = true

    if (!sessionUserId || !cloudConfigured) {
      setProfile(null)
      setProfileLoading(false)
      return () => { active = false }
    }

    setProfileLoading(true)

    // Anonymous Supabase users are always treated as read-only guests.
    if (isAnonymousUser) {
      setProfile({ id: sessionUserId, email: '', full_name: 'אורח', role: 'viewer', is_active: true, is_guest: true })
      setMessage('')
      setProfileLoading(false)
      return () => { active = false }
    }

    supabase.from('profiles').select('id,email,full_name,role,is_active').eq('id', sessionUserId).maybeSingle()
      .then(({ data, error }) => {
        if (!active) return
        if (error) setMessage('המשתמש התחבר, אך פרופיל ההרשאה עדיין לא הוגדר.')
        setProfile(data || null)
        setProfileLoading(false)
      })

    return () => { active = false }
  }, [sessionUserId, isAnonymousUser])

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

  const requestAdminLogin = () => { setMessage(''); setAdminLoginOpen(true) }

  const signOut = async () => {
    if (supabase) await supabase.auth.signOut({ scope: 'local' })
    setAdminLoginOpen(false)
  }

  if (!cloudConfigured || configurationError) return <SetupRequired error={configurationError} />
  const dashboard = <DashboardApp currentUser={session?.user || null} userRole={profile?.role || 'viewer'} isGuest={!session || Boolean(profile?.is_guest || session?.user?.is_anonymous)} onSignOut={signOut} onRequestAdminLogin={requestAdminLogin}/>
  if (session && !profileLoading && !profile) return <div className="auth-page" dir="rtl"><div className="auth-card"><AlertCircle className="blocked-icon"/><h1>לא נמצא פרופיל הרשאה</h1><p>יש להגדיר למשתמש פרופיל פעיל בטבלת profiles.</p><button className="auth-submit" onClick={signOut}>יציאה</button></div></div>
  if (profile && profile.is_active === false) return <div className="auth-page" dir="rtl"><div className="auth-card"><AlertCircle className="blocked-icon"/><h1>החשבון חסום</h1><p>פנה למנהל המערכת להפעלת המשתמש.</p><button className="auth-submit" onClick={signOut}>יציאה</button></div></div>
  return <>{dashboard}{!session && adminLoginOpen && <div className="auth-modal-backdrop" dir="rtl" onMouseDown={e=>{if(e.target===e.currentTarget)setAdminLoginOpen(false)}}><form className="auth-card auth-modal-card" onSubmit={signIn}><button type="button" className="auth-modal-close" onClick={()=>setAdminLoginOpen(false)}>×</button><div className="auth-brand"><img src="/icons/icon-192.png" alt="IML Control"/><div className="auth-logo">IML<span>CONTROL</span></div></div><div className="auth-icon"><LockKeyhole/></div><h1>כניסת מנהל</h1><p>התחברות נדרשת רק לפעולות ניהול וטעינת קבצים.</p><label><span>דוא״ל</span><div><Mail size={18}/><input type="email" value={email} onChange={e=>setEmail(e.target.value)} required autoComplete="email"/></div></label><label><span>סיסמה</span><div><LockKeyhole size={18}/><input type={showPassword?'text':'password'} value={password} onChange={e=>setPassword(e.target.value)} required autoComplete="current-password"/><button type="button" onClick={()=>setShowPassword(v=>!v)}>{showPassword?<EyeOff size={18}/>:<Eye size={18}/>}</button></div></label>{message&&<div className="auth-error"><AlertCircle size={18}/>{message}</div>}<button className="auth-submit" disabled={busy}>{busy?'מתחבר...':'כניסת מנהל'}</button></form></div>}</>
}

function SetupRequired({ error }){ return <div className="auth-page" dir="rtl"><div className="auth-card setup-card"><div className="auth-icon"><Cloud/></div><h1>נדרש חיבור לענן</h1><p>{error || 'חסרים משתני החיבור של Supabase.'}</p><code>VITE_SUPABASE_URL</code><code>VITE_SUPABASE_ANON_KEY</code><small>נתמך גם: VITE_SUPABASE_PUBLISHABLE_KEY</small><div className="setup-note">העתק את Project URL ישירות מ־Supabase, ללא /rest/v1/, ושמור ב־Netlify. לאחר מכן בצע Clear cache and deploy site.</div></div></div> }
