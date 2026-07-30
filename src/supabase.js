import { createClient } from '@supabase/supabase-js'

// Canonical project URL for this IML deployment. The project URL is public;
// the publishable key remains supplied by Netlify environment variables.
const IML_CANONICAL_URL = 'https://gfrshnnfvdlwdkoeeear.supabase.co'
const IML_PROJECT_PREFIX = 'gfrshnnfvdlwdko'

function normalizeSupabaseUrl(value = '') {
  const cleaned = String(value).trim().replace(/[\u200B-\u200D\uFEFF]/g, '')
  if (!cleaned) return ''

  const normalized = cleaned
    .replace(/\/rest\/v1\/?$/i, '')
    .replace(/\/auth\/v1\/?$/i, '')
    .replace(/\/+$/, '')

  // Older deploys contained a typo with extra "e" characters in the project
  // hostname. Correct only this project's known prefix; never alter other URLs.
  try {
    const parsed = new URL(normalized)
    if (
      parsed.hostname.startsWith(IML_PROJECT_PREFIX) &&
      parsed.hostname.endsWith('.supabase.co') &&
      parsed.hostname !== new URL(IML_CANONICAL_URL).hostname
    ) {
      return IML_CANONICAL_URL
    }
  } catch {
    return normalized
  }

  return normalized
}

export const rawSupabaseUrl = import.meta.env.VITE_SUPABASE_URL || ''
export const supabaseUrl = normalizeSupabaseUrl(rawSupabaseUrl || IML_CANONICAL_URL)
export const supabaseAnonKey = (
  import.meta.env.VITE_SUPABASE_ANON_KEY ||
  import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY ||
  ''
).trim()

export const urlWasNormalized = Boolean(rawSupabaseUrl && rawSupabaseUrl.trim() !== supabaseUrl)
export const cloudConfigured = Boolean(supabaseUrl && supabaseAnonKey)
export const configurationError = !supabaseAnonKey
  ? 'חסר מפתח Publishable של Supabase ב־Netlify.'
  : !/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(supabaseUrl)
    ? 'כתובת Supabase אינה בפורמט תקין.'
    : ''

function supabaseFetch(input, init = {}) {
  const requestUrl = typeof input === 'string' ? input : input?.url || ''
  const isSupabaseRequest = requestUrl.startsWith(supabaseUrl)

  if (!isSupabaseRequest) return fetch(input, init)

  const headers = new Headers(input instanceof Request ? input.headers : undefined)
  const initHeaders = new Headers(init.headers || {})
  initHeaders.forEach((value, key) => headers.set(key, value))

  // Every Supabase API request must carry the project key. Supabase-js normally
  // adds it, but this explicit guard prevents browser/build edge cases that
  // produced "No API key found in request" during Excel uploads.
  if (!headers.has('apikey')) headers.set('apikey', supabaseAnonKey)

  const requestInit = { ...init, headers }

  if (input instanceof Request) {
    return fetch(new Request(input, requestInit))
  }
  return fetch(input, requestInit)
}

export const supabase = cloudConfigured && !configurationError
  ? createClient(supabaseUrl, supabaseAnonKey, {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true,
        storageKey: 'iml-control-auth-v9212',
      },
      global: {
        fetch: supabaseFetch,
        headers: {
          apikey: supabaseAnonKey,
        },
      },
    })
  : null

export async function testSupabaseConnection() {
  if (!supabaseUrl || configurationError) {
    return { ok: false, message: configurationError || 'חיבור Supabase לא הוגדר.' }
  }

  try {
    const response = await fetch(`${supabaseUrl}/auth/v1/settings`, {
      headers: { apikey: supabaseAnonKey },
      cache: 'no-store',
    })
    if (!response.ok) {
      return { ok: false, message: `Supabase החזיר שגיאה ${response.status}.` }
    }
    return { ok: true, message: 'החיבור ל־Supabase תקין.' }
  } catch (error) {
    const hostname = (() => {
      try { return new URL(supabaseUrl).hostname } catch { return supabaseUrl }
    })()
    return {
      ok: false,
      message: `לא ניתן להגיע לשרת ${hostname}. בדוק חיבור לאינטרנט ובצע רענון מלא.`,
      technical: error?.message || String(error),
    }
  }
}
