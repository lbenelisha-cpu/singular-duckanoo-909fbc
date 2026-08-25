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
  // Normalize every fetch input form supported by the Fetch API: string, URL,
  // or Request. The previous implementation did not recognize URL objects,
  // so some PostgREST requests bypassed the API-key guard and Supabase replied
  // with "No API key found in request".
  const baseRequest = new Request(input, init)
  let isSupabaseRequest = false

  try {
    const requestUrl = new URL(baseRequest.url)
    const projectUrl = new URL(supabaseUrl)
    isSupabaseRequest = requestUrl.origin === projectUrl.origin
  } catch {
    isSupabaseRequest = String(baseRequest.url || '').startsWith(supabaseUrl)
  }

  if (!isSupabaseRequest) return fetch(baseRequest)

  const headers = new Headers(baseRequest.headers)
  headers.set('apikey', supabaseAnonKey)

  // Preserve the authenticated user's Bearer token when supabase-js supplied
  // one. For anonymous project calls, provide the publishable key as fallback.
  if (!headers.has('Authorization')) {
    headers.set('Authorization', `Bearer ${supabaseAnonKey}`)
  }

  return fetch(new Request(baseRequest, { headers }))
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
