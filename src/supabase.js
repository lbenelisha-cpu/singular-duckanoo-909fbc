import { createClient } from '@supabase/supabase-js'

const RUNTIME_URL_KEY = 'iml_supabase_url_override'

export function normalizeSupabaseUrl(value = '') {
  const cleaned = String(value).trim().replace(/[\u200B-\u200D\uFEFF]/g, '')
  if (!cleaned) return ''
  return cleaned
    .replace(/\/rest\/v1\/?$/i, '')
    .replace(/\/auth\/v1\/?$/i, '')
    .replace(/\/+$/, '')
}

function readRuntimeUrl() {
  try {
    return normalizeSupabaseUrl(window.localStorage.getItem(RUNTIME_URL_KEY) || '')
  } catch {
    return ''
  }
}

export function saveRuntimeSupabaseUrl(value) {
  const normalized = normalizeSupabaseUrl(value)
  if (!/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(normalized)) {
    throw new Error('כתובת Supabase אינה בפורמט תקין.')
  }
  window.localStorage.setItem(RUNTIME_URL_KEY, normalized)
  return normalized
}

export function clearRuntimeSupabaseUrl() {
  try { window.localStorage.removeItem(RUNTIME_URL_KEY) } catch { /* no-op */ }
}

export const envSupabaseUrl = normalizeSupabaseUrl(import.meta.env.VITE_SUPABASE_URL || '')
export const runtimeSupabaseUrl = typeof window !== 'undefined' ? readRuntimeUrl() : ''
export const supabaseUrl = runtimeSupabaseUrl || envSupabaseUrl
export const rawSupabaseUrl = import.meta.env.VITE_SUPABASE_URL || ''
export const supabaseAnonKey = (
  import.meta.env.VITE_SUPABASE_ANON_KEY ||
  import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY ||
  ''
).trim()

export const urlWasNormalized = Boolean(rawSupabaseUrl && rawSupabaseUrl.trim() !== envSupabaseUrl)
export const usingRuntimeUrl = Boolean(runtimeSupabaseUrl)
export const cloudConfigured = Boolean(supabaseUrl && supabaseAnonKey)
export const configurationError = !cloudConfigured
  ? 'חסרים משתני החיבור של Supabase.'
  : !/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(supabaseUrl)
    ? 'כתובת Supabase אינה בפורמט תקין.'
    : ''

export const supabase = cloudConfigured && !configurationError
  ? createClient(supabaseUrl, supabaseAnonKey, {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true,
      },
    })
  : null

export async function testSupabaseConnection() {
  if (!supabaseUrl || configurationError) {
    return { ok: false, message: configurationError || 'חיבור Supabase לא הוגדר.' }
  }

  try {
    const controller = new AbortController()
    const timeout = window.setTimeout(() => controller.abort(), 10000)
    const response = await fetch(`${supabaseUrl}/auth/v1/settings`, {
      headers: { apikey: supabaseAnonKey },
      signal: controller.signal,
    })
    window.clearTimeout(timeout)
    if (!response.ok) {
      return { ok: false, message: `Supabase החזיר שגיאה ${response.status}.`, url: supabaseUrl }
    }
    return { ok: true, message: 'החיבור ל־Supabase תקין.', url: supabaseUrl }
  } catch (error) {
    const hostname = (() => {
      try { return new URL(supabaseUrl).hostname } catch { return supabaseUrl }
    })()
    return {
      ok: false,
      message: `לא ניתן להגיע לשרת ${hostname}. ניתן לתקן את כתובת הפרויקט במסך זה ללא שינוי קוד.`,
      technical: error?.message || String(error),
      url: supabaseUrl,
    }
  }
}
