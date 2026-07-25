import { createClient } from '@supabase/supabase-js'

function normalizeSupabaseUrl(value = '') {
  const cleaned = String(value).trim().replace(/[\u200B-\u200D\uFEFF]/g, '')
  if (!cleaned) return ''
  return cleaned
    .replace(/\/rest\/v1\/?$/i, '')
    .replace(/\/auth\/v1\/?$/i, '')
    .replace(/\/+$/, '')
}

export const rawSupabaseUrl = import.meta.env.VITE_SUPABASE_URL || ''
export const supabaseUrl = normalizeSupabaseUrl(rawSupabaseUrl)
export const supabaseAnonKey = (
  import.meta.env.VITE_SUPABASE_ANON_KEY ||
  import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY ||
  ''
).trim()

export const urlWasNormalized = Boolean(rawSupabaseUrl && rawSupabaseUrl.trim() !== supabaseUrl)
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
    const response = await fetch(`${supabaseUrl}/auth/v1/settings`, {
      headers: { apikey: supabaseAnonKey },
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
      message: `לא ניתן להגיע לשרת ${hostname}. יש להעתיק מחדש את Project URL המדויק מ־Supabase ולבצע Deploy חדש.`,
      technical: error?.message || String(error),
    }
  }
}
