import { supabase } from './supabase'

export const CLOUD_KINDS = ['production', 'quality', 'deviations', 'targets']
export const FUTURE_CLOUD_KINDS = [...CLOUD_KINDS, 'packaging_plan']

// Keep every request deliberately small. A quality file can contain hundreds of
// thousands of rows, and large JSONB responses are what caused the database
// statement timeouts in the previous build.
// Build 11.6.1 Large File Fast Upload: compact rows are small, so 300 rows per
// JSONB chunk stays comfortably below request limits while cutting the number of
// round trips by ~4x. Two chunks are sent per request and up to four requests
// run concurrently (max ~2,400 compact rows in flight).
const ROWS_PER_CHUNK = 300
const CHUNKS_PER_UPLOAD_REQUEST = 2
const UPLOAD_CONCURRENCY = 4
const CHUNKS_PER_SERVER_COPY_REQUEST = 200
const CHUNKS_PER_DOWNLOAD_PAGE = 100
const IS_IOS_DOWNLOAD_CLIENT = (() => {
  if (typeof navigator === 'undefined') return false
  const ua = String(navigator.userAgent || '')
  const touchMac = /Macintosh/i.test(ua) && Number(navigator.maxTouchPoints || 0) > 1
  return /iPad|iPhone|iPod/i.test(ua) || touchMac
})()
const DOWNLOAD_PAGE_SIZE = IS_IOS_DOWNLOAD_CLIENT ? 12 : CHUNKS_PER_DOWNLOAD_PAGE
const DOWNLOAD_PAGE_PAUSE_MS = IS_IOS_DOWNLOAD_CLIENT ? 25 : 0
const MAX_RETRIES = 4
const RETRY_BASE_MS = 700

let schemaCapability = null

function requireClient() {
  if (!supabase) throw new Error('Supabase client is not configured')
}

function emit(onProgress, phase, completed, total, message, extra = {}) {
  const percent = total > 0 ? Math.min(100, Math.round((completed / total) * 100)) : 0
  onProgress?.({ phase, completed, total, percent, message, ...extra })
}

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms))

function retryable(error) {
  const text = `${error?.message || ''} ${error?.details || ''} ${error?.hint || ''}`.toLowerCase()
  return text.includes('statement timeout') ||
    text.includes('canceling statement') ||
    text.includes('failed to fetch') ||
    text.includes('load failed') ||
    text.includes('network request failed') ||
    text.includes('network') ||
    text.includes('timeout') ||
    text.includes('502') ||
    text.includes('503') ||
    text.includes('504')
}

async function withRetry(operation, label = 'פעולת ענן') {
  let lastError
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
    try {
      return await operation(attempt)
    } catch (error) {
      lastError = error
      if (!retryable(error) || attempt === MAX_RETRIES) throw error
      await sleep(RETRY_BASE_MS * (2 ** attempt))
    }
  }
  throw new Error(`${label} נכשלה: ${lastError?.message || 'שגיאה לא ידועה'}`)
}

function isMissingSchema(error, names = []) {
  const text = `${error?.message || ''} ${error?.details || ''} ${error?.hint || ''}`.toLowerCase()
  return names.some(name => text.includes(name.toLowerCase())) && (
    text.includes('does not exist') || text.includes('could not find') || text.includes('schema cache')
  )
}

export async function detectCloudSchema(force = false) {
  requireClient()
  if (schemaCapability && !force) return schemaCapability

  const { error: sourceError } = await withRetry(async () => {
    const result = await supabase
      .from('iml_data_sources')
      .select('kind,active_version_id')
      .limit(1)
    if (result.error && retryable(result.error)) throw result.error
    return result
  }, 'בדיקת חיבור ל-Supabase')

  if (!sourceError) {
    const { error: versionError } = await withRetry(async () => {
      const result = await supabase
        .from('iml_dataset_versions')
        .select('id')
        .limit(1)
      if (result.error && retryable(result.error)) throw result.error
      return result
    }, 'בדיקת סכמת הנתונים')
    schemaCapability = {
      legacy: true,
      versioned: !versionError,
      message: versionError ? 'טבלת גרסאות הנתונים עדיין אינה מותקנת' : 'סכמת Sprint 10.2 פעילה',
      error: versionError || null,
    }
    return schemaCapability
  }

  if (isMissingSchema(sourceError, ['active_version_id'])) {
    const { error: legacyError } = await withRetry(async () => {
      const result = await supabase.from('iml_data_sources').select('kind').limit(1)
      if (result.error && retryable(result.error)) throw result.error
      return result
    }, 'בדיקת סכמת ענן ישנה')
    if (!legacyError) {
      schemaCapability = {
        legacy: true,
        versioned: false,
        message: 'מבנה הענן הישן פעיל; יש להריץ את קובץ ההגירה',
        error: sourceError,
      }
      return schemaCapability
    }
  }

  throw sourceError
}

async function readChunkPages({ table, filterColumn, filterValue, expectedChunks = 0, onProgress, kind }) {
  const rows = []
  let from = 0

  while (true) {
    const to = from + DOWNLOAD_PAGE_SIZE - 1
    const page = await withRetry(async attempt => {
      const { data, error } = await supabase
        .from(table)
        .select('chunk_index,payload')
        .eq(filterColumn, filterValue)
        .order('chunk_index', { ascending: true })
        .range(from, to)
      if (error) throw error
      return data || []
    }, `קריאת ${kind || 'נתונים'} מהענן`)

    for (const chunk of page) {
      if (Array.isArray(chunk.payload)) rows.push(...chunk.payload)
    }

    const completedChunks = Math.min(from + page.length, expectedChunks || from + page.length)
    emit(
      onProgress,
      'download',
      completedChunks,
      expectedChunks || Math.max(completedChunks, 1),
      `מוריד ${kind || 'נתונים'} מהענן`,
      { downloadedRows: rows.length }
    )

    if (page.length < DOWNLOAD_PAGE_SIZE) break
    from += DOWNLOAD_PAGE_SIZE

    // Yield to the browser between pages so the UI remains responsive.
    await sleep(DOWNLOAD_PAGE_PAUSE_MS)
  }

  return rows
}


const inFlightLoads = new Map()

export async function getCloudDatasetMeta(kind) {
  requireClient()
  const capability = await detectCloudSchema()
  const fields = capability.versioned
    ? 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at,active_version_id'
    : 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at'
  const { data, error } = await withRetry(async () => {
    const result = await supabase.from('iml_data_sources').select(fields).eq('kind', kind).maybeSingle()
    if (result.error) throw result.error
    return result
  }, `קריאת מטא-דאטה ${kind}`)
  if (error) throw error
  return data || null
}

export function loadCloudDatasetOnce(kind, onProgress) {
  if (inFlightLoads.has(kind)) return inFlightLoads.get(kind)
  const request = loadCloudDataset(kind, onProgress).finally(() => inFlightLoads.delete(kind))
  inFlightLoads.set(kind, request)
  return request
}

export async function loadCloudDataset(kind, onProgress) {
  requireClient()
  const capability = await detectCloudSchema()
  const sourceFields = capability.versioned
    ? 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at,active_version_id'
    : 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at'

  const source = await withRetry(async () => {
    const { data, error } = await supabase
      .from('iml_data_sources')
      .select(sourceFields)
      .eq('kind', kind)
      .maybeSingle()
    if (error) throw error
    return data
  }, `קריאת מקור ${kind}`)

  if (!source) return { rows: [], meta: null }

  if (capability.versioned && source.active_version_id) {
    const version = await withRetry(async () => {
      const { data, error } = await supabase
        .from('iml_dataset_versions')
        .select('id,version_no,chunk_count,status,created_at,activated_at')
        .eq('id', source.active_version_id)
        .single()
      if (error) throw error
      return data
    }, `קריאת גרסת ${kind}`)

    const rows = await readChunkPages({
      table: 'iml_dataset_chunks',
      filterColumn: 'version_id',
      filterValue: version.id,
      expectedChunks: Number(version.chunk_count || 0),
      onProgress,
      kind,
    })

    return {
      rows,
      meta: {
        fileName: source.file_name,
        rows: Number(source.row_count || rows.length),
        rawRows: Number(source.raw_row_count || 0),
        facilities: source.facilities,
        loadedAt: source.loaded_at || source.updated_at,
        loadedBy: source.loaded_by_email,
        valid: true,
        source: 'cloud',
        version: version.version_no,
        versionId: version.id,
      },
    }
  }

  const rows = await readChunkPages({
    table: 'iml_data_chunks',
    filterColumn: 'kind',
    filterValue: kind,
    onProgress,
    kind,
  })

  return {
    rows,
    meta: {
      fileName: source.file_name,
      rows: Number(source.row_count || rows.length),
      rawRows: Number(source.raw_row_count || 0),
      facilities: source.facilities,
      loadedAt: source.loaded_at || source.updated_at,
      loadedBy: source.loaded_by_email,
      valid: true,
      source: 'cloud',
      version: 0,
      legacy: true,
    },
  }
}

// Load the newest saved snapshot for every calendar month from the versioned
// dataset archive. This keeps historical quantity months available even when
// the active production upload contains only the current month-to-date file.
export async function loadCloudDatasetHistory(kind, { maxVersions = 60, maxMonths = 36 } = {}, onProgress) {
  requireClient()
  const capability = await detectCloudSchema()
  if (!capability.versioned) return loadCloudDataset(kind, onProgress)

  const versions = await withRetry(async () => {
    const { data, error } = await supabase
      .from('iml_dataset_versions')
      .select('id,version_no,chunk_count,status,created_at,activated_at')
      .eq('kind', kind)
      .in('status', ['active', 'archived'])
      .order('version_no', { ascending:false })
      .limit(maxVersions)
    if (error) throw error
    return data || []
  }, `קריאת היסטוריית גרסאות ${kind}`)

  const claimedMonths = new Set()
  const historyRows = []
  const usedVersions = []
  for (let index = 0; index < versions.length && claimedMonths.size < maxMonths; index += 1) {
    const version = versions[index]
    const rows = await readChunkPages({
      table:'iml_dataset_chunks', filterColumn:'version_id', filterValue:version.id,
      expectedChunks:Number(version.chunk_count || 0), kind:`${kind} — היסטוריה`,
    })
    const byMonth = new Map()
    rows.forEach(row => {
      const raw = row?.productionDay || row?.finishDate || row?.date || row?.Date || ''
      const text = raw instanceof Date ? raw.toISOString() : String(raw || '')
      const match = text.match(/(20\d{2})[-/.](\d{1,2})/)
      if (!match) return
      const month = `${match[1]}-${String(match[2]).padStart(2, '0')}`
      if (!byMonth.has(month)) byMonth.set(month, [])
      byMonth.get(month).push(row)
    })
    let used = false
    for (const [month, monthRows] of byMonth) {
      if (claimedMonths.has(month) || claimedMonths.size >= maxMonths) continue
      claimedMonths.add(month)
      historyRows.push(...monthRows)
      used = true
    }
    if (used) usedVersions.push(version.id)
    emit(onProgress, 'history', index + 1, versions.length || 1, `משחזר היסטוריית ${kind}`, { months:claimedMonths.size, downloadedRows:historyRows.length })
  }

  return {
    rows:historyRows,
    meta:{ source:'cloud-history', valid:true, rows:historyRows.length, months:[...claimedMonths].sort(), versions:usedVersions }
  }
}


async function readChunkPagesMatching({ table, filterColumn, filterValue, expectedChunks = 0, matcher, onProgress, kind }) {
  const matches = []
  let from = 0

  while (true) {
    const to = from + DOWNLOAD_PAGE_SIZE - 1
    const page = await withRetry(async () => {
      const { data, error } = await supabase
        .from(table)
        .select('chunk_index,payload')
        .eq(filterColumn, filterValue)
        .order('chunk_index', { ascending: true })
        .range(from, to)
      if (error) throw error
      return data || []
    }, `קריאת ${kind || 'נתונים'} מסוננים מהענן`)

    for (const chunk of page) {
      if (!Array.isArray(chunk.payload)) continue
      for (const row of chunk.payload) {
        if (!matcher || matcher(row)) matches.push(row)
      }
    }

    const completedChunks = Math.min(from + page.length, expectedChunks || from + page.length)
    emit(
      onProgress,
      'download-filtered',
      completedChunks,
      expectedChunks || Math.max(completedChunks, 1),
      `מחפש ${kind || 'נתונים'} עבור המנה שנבחרה`,
      { matchedRows: matches.length }
    )

    if (page.length < DOWNLOAD_PAGE_SIZE) break
    from += DOWNLOAD_PAGE_SIZE
    await sleep(DOWNLOAD_PAGE_PAUSE_MS)
  }

  return matches
}

export async function loadCloudDatasetMatching(kind, matcher, onProgress) {
  requireClient()
  const capability = await detectCloudSchema()
  const sourceFields = capability.versioned
    ? 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at,active_version_id'
    : 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at'

  const source = await withRetry(async () => {
    const { data, error } = await supabase
      .from('iml_data_sources')
      .select(sourceFields)
      .eq('kind', kind)
      .maybeSingle()
    if (error) throw error
    return data
  }, `קריאת מקור ${kind}`)

  if (!source) return { rows: [], meta: null }

  if (capability.versioned && source.active_version_id) {
    const version = await withRetry(async () => {
      const { data, error } = await supabase
        .from('iml_dataset_versions')
        .select('id,version_no,chunk_count,status,created_at,activated_at')
        .eq('id', source.active_version_id)
        .single()
      if (error) throw error
      return data
    }, `קריאת גרסת ${kind}`)

    const rows = await readChunkPagesMatching({
      table: 'iml_dataset_chunks',
      filterColumn: 'version_id',
      filterValue: version.id,
      expectedChunks: Number(version.chunk_count || 0),
      matcher,
      onProgress,
      kind,
    })

    return {
      rows,
      meta: {
        fileName: source.file_name,
        rows: Number(source.row_count || 0),
        rawRows: Number(source.raw_row_count || 0),
        facilities: source.facilities,
        loadedAt: source.loaded_at || source.updated_at,
        loadedBy: source.loaded_by_email,
        valid: true,
        source: 'cloud-filtered',
        version: version.version_no,
        versionId: version.id,
      },
    }
  }

  const rows = await readChunkPagesMatching({
    table: 'iml_data_chunks',
    filterColumn: 'kind',
    filterValue: kind,
    matcher,
    onProgress,
    kind,
  })
  return { rows, meta: { fileName: source.file_name, loadedAt: source.loaded_at || source.updated_at, source:'cloud-filtered', valid:true } }
}

/**
 * iPhone fast path: QUALITY is filtered inside PostgreSQL, before any rows are
 * transferred to Safari. Requires SPRINT_11_9_42_IPHONE_QUALITY_MONTH_RPC.sql.
 */
export async function loadCloudQualityMonth(month, onProgress) {
  requireClient()
  if (!/^\d{4}-\d{2}$/.test(String(month || ''))) throw new Error('Invalid quality month')

  const [yearText, monthText] = String(month).split('-')
  const year = Number(yearText), monthNumber = Number(monthText)
  const nextYear = monthNumber === 12 ? year + 1 : year
  const nextMonth = monthNumber === 12 ? 1 : monthNumber + 1
  const fromDate = `${yearText}-${monthText}-01`
  const toDate = `${nextYear}-${String(nextMonth).padStart(2, '0')}-01`

  emit(onProgress, 'quality-month', 0, 1, `מבקש מהענן רק את איכות ${month}`)
  const { data, error } = await supabase.rpc('iml_get_quality_month', {
    p_from_date: fromDate,
    p_to_date: toDate,
  })
  if (error) {
    const message = String(error.message || error.details || error.hint || '')
    if (/iml_get_quality_month|function.*does not exist|schema cache|PGRST202/i.test(message)) {
      throw new Error('מסלול איכות מהיר לאייפון עדיין לא מותקן ב-Supabase. יש להריץ פעם אחת את SPRINT_11_9_42_IPHONE_QUALITY_MONTH_RPC.sql')
    }
    throw error
  }

  const rows = (data || []).map(item => item?.payload ?? item).filter(Boolean)
  emit(onProgress, 'quality-month', 1, 1, `איכות ${month} ירדה מהענן — ${rows.length.toLocaleString()} רשומות`, { downloadedRows:rows.length })
  return { rows, meta:{ source:'cloud-quality-month', valid:true, mobileMonth:month, visibleRows:rows.length } }
}


// Sprint 11.9.43 — true monthly QUALITY cache for iPhone.
// Unlike loadCloudDatasetMatching(), this table is already partitioned by month,
// so iPhone never scans the full QUALITY dataset.
const MOBILE_QUALITY_ROWS_PER_CHUNK = 750

export async function getMobileQualityCacheMeta() {
  requireClient()
  const { data, error } = await supabase
    .from('iml_mobile_quality_cache_meta')
    .select('version_id,months,row_count,updated_at')
    .eq('id', 1)
    .maybeSingle()
  if (error) throw error
  return data || null
}

export async function rebuildMobileQualityCache(rows, versionId, onProgress) {
  requireClient()
  if (!Array.isArray(rows) || !rows.length || !versionId) return { rows:0, months:[] }

  const dated = []
  let maxMs = 0
  for (const row of rows) {
    const raw = row?.date
    const d = raw instanceof Date ? raw : new Date(raw)
    const ms = d?.getTime?.()
    if (!Number.isFinite(ms)) continue
    if (ms > maxMs) maxMs = ms
    dated.push({ row, d, ms })
  }
  if (!dated.length || !maxMs) return { rows:0, months:[] }

  // iPhone is intentionally a lightweight view. Keep the latest 3 QUALITY months
  // available in the dedicated cache; older history stays available on desktop.
  const latest = new Date(maxMs)
  const wantedMonths = []
  for (let offset = 0; offset < 3; offset += 1) {
    const d = new Date(latest.getFullYear(), latest.getMonth() - offset, 1)
    wantedMonths.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`)
  }
  const wanted = new Set(wantedMonths)

  const byMonth = new Map(wantedMonths.map(month => [month, []]))
  for (const { row, d } of dated) {
    const month = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`
    if (!wanted.has(month)) continue
    const serialized = { ...row, date: d.toISOString() }
    byMonth.get(month).push(serialized)
  }

  const { error: resetError } = await supabase.rpc('iml_reset_mobile_quality_cache', {
    p_version_id: versionId,
    p_months: wantedMonths,
  })
  if (resetError) throw resetError

  const chunks = []
  for (const month of wantedMonths) {
    const monthRows = byMonth.get(month) || []
    for (let i = 0; i < monthRows.length; i += MOBILE_QUALITY_ROWS_PER_CHUNK) {
      chunks.push({
        month_key: month,
        cache_version_id: versionId,
        chunk_index: Math.floor(i / MOBILE_QUALITY_ROWS_PER_CHUNK),
        payload: monthRows.slice(i, i + MOBILE_QUALITY_ROWS_PER_CHUNK),
        row_count: Math.min(MOBILE_QUALITY_ROWS_PER_CHUNK, monthRows.length - i),
      })
    }
  }

  let done = 0
  for (let i = 0; i < chunks.length; i += 3) {
    const batch = chunks.slice(i, i + 3)
    const { error } = await supabase
      .from('iml_mobile_quality_chunks')
      .upsert(batch, { onConflict:'month_key,cache_version_id,chunk_index' })
    if (error) throw error
    done += batch.length
    emit(onProgress, 'mobile-quality-cache', done, chunks.length || 1, 'מכין איכות מהירה לאייפון')
    await sleep(20)
  }

  const totalRows = wantedMonths.reduce((sum, month) => sum + (byMonth.get(month)?.length || 0), 0)
  const { error: completeError } = await supabase.rpc('iml_complete_mobile_quality_cache', {
    p_version_id: versionId,
    p_months: wantedMonths,
    p_row_count: totalRows,
  })
  if (completeError) throw completeError

  return { rows:totalRows, months:wantedMonths }
}

export async function loadMobileQualityMonth(month, onProgress) {
  requireClient()
  const meta = await getMobileQualityCacheMeta()
  if (!meta?.version_id) {
    throw new Error('מטמון האיכות לאייפון עדיין לא הוכן. יש לפתוח את IML CONTROL פעם אחת במחשב מנהל לאחר התקנת עדכון 11.9.43.')
  }
  const available = Array.isArray(meta.months) ? meta.months : []
  if (!available.includes(month)) {
    return { rows:[], meta:{ source:'mobile-quality-cache', valid:true, mobileMonth:month, availableMonths:available, cacheVersionId:meta.version_id } }
  }

  const rows = []
  let from = 0
  const pageSize = 20
  while (true) {
    const { data, error } = await supabase
      .from('iml_mobile_quality_chunks')
      .select('chunk_index,payload,row_count')
      .eq('cache_version_id', meta.version_id)
      .eq('month_key', month)
      .order('chunk_index', { ascending:true })
      .range(from, from + pageSize - 1)
    if (error) throw error
    const page = data || []
    for (const chunk of page) {
      if (Array.isArray(chunk.payload)) rows.push(...chunk.payload)
    }
    emit(onProgress, 'mobile-quality-month', rows.length, Math.max(Number(meta.row_count || rows.length), 1), `מוריד איכות ${month} לאייפון`, { downloadedRows:rows.length })
    if (page.length < pageSize) break
    from += pageSize
    await sleep(25)
  }

  return {
    rows,
    meta:{
      source:'mobile-quality-cache',
      valid:true,
      mobileMonth:month,
      visibleRows:rows.length,
      availableMonths:available,
      cacheVersionId:meta.version_id,
      loadedAt:meta.updated_at,
    }
  }
}


export async function loadAllCloudDatasets(onProgress) {
  const result = {}
  for (let index = 0; index < CLOUD_KINDS.length; index += 1) {
    const kind = CLOUD_KINDS[index]
    onProgress?.({ kind, phase: 'dataset', percent: Math.round((index / CLOUD_KINDS.length) * 100) })
    result[kind] = await loadCloudDatasetOnce(kind, progress => onProgress?.({ kind, ...progress }))
  }
  onProgress?.({ kind: '', phase: 'complete', percent: 100 })
  return result
}


export async function assertCloudWriteAccess() {
  requireClient()
  const { data: authData, error: authError } = await supabase.auth.getUser()
  if (authError) throw authError
  if (!authData?.user) throw new Error('אין משתמש מחובר ל־Supabase')

  const { data, error } = await supabase.rpc('iml_is_admin')
  if (error) throw error
  if (!data) {
    throw new Error(`למשתמש ${authData.user.email || ''} אין הרשאת מנהל ב־Supabase. יש לעדכן את role ל־admin בטבלת profiles.`)
  }
  return authData.user
}

export async function uploadCloudDataset(kind, rows, meta, user, onProgress) {
  requireClient()
  if (!CLOUD_KINDS.includes(kind)) throw new Error(`Unsupported dataset kind: ${kind}`)
  if (!Array.isArray(rows)) throw new Error('Dataset rows must be an array')

  await assertCloudWriteAccess()
  const capability = await detectCloudSchema(true)
  if (!capability.versioned) {
    throw new Error('מנוע הגרסאות עדיין לא מותקן ב־Supabase. יש להריץ את קובץ ההגירה ולאחר מכן לרענן את האתר.')
  }

  const startedAt = Date.now()
  const chunks = []
  for (let index = 0; index < rows.length; index += ROWS_PER_CHUNK) chunks.push(rows.slice(index, index + ROWS_PER_CHUNK))
  emit(onProgress, 'prepare', 1, 1, `הוכנו ${chunks.length} מקטעים קטנים ובטוחים`)

  const version = await withRetry(async () => {
    const { data, error } = await supabase
      .from('iml_dataset_versions')
      .insert({
        kind,
        file_name: meta.fileName || '',
        row_count: rows.length,
        raw_row_count: meta.rawRows ?? rows.length,
        facilities: meta.facilities || 0,
        chunk_count: chunks.length,
        uploaded_by: user?.id || null,
        uploaded_by_email: user?.email || '',
        status: 'uploading',
      })
      .select('id,version_no')
      .single()
    if (error) throw error
    return data
  }, 'יצירת גרסת נתונים')

  try {
    let uploadedChunks = 0
    const uploadJobs = []
    for (let offset = 0; offset < chunks.length; offset += CHUNKS_PER_UPLOAD_REQUEST) {
      uploadJobs.push({
        offset,
        batch: chunks.slice(offset, offset + CHUNKS_PER_UPLOAD_REQUEST).map((payload, batchIndex) => ({
          version_id: version.id,
          chunk_index: offset + batchIndex,
          payload,
          row_count: payload.length,
        })),
      })
    }
    for (let jobOffset = 0; jobOffset < uploadJobs.length; jobOffset += UPLOAD_CONCURRENCY) {
      const wave = uploadJobs.slice(jobOffset, jobOffset + UPLOAD_CONCURRENCY)
      await Promise.all(wave.map(async ({ offset, batch }) => {
        await withRetry(async () => {
          const { error } = await supabase.from('iml_dataset_chunks').upsert(batch, { onConflict:'version_id,chunk_index' })
          if (error) throw error
        }, `העלאת מקטעים ${offset + 1}-${offset + batch.length}`)
        uploadedChunks += batch.length
        emit(onProgress, 'upload', Math.min(uploadedChunks, chunks.length), chunks.length || 1, 'מעלה נתונים במקביל ל־Supabase')
      }))
      await sleep(DOWNLOAD_PAGE_PAUSE_MS)
    }

    emit(onProgress, 'verify', 0, 1, 'Supabase מאמת את הגרסה בצד השרת')
    await withRetry(async () => {
      // The activation function performs count/sum verification entirely in
      // PostgreSQL. We intentionally do not download thousands of row_count
      // records to the browser anymore.
      const { error } = await supabase.rpc('iml_activate_dataset_version', { p_version_id: version.id })
      if (error) throw error
    }, 'אימות והפעלת הגרסה')
    emit(onProgress, 'verify', 1, 1, 'האימות הסתיים בהצלחה')

    await supabase.from('iml_upload_history').insert({
      kind,
      file_name: meta.fileName || '',
      row_count: rows.length,
      raw_row_count: meta.rawRows ?? rows.length,
      uploaded_by: user?.id || null,
      uploaded_by_email: user?.email || '',
      status: 'success',
      version_id: version.id,
      duration_ms: Date.now() - startedAt,
    })

    emit(onProgress, 'complete', 1, 1, 'הגרסה הופעלה וזמינה לכל המשתמשים')
    return { ...meta, rows: rows.length, loadedAt: new Date().toISOString(), loadedBy: user?.email || '', source: 'cloud', version: version.version_no, versionId: version.id }
  } catch (error) {
    await supabase.from('iml_dataset_versions').update({ status: 'failed', error_message: error.message }).eq('id', version.id)
    await supabase.from('iml_upload_history').insert({
      kind,
      file_name: meta.fileName || '',
      row_count: rows.length,
      raw_row_count: meta.rawRows ?? rows.length,
      uploaded_by: user?.id || null,
      uploaded_by_email: user?.email || '',
      status: 'failed',
      error_message: error.message,
      version_id: version.id,
      duration_ms: Date.now() - startedAt,
    })
    throw error
  }
}

export async function loadUploadHistory(limit = 25) {
  requireClient()
  const { data, error } = await supabase
    .from('iml_upload_history')
    .select('id,kind,file_name,row_count,raw_row_count,uploaded_by_email,status,error_message,created_at,version_id,duration_ms')
    .order('created_at', { ascending: false })
    .limit(limit)
  if (error) throw error
  return data || []
}

export async function getCloudHealth() {
  requireClient()
  const started = Date.now()
  const capability = await detectCloudSchema()
  const { data, error } = await supabase.from('iml_data_sources').select('kind').limit(10)
  if (error) throw error
  return {
    online: true,
    datasets: data?.length || 0,
    latencyMs: Date.now() - started,
    checkedAt: new Date().toISOString(),
    versioned: capability.versioned,
    schemaMessage: capability.message,
  }
}

export async function deleteAllCloudDatasets(user) {
  requireClient()
  const { error } = await supabase.rpc('iml_delete_all_datasets')
  if (error) throw error
  await supabase.from('iml_upload_history').insert({
    kind: 'all', file_name: '', row_count: 0, raw_row_count: 0,
    uploaded_by: user?.id || null, uploaded_by_email: user?.email || '', status: 'deleted',
  })
}


export async function uploadCloudDatasetIncremental(kind, newRows, meta, user, onProgress) {
  requireClient()
  if (kind !== 'quality') throw new Error('טעינה מצטברת נתמכת כעת עבור איכות בלבד')
  if (!Array.isArray(newRows) || !newRows.length) throw new Error('לא נמצאו רשומות חדשות להעלאה')
  await assertCloudWriteAccess()
  const startedAt = Date.now()
  const chunks = []
  for (let index = 0; index < newRows.length; index += ROWS_PER_CHUNK) chunks.push(newRows.slice(index, index + ROWS_PER_CHUNK))
  emit(onProgress, 'prepare', 1, 1, `נמצאו ${newRows.length} רשומות חדשות בלבד`)

  // Build 2: do NOT clone the complete active quality dataset in one PostgreSQL
  // statement. With 800K+ quality rows that single INSERT ... SELECT exceeded
  // Supabase's statement timeout (57014). We now create the new version first
  // and copy the old chunks server-side in small bounded pages.
  const { data: seed, error: seedError } = await supabase.rpc('iml_prepare_incremental_dataset_version', {
    p_kind: kind,
    p_file_name: meta.fileName || '',
    p_raw_row_count: meta.rawRows ?? newRows.length,
    p_facilities: meta.facilities || 0,
    p_new_row_count: newRows.length,
    p_new_chunk_count: chunks.length,
    p_uploaded_by: user?.id || null,
    p_uploaded_by_email: user?.email || '',
  })
  if (seedError) {
    const message = `${seedError.message || ''} ${seedError.details || ''} ${seedError.hint || ''}`
    if (/iml_prepare_incremental_dataset_version|schema cache|could not find|does not exist/i.test(message)) {
      throw new Error('מנגנון העלאת איכות מהירה עדיין לא מותקן ב-Supabase. יש להריץ פעם אחת את SPRINT_11_9_1_QUALITY_INCREMENTAL_TIMEOUT_FIX.sql.')
    }
    throw seedError
  }

  const version = Array.isArray(seed) ? seed[0] : seed
  if (!version?.version_id) throw new Error('Supabase לא החזיר מזהה לגרסת האיכות המצטברת')
  const baseChunkIndex = Number(version.base_chunk_count) || 0
  const previousVersionId = version.previous_version_id || null

  try {
    if (previousVersionId && baseChunkIndex > 0) {
      emit(onProgress, 'copy', 0, baseChunkIndex, 'מעתיק את נתוני האיכות הקיימים במנות קטנות')
      for (let fromIndex = 0; fromIndex < baseChunkIndex; fromIndex += CHUNKS_PER_SERVER_COPY_REQUEST) {
        const toIndex = Math.min(baseChunkIndex - 1, fromIndex + CHUNKS_PER_SERVER_COPY_REQUEST - 1)
        await withRetry(async () => {
          const { error } = await supabase.rpc('iml_copy_dataset_chunks_range', {
            p_source_version_id: previousVersionId,
            p_target_version_id: version.version_id,
            p_from_chunk_index: fromIndex,
            p_to_chunk_index: toIndex,
          })
          if (error) throw error
        }, `העתקת איכות קיימת ${fromIndex + 1}-${toIndex + 1}`)
        emit(onProgress, 'copy', toIndex + 1, baseChunkIndex, 'מעתיק את נתוני האיכות הקיימים במנות קטנות')
        await sleep(DOWNLOAD_PAGE_PAUSE_MS)
      }
    }

    let uploadedChunks = 0
    const uploadJobs = []
    for (let offset = 0; offset < chunks.length; offset += CHUNKS_PER_UPLOAD_REQUEST) {
      uploadJobs.push({
        offset,
        batch: chunks.slice(offset, offset + CHUNKS_PER_UPLOAD_REQUEST).map((payload, batchIndex) => ({
          version_id: version.version_id,
          chunk_index: baseChunkIndex + offset + batchIndex,
          payload,
          row_count: payload.length,
        })),
      })
    }
    for (let jobOffset = 0; jobOffset < uploadJobs.length; jobOffset += UPLOAD_CONCURRENCY) {
      const wave = uploadJobs.slice(jobOffset, jobOffset + UPLOAD_CONCURRENCY)
      await Promise.all(wave.map(async ({ offset, batch }) => {
        await withRetry(async () => {
          const { error } = await supabase.from('iml_dataset_chunks').upsert(batch, { onConflict:'version_id,chunk_index' })
          if (error) throw error
        }, `העלאת רשומות איכות חדשות ${offset + 1}-${offset + batch.length}`)
        uploadedChunks += batch.length
        emit(onProgress, 'upload', Math.min(uploadedChunks, chunks.length), chunks.length || 1, 'מעלה רשומות איכות חדשות במקביל')
      }))
      await sleep(DOWNLOAD_PAGE_PAUSE_MS)
    }

    emit(onProgress, 'verify', 0, 1, 'מאמת ומפעיל את גרסת האיכות המצטברת')
    await withRetry(async () => {
      const { error } = await supabase.rpc('iml_activate_dataset_version', { p_version_id:version.version_id })
      if (error) throw error
    }, 'אימות גרסת האיכות')
    emit(onProgress, 'verify', 1, 1, 'האימות הסתיים בהצלחה')

    const totalRows = Number(version.total_row_count) || ((Number(meta.existingRows) || 0) + newRows.length)
    await supabase.from('iml_upload_history').insert({
      kind, file_name:meta.fileName || '', row_count:newRows.length, raw_row_count:meta.rawRows ?? newRows.length,
      uploaded_by:user?.id || null, uploaded_by_email:user?.email || '', status:'success', version_id:version.version_id,
      duration_ms:Date.now() - startedAt,
    })
    emit(onProgress, 'complete', 1, 1, `נוספו ${newRows.length} רשומות חדשות; סה״כ ${totalRows}`)
    return { ...meta, rows:totalRows, newRows:newRows.length, loadedAt:new Date().toISOString(), loadedBy:user?.email || '', source:'cloud', version:version.version_no, versionId:version.version_id }
  } catch (error) {
    await supabase.from('iml_dataset_versions').update({ status:'failed', error_message:error.message }).eq('id', version.version_id)
    throw error
  }
}


// Sprint 11.9.2 — keep the original monthly-target workbook as a shared cloud asset.
// The normalized target rows remain the source for calculations; this singleton copy
// preserves the exact Excel workbook so every computer downloads the same template.
function arrayBufferToBase64(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer)
  let binary = ''
  const step = 0x8000
  for (let offset = 0; offset < bytes.length; offset += step) {
    binary += String.fromCharCode(...bytes.subarray(offset, Math.min(offset + step, bytes.length)))
  }
  return btoa(binary)
}

function base64ToUint8Array(base64) {
  const binary = atob(base64 || '')
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}

function targetWorkbookSchemaError(error) {
  if (isMissingSchema(error, ['iml_target_workbook'])) {
    return new Error('סנכרון קובץ היעדים המלא עדיין לא מותקן ב-Supabase. יש להריץ SPRINT_11_9_2_TARGET_WORKBOOK_SYNC.sql פעם אחת.')
  }
  return error
}

export async function saveActiveTargetWorkbook(workbook, user = null) {
  requireClient()
  if (!workbook?.bytes) throw new Error('קובץ היעדים המקורי חסר')
  const bytes = workbook.bytes instanceof Uint8Array ? workbook.bytes : new Uint8Array(workbook.bytes)
  const row = {
    id: 1,
    file_name: workbook.name || 'IML_Monthly_Targets.xlsx',
    mime_type: workbook.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    byte_count: bytes.byteLength,
    bytes_base64: arrayBufferToBase64(bytes),
    last_modified: workbook.lastModified ? new Date(workbook.lastModified).toISOString() : null,
    target_version_id: workbook.versionId || null,
    uploaded_by: user?.id || null,
    uploaded_by_email: user?.email || '',
    updated_at: new Date().toISOString(),
  }
  const { data, error } = await supabase
    .from('iml_target_workbook')
    .upsert(row, { onConflict:'id' })
    .select('id,file_name,mime_type,byte_count,last_modified,target_version_id,uploaded_by_email,updated_at')
    .single()
  if (error) throw targetWorkbookSchemaError(error)
  return data
}

export async function loadActiveTargetWorkbook() {
  requireClient()
  const { data, error } = await supabase
    .from('iml_target_workbook')
    .select('id,file_name,mime_type,byte_count,bytes_base64,last_modified,target_version_id,uploaded_by_email,updated_at')
    .eq('id', 1)
    .maybeSingle()
  if (error) throw targetWorkbookSchemaError(error)
  if (!data?.bytes_base64) return null
  const bytes = base64ToUint8Array(data.bytes_base64)
  return {
    name: data.file_name,
    type: data.mime_type,
    lastModified: data.last_modified ? new Date(data.last_modified).getTime() : 0,
    savedAt: data.updated_at,
    targetVersionId: data.target_version_id || '',
    uploadedBy: data.uploaded_by_email || '',
    byteCount: Number(data.byte_count || bytes.byteLength),
    bytes,
    source: 'cloud',
  }
}


// Sprint 11.9.34 — monthly target history.
// Unlike the legacy `targets` dataset (which keeps one active version), this
// archive stores one normalized target snapshot per planning month.
function monthlyTargetsSchemaError(error) {
  if (isMissingSchema(error, ['iml_monthly_targets', 'iml_monthly_target_workbooks'])) {
    return new Error('היסטוריית היעדים החודשית עדיין לא מותקנת ב-Supabase. יש להריץ SPRINT_11_9_34_MONTHLY_TARGET_HISTORY.sql פעם אחת.')
  }
  return error
}

export async function saveMonthlyTargetDataset(month, rows, meta = {}, user = null) {
  requireClient()
  if (!/^\d{4}-\d{2}$/.test(String(month || ''))) throw new Error('חודש היעדים אינו תקין')
  if (!Array.isArray(rows) || !rows.length) throw new Error('לא נמצאו יעדים לשמירה חודשית')
  await assertCloudWriteAccess()
  const row = {
    target_month: month,
    rows_json: rows,
    row_count: rows.length,
    file_name: meta.originalFileName || meta.fileName || '',
    target_version_id: meta.versionId || null,
    uploaded_by: user?.id || null,
    uploaded_by_email: user?.email || '',
    updated_at: new Date().toISOString(),
  }
  const { data, error } = await supabase
    .from('iml_monthly_targets')
    .upsert(row, { onConflict:'target_month' })
    .select('target_month,row_count,file_name,target_version_id,uploaded_by_email,updated_at')
    .single()
  if (error) throw monthlyTargetsSchemaError(error)
  return data
}

export async function loadAllMonthlyTargetDatasets() {
  requireClient()
  const { data, error } = await supabase
    .from('iml_monthly_targets')
    .select('target_month,rows_json,row_count,file_name,target_version_id,uploaded_by_email,updated_at')
    .order('target_month', { ascending:true })
  if (error) throw monthlyTargetsSchemaError(error)
  const months = data || []
  const rows = months.flatMap(item => Array.isArray(item.rows_json) ? item.rows_json : [])
  const latest = months.at(-1) || null
  return {
    rows,
    months: months.map(item => ({
      month:item.target_month, rows:Number(item.row_count || 0), fileName:item.file_name || '',
      versionId:item.target_version_id || '', loadedBy:item.uploaded_by_email || '', loadedAt:item.updated_at || '',
    })),
    meta: latest ? {
      fileName: latest.file_name || `יעדים ${latest.target_month}`,
      rows: Number(latest.row_count || 0), loadedAt:latest.updated_at, loadedBy:latest.uploaded_by_email || '',
      versionId:latest.target_version_id || '', source:'cloud-monthly', valid:true,
    } : null,
  }
}

export async function saveMonthlyTargetWorkbook(month, workbook, user = null) {
  requireClient()
  if (!/^\d{4}-\d{2}$/.test(String(month || ''))) throw new Error('חודש היעדים אינו תקין')
  if (!workbook?.bytes) throw new Error('קובץ היעדים המקורי חסר')
  await assertCloudWriteAccess()
  const bytes = workbook.bytes instanceof Uint8Array ? workbook.bytes : new Uint8Array(workbook.bytes)
  const row = {
    target_month:month,
    file_name:workbook.name || `IML_Monthly_Targets_${month}.xlsx`,
    mime_type:workbook.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    byte_count:bytes.byteLength,
    bytes_base64:arrayBufferToBase64(bytes),
    last_modified:workbook.lastModified ? new Date(workbook.lastModified).toISOString() : null,
    target_version_id:workbook.versionId || null,
    uploaded_by:user?.id || null,
    uploaded_by_email:user?.email || '',
    updated_at:new Date().toISOString(),
  }
  const { data, error } = await supabase
    .from('iml_monthly_target_workbooks')
    .upsert(row, { onConflict:'target_month' })
    .select('target_month,file_name,mime_type,byte_count,last_modified,target_version_id,uploaded_by_email,updated_at')
    .single()
  if (error) throw monthlyTargetsSchemaError(error)
  return data
}

export async function loadMonthlyTargetWorkbook(month) {
  requireClient()
  if (!month) return null
  const { data, error } = await supabase
    .from('iml_monthly_target_workbooks')
    .select('target_month,file_name,mime_type,byte_count,bytes_base64,last_modified,target_version_id,uploaded_by_email,updated_at')
    .eq('target_month', month)
    .maybeSingle()
  if (error) throw monthlyTargetsSchemaError(error)
  if (!data?.bytes_base64) return null
  const bytes = base64ToUint8Array(data.bytes_base64)
  return {
    month:data.target_month,
    name:data.file_name,
    type:data.mime_type,
    lastModified:data.last_modified ? new Date(data.last_modified).getTime() : 0,
    savedAt:data.updated_at,
    targetVersionId:data.target_version_id || '',
    uploadedBy:data.uploaded_by_email || '',
    byteCount:Number(data.byte_count || bytes.byteLength),
    bytes,
    source:'cloud-monthly',
  }
}
