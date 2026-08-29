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
