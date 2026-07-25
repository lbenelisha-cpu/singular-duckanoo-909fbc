import { supabase } from './supabase'

export const CLOUD_KINDS = ['production', 'quality', 'deviations', 'targets']
export const FUTURE_CLOUD_KINDS = [...CLOUD_KINDS, 'packaging_plan']

// Keep every request deliberately small. A quality file can contain hundreds of
// thousands of rows, and large JSONB responses are what caused the database
// statement timeouts in the previous build.
const ROWS_PER_CHUNK = 150
const CHUNKS_PER_UPLOAD_REQUEST = 3
const CHUNKS_PER_DOWNLOAD_PAGE = 12
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

  const { error: sourceError } = await supabase
    .from('iml_data_sources')
    .select('kind,active_version_id')
    .limit(1)

  if (!sourceError) {
    const { error: versionError } = await supabase
      .from('iml_dataset_versions')
      .select('id')
      .limit(1)
    schemaCapability = {
      legacy: true,
      versioned: !versionError,
      message: versionError ? 'טבלת גרסאות הנתונים עדיין אינה מותקנת' : 'סכמת Sprint 10.2 פעילה',
      error: versionError || null,
    }
    return schemaCapability
  }

  if (isMissingSchema(sourceError, ['active_version_id'])) {
    const { error: legacyError } = await supabase.from('iml_data_sources').select('kind').limit(1)
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
    const to = from + CHUNKS_PER_DOWNLOAD_PAGE - 1
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

    if (page.length < CHUNKS_PER_DOWNLOAD_PAGE) break
    from += CHUNKS_PER_DOWNLOAD_PAGE

    // Yield to the browser between pages so the UI remains responsive.
    await sleep(0)
  }

  return rows
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

export async function loadAllCloudDatasets(onProgress) {
  const result = {}
  for (let index = 0; index < CLOUD_KINDS.length; index += 1) {
    const kind = CLOUD_KINDS[index]
    onProgress?.({ kind, phase: 'dataset', percent: Math.round((index / CLOUD_KINDS.length) * 100) })
    result[kind] = await loadCloudDataset(kind, progress => onProgress?.({ kind, ...progress }))
  }
  onProgress?.({ kind: '', phase: 'complete', percent: 100 })
  return result
}

export async function uploadCloudDataset(kind, rows, meta, user, onProgress) {
  requireClient()
  if (!CLOUD_KINDS.includes(kind)) throw new Error(`Unsupported dataset kind: ${kind}`)
  if (!Array.isArray(rows)) throw new Error('Dataset rows must be an array')

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
    for (let offset = 0; offset < chunks.length; offset += CHUNKS_PER_UPLOAD_REQUEST) {
      const batch = chunks.slice(offset, offset + CHUNKS_PER_UPLOAD_REQUEST).map((payload, batchIndex) => ({
        version_id: version.id,
        chunk_index: offset + batchIndex,
        payload,
        row_count: payload.length,
      }))

      await withRetry(async () => {
        // Upsert makes a retried request idempotent if the network response was
        // lost after Supabase had already committed the batch.
        const { error } = await supabase
          .from('iml_dataset_chunks')
          .upsert(batch, { onConflict: 'version_id,chunk_index' })
        if (error) throw error
      }, `העלאת מקטעים ${offset + 1}-${offset + batch.length}`)

      emit(onProgress, 'upload', Math.min(offset + batch.length, chunks.length), chunks.length || 1, 'מעלה מקטעים קטנים ל־Supabase')
      await sleep(0)
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
