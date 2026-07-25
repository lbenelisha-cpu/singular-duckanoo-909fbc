import { supabase } from './supabase'

export const CLOUD_KINDS = ['production', 'quality', 'deviations', 'targets']
export const FUTURE_CLOUD_KINDS = [...CLOUD_KINDS, 'packaging_plan']
const ROWS_PER_CHUNK = 350
const CHUNKS_PER_REQUEST = 8

let schemaCapability = null

function requireClient() {
  if (!supabase) throw new Error('Supabase client is not configured')
}

function emit(onProgress, phase, completed, total, message) {
  const percent = total > 0 ? Math.min(100, Math.round((completed / total) * 100)) : 0
  onProgress?.({ phase, completed, total, percent, message })
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
      message: versionError ? 'טבלת גרסאות הנתונים עדיין אינה מותקנת' : 'סכמת Sprint 10.1.1 פעילה',
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
        message: 'מבנה הענן הישן פעיל; יש להריץ את קובץ ההגירה של Sprint 10.1.1',
        error: sourceError,
      }
      return schemaCapability
    }
  }

  throw sourceError
}

async function readVersionChunks(versionId, expectedChunks = 0, onProgress) {
  const chunks = []
  const pageSize = 500
  for (let from = 0; ; from += pageSize) {
    const { data, error } = await supabase
      .from('iml_dataset_chunks')
      .select('chunk_index,payload')
      .eq('version_id', versionId)
      .order('chunk_index', { ascending: true })
      .range(from, from + pageSize - 1)
    if (error) throw error
    chunks.push(...(data || []))
    emit(onProgress, 'download', chunks.length, expectedChunks || Math.max(chunks.length, 1), 'מוריד נתונים מהענן')
    if (!data || data.length < pageSize) break
  }
  return chunks
}

async function readLegacyChunks(kind) {
  const chunks = []
  const pageSize = 500
  for (let from = 0; ; from += pageSize) {
    const { data, error } = await supabase
      .from('iml_data_chunks')
      .select('chunk_index,payload')
      .eq('kind', kind)
      .order('chunk_index', { ascending: true })
      .range(from, from + pageSize - 1)
    if (error) throw error
    chunks.push(...(data || []))
    if (!data || data.length < pageSize) break
  }
  return chunks
}

export async function loadCloudDataset(kind, onProgress) {
  requireClient()
  const capability = await detectCloudSchema()
  const sourceFields = capability.versioned
    ? 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at,active_version_id'
    : 'kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at'

  const { data: source, error: sourceError } = await supabase
    .from('iml_data_sources')
    .select(sourceFields)
    .eq('kind', kind)
    .maybeSingle()
  if (sourceError) throw sourceError
  if (!source) return { rows: [], meta: null }

  if (capability.versioned && source.active_version_id) {
    const { data: version, error: versionError } = await supabase
      .from('iml_dataset_versions')
      .select('id,version_no,chunk_count,status,created_at,activated_at')
      .eq('id', source.active_version_id)
      .single()
    if (versionError) throw versionError
    const chunks = await readVersionChunks(version.id, version.chunk_count, onProgress)
    return {
      rows: chunks.flatMap(chunk => Array.isArray(chunk.payload) ? chunk.payload : []),
      meta: {
        fileName: source.file_name,
        rows: Number(source.row_count || 0),
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

  const chunks = await readLegacyChunks(kind)
  return {
    rows: chunks.flatMap(chunk => Array.isArray(chunk.payload) ? chunk.payload : []),
    meta: {
      fileName: source.file_name,
      rows: Number(source.row_count || 0),
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
    result[kind] = await loadCloudDataset(kind)
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
    throw new Error('מנוע הגרסאות עדיין לא מותקן ב־Supabase. יש להריץ את SPRINT_10_1_1_MIGRATION.sql ולאחר מכן לרענן את האתר.')
  }

  const startedAt = Date.now()
  const chunks = []
  for (let index = 0; index < rows.length; index += ROWS_PER_CHUNK) chunks.push(rows.slice(index, index + ROWS_PER_CHUNK))
  emit(onProgress, 'prepare', 1, 1, `הוכנו ${chunks.length} מקטעים`)

  const { data: version, error: versionError } = await supabase
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
  if (versionError) throw versionError

  try {
    for (let offset = 0; offset < chunks.length; offset += CHUNKS_PER_REQUEST) {
      const batch = chunks.slice(offset, offset + CHUNKS_PER_REQUEST).map((payload, batchIndex) => ({
        version_id: version.id,
        chunk_index: offset + batchIndex,
        payload,
        row_count: payload.length,
      }))
      const { error } = await supabase.from('iml_dataset_chunks').insert(batch)
      if (error) throw error
      emit(onProgress, 'upload', Math.min(offset + batch.length, chunks.length), chunks.length || 1, 'מעלה מקטעים ל־Supabase')
    }

    emit(onProgress, 'verify', 0, 1, 'מאמת את הטעינה לפני הפעלה')
    const { data: verification, error: verifyError } = await supabase
      .from('iml_dataset_chunks')
      .select('row_count')
      .eq('version_id', version.id)
    if (verifyError) throw verifyError
    const verifiedRows = (verification || []).reduce((sum, item) => sum + Number(item.row_count || 0), 0)
    if (verifiedRows !== rows.length || (verification || []).length !== chunks.length) {
      throw new Error(`אימות הטעינה נכשל: נשמרו ${verifiedRows} מתוך ${rows.length} רשומות`)
    }
    emit(onProgress, 'verify', 1, 1, 'האימות הסתיים בהצלחה')

    const { error: activateError } = await supabase.rpc('iml_activate_dataset_version', { p_version_id: version.id })
    if (activateError) throw activateError

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
