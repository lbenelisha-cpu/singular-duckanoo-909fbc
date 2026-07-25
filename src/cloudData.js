import { supabase } from './supabase'

export const CLOUD_KINDS = ['production', 'quality', 'deviations', 'targets']
export const FUTURE_CLOUD_KINDS = [...CLOUD_KINDS, 'packaging_plan']
const CHUNK_SIZE = 500
const INSERT_BATCH_SIZE = 12

function requireClient() {
  if (!supabase) throw new Error('Supabase client is not configured')
}

export async function loadCloudDataset(kind) {
  requireClient()
  const { data: source, error: sourceError } = await supabase
    .from('iml_data_sources')
    .select('kind,file_name,row_count,raw_row_count,facilities,loaded_at,loaded_by_email,updated_at')
    .eq('kind', kind)
    .maybeSingle()
  if (sourceError) throw sourceError
  if (!source) return { rows: [], meta: null }

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

  return {
    rows: chunks.flatMap(chunk => Array.isArray(chunk.payload) ? chunk.payload : []),
    meta: {
      fileName: source.file_name,
      rows: source.row_count,
      rawRows: source.raw_row_count,
      facilities: source.facilities,
      loadedAt: source.loaded_at || source.updated_at,
      loadedBy: source.loaded_by_email,
      valid: true,
      source: 'cloud',
    },
  }
}

export async function loadAllCloudDatasets(onProgress) {
  const result = {}
  for (const kind of CLOUD_KINDS) {
    onProgress?.(kind)
    result[kind] = await loadCloudDataset(kind)
  }
  return result
}

export async function uploadCloudDataset(kind, rows, meta, user) {
  requireClient()
  if (!CLOUD_KINDS.includes(kind)) throw new Error(`Unsupported dataset kind: ${kind}`)

  const chunks = []
  for (let index = 0; index < rows.length; index += CHUNK_SIZE) {
    chunks.push({
      kind,
      chunk_index: Math.floor(index / CHUNK_SIZE),
      payload: rows.slice(index, index + CHUNK_SIZE),
    })
  }

  const cloudMeta = {
    kind,
    file_name: meta.fileName || '',
    row_count: rows.length,
    raw_row_count: meta.rawRows ?? rows.length,
    facilities: meta.facilities || 0,
    loaded_at: meta.loadedAt || new Date().toISOString(),
    loaded_by: user?.id || null,
    loaded_by_email: user?.email || '',
    updated_at: new Date().toISOString(),
  }
  const { error: sourceError } = await supabase.from('iml_data_sources').upsert(cloudMeta, { onConflict: 'kind' })
  if (sourceError) throw sourceError

  const { error: deleteError } = await supabase.from('iml_data_chunks').delete().eq('kind', kind)
  if (deleteError) throw deleteError

  for (let index = 0; index < chunks.length; index += INSERT_BATCH_SIZE) {
    const { error } = await supabase.from('iml_data_chunks').insert(chunks.slice(index, index + INSERT_BATCH_SIZE))
    if (error) throw error
  }

  await supabase.from('iml_upload_history').insert({
    kind,
    file_name: cloudMeta.file_name,
    row_count: cloudMeta.row_count,
    raw_row_count: cloudMeta.raw_row_count,
    uploaded_by: cloudMeta.loaded_by,
    uploaded_by_email: cloudMeta.loaded_by_email,
    status: 'success',
  })

  return { ...meta, rows: rows.length, loadedBy: user?.email || '', source: 'cloud' }
}

export async function loadUploadHistory(limit = 25) {
  requireClient()
  const { data, error } = await supabase
    .from('iml_upload_history')
    .select('id,kind,file_name,row_count,raw_row_count,uploaded_by_email,status,error_message,created_at')
    .order('created_at', { ascending: false })
    .limit(limit)
  if (error) throw error
  return data || []
}

export async function getCloudHealth() {
  requireClient()
  const started = Date.now()
  const { count, error } = await supabase
    .from('iml_data_sources')
    .select('*', { count: 'exact', head: true })
  if (error) throw error
  return { online: true, datasets: count || 0, latencyMs: Date.now() - started, checkedAt: new Date().toISOString() }
}

export async function deleteAllCloudDatasets(user) {
  requireClient()
  const { error: chunksError } = await supabase.from('iml_data_chunks').delete().in('kind', CLOUD_KINDS)
  if (chunksError) throw chunksError
  const { error: sourcesError } = await supabase.from('iml_data_sources').delete().in('kind', CLOUD_KINDS)
  if (sourcesError) throw sourcesError
  await supabase.from('iml_upload_history').insert({
    kind: 'all', file_name: '', row_count: 0, raw_row_count: 0,
    uploaded_by: user?.id || null, uploaded_by_email: user?.email || '', status: 'deleted',
  })
}
