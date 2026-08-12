-- Sprint 11.9.1
-- Quality incremental upload timeout fix (PostgreSQL error 57014)
--
-- Problem:
-- The previous incremental function cloned ALL active quality chunks in one
-- INSERT ... SELECT statement. With 800K+ quality rows this can exceed the
-- Supabase/PostgREST statement timeout.
--
-- Fix:
-- 1. Prepare an empty new version and return the previous version id.
-- 2. Copy old chunks in small server-side ranges from the browser.
-- 3. Upload only new chunks and activate with the existing verification RPC.

create or replace function public.iml_prepare_incremental_dataset_version(
  p_kind text,
  p_file_name text,
  p_raw_row_count bigint,
  p_facilities integer,
  p_new_row_count bigint,
  p_new_chunk_count integer,
  p_uploaded_by uuid,
  p_uploaded_by_email text
)
returns table(
  version_id uuid,
  version_no bigint,
  previous_version_id uuid,
  base_chunk_count integer,
  total_row_count bigint
)
language plpgsql
security definer
set search_path = public
as $$
declare
  previous_version public.iml_dataset_versions%rowtype;
  created_version public.iml_dataset_versions%rowtype;
begin
  if not public.iml_is_admin() then
    raise exception 'Admin permission required';
  end if;

  if p_kind <> 'quality' then
    raise exception 'Incremental upload is currently supported for quality only';
  end if;

  if coalesce(p_new_row_count, 0) <= 0 then
    raise exception 'No new rows were supplied';
  end if;

  select * into previous_version
  from public.iml_dataset_versions
  where kind = p_kind and status = 'active'
  order by activated_at desc nulls last, created_at desc
  limit 1;

  insert into public.iml_dataset_versions(
    kind, version_no, file_name, row_count, raw_row_count, facilities,
    chunk_count, uploaded_by, uploaded_by_email, status
  ) values (
    p_kind,
    0,
    coalesce(p_file_name, ''),
    coalesce(previous_version.row_count, 0) + p_new_row_count,
    coalesce(previous_version.raw_row_count, 0) + coalesce(p_raw_row_count, p_new_row_count),
    greatest(coalesce(previous_version.facilities, 0), coalesce(p_facilities, 0)),
    coalesce(previous_version.chunk_count, 0) + coalesce(p_new_chunk_count, 0),
    p_uploaded_by,
    coalesce(p_uploaded_by_email, ''),
    'uploading'
  ) returning * into created_version;

  return query select
    created_version.id,
    created_version.version_no,
    previous_version.id,
    coalesce(previous_version.chunk_count, 0),
    created_version.row_count;
end;
$$;

grant execute on function public.iml_prepare_incremental_dataset_version(
  text,text,bigint,integer,bigint,integer,uuid,text
) to authenticated;

create or replace function public.iml_copy_dataset_chunks_range(
  p_source_version_id uuid,
  p_target_version_id uuid,
  p_from_chunk_index integer,
  p_to_chunk_index integer
)
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  source_version public.iml_dataset_versions%rowtype;
  target_version public.iml_dataset_versions%rowtype;
  copied integer := 0;
begin
  if not public.iml_is_admin() then
    raise exception 'Admin permission required';
  end if;

  if p_from_chunk_index < 0 or p_to_chunk_index < p_from_chunk_index then
    raise exception 'Invalid chunk range';
  end if;

  select * into source_version
  from public.iml_dataset_versions
  where id = p_source_version_id;

  if not found then
    raise exception 'Source dataset version not found';
  end if;

  select * into target_version
  from public.iml_dataset_versions
  where id = p_target_version_id;

  if not found then
    raise exception 'Target dataset version not found';
  end if;

  if target_version.status <> 'uploading' then
    raise exception 'Target dataset version is not awaiting upload';
  end if;

  if source_version.kind <> target_version.kind then
    raise exception 'Source and target dataset kinds do not match';
  end if;

  insert into public.iml_dataset_chunks(version_id, chunk_index, payload, row_count)
  select p_target_version_id, chunk_index, payload, row_count
  from public.iml_dataset_chunks
  where version_id = p_source_version_id
    and chunk_index between p_from_chunk_index and p_to_chunk_index
  order by chunk_index
  on conflict (version_id, chunk_index) do update
    set payload = excluded.payload,
        row_count = excluded.row_count;

  get diagnostics copied = row_count;
  return copied;
end;
$$;

grant execute on function public.iml_copy_dataset_chunks_range(
  uuid,uuid,integer,integer
) to authenticated;

select 'Sprint 11.9.1 quality incremental timeout fix installed' as status;
