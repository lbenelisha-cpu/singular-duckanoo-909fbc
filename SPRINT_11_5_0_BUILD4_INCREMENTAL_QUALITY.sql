-- Sprint 11.5.0 Build 4
-- Incremental quality uploads: copies the active quality chunks inside PostgreSQL
-- and lets the browser upload only rows that are not already present.

create or replace function public.iml_create_incremental_dataset_version(
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
    p_kind, 0, coalesce(p_file_name, ''),
    coalesce(previous_version.row_count, 0) + p_new_row_count,
    coalesce(previous_version.raw_row_count, 0) + coalesce(p_raw_row_count, p_new_row_count),
    greatest(coalesce(previous_version.facilities, 0), coalesce(p_facilities, 0)),
    coalesce(previous_version.chunk_count, 0) + coalesce(p_new_chunk_count, 0),
    p_uploaded_by, coalesce(p_uploaded_by_email, ''), 'uploading'
  ) returning * into created_version;

  if previous_version.id is not null then
    insert into public.iml_dataset_chunks(version_id, chunk_index, payload, row_count)
    select created_version.id, chunk_index, payload, row_count
    from public.iml_dataset_chunks
    where version_id = previous_version.id
    order by chunk_index;
  end if;

  return query select
    created_version.id,
    created_version.version_no,
    coalesce(previous_version.chunk_count, 0),
    created_version.row_count;
end;
$$;

grant execute on function public.iml_create_incremental_dataset_version(
  text,text,bigint,integer,bigint,integer,uuid,text
) to authenticated;

select 'Sprint 11.5.0 Build 4 incremental quality is ready' as status;
