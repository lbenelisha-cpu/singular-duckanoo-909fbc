-- IML Control Sprint 10.1.1
-- Dedicated schema recovery migration.
-- Run the ENTIRE file in Supabase SQL Editor, then confirm the verification row at the end.

create extension if not exists pgcrypto;

-- 1. Active version pointer on the existing metadata table.
alter table public.iml_data_sources
  add column if not exists active_version_id uuid;

-- 2. Dataset version header.
create table if not exists public.iml_dataset_versions (
  id uuid primary key default gen_random_uuid(),
  kind text not null check (kind in ('production','quality','deviations','targets','packaging_plan')),
  version_no bigint not null,
  file_name text not null default '',
  row_count bigint not null default 0,
  raw_row_count bigint not null default 0,
  facilities integer not null default 0,
  chunk_count integer not null default 0,
  uploaded_by uuid references auth.users(id) on delete set null,
  uploaded_by_email text not null default '',
  status text not null default 'uploading' check (status in ('uploading','active','archived','failed')),
  error_message text not null default '',
  created_at timestamptz not null default now(),
  activated_at timestamptz,
  unique(kind, version_no)
);

-- 3. Automatic sequential version number per dataset kind.
create or replace function public.iml_assign_version_no()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if new.version_no is null or new.version_no = 0 then
    perform pg_advisory_xact_lock(hashtext('iml-version-' || new.kind));
    select coalesce(max(version_no), 0) + 1 into new.version_no
    from public.iml_dataset_versions where kind = new.kind;
  end if;
  return new;
end;
$$;

drop trigger if exists iml_dataset_version_number on public.iml_dataset_versions;
create trigger iml_dataset_version_number
before insert on public.iml_dataset_versions
for each row execute function public.iml_assign_version_no();

-- 4. Versioned data chunks.
create table if not exists public.iml_dataset_chunks (
  version_id uuid not null references public.iml_dataset_versions(id) on delete cascade,
  chunk_index integer not null,
  payload jsonb not null default '[]'::jsonb,
  row_count integer not null default 0,
  created_at timestamptz not null default now(),
  primary key(version_id, chunk_index)
);

alter table public.iml_upload_history add column if not exists version_id uuid;
alter table public.iml_upload_history add column if not exists duration_ms bigint;

create index if not exists iml_dataset_versions_kind_status_idx
  on public.iml_dataset_versions(kind, status, created_at desc);
create index if not exists iml_dataset_chunks_version_idx
  on public.iml_dataset_chunks(version_id, chunk_index);

-- 5. Security policies.
alter table public.iml_dataset_versions enable row level security;
alter table public.iml_dataset_chunks enable row level security;

drop policy if exists "authenticated read dataset versions" on public.iml_dataset_versions;
create policy "authenticated read dataset versions" on public.iml_dataset_versions
for select to authenticated using (true);

drop policy if exists "admin manage dataset versions" on public.iml_dataset_versions;
create policy "admin manage dataset versions" on public.iml_dataset_versions
for all to authenticated using (public.iml_is_admin()) with check (public.iml_is_admin());

drop policy if exists "authenticated read dataset chunks" on public.iml_dataset_chunks;
create policy "authenticated read dataset chunks" on public.iml_dataset_chunks
for select to authenticated using (true);

drop policy if exists "admin manage dataset chunks" on public.iml_dataset_chunks;
create policy "admin manage dataset chunks" on public.iml_dataset_chunks
for all to authenticated using (public.iml_is_admin()) with check (public.iml_is_admin());

-- 6. Atomic activation: a new version becomes visible only after verification.
create or replace function public.iml_activate_dataset_version(p_version_id uuid)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v public.iml_dataset_versions%rowtype;
  actual_chunks integer;
  actual_rows bigint;
begin
  if not public.iml_is_admin() then
    raise exception 'Admin permission required';
  end if;

  select * into v
  from public.iml_dataset_versions
  where id = p_version_id
  for update;

  if not found then raise exception 'Dataset version not found'; end if;
  if v.status <> 'uploading' then raise exception 'Dataset version is not awaiting activation'; end if;

  select count(*), coalesce(sum(row_count),0)
  into actual_chunks, actual_rows
  from public.iml_dataset_chunks
  where version_id = p_version_id;

  if actual_chunks <> v.chunk_count or actual_rows <> v.row_count then
    raise exception 'Dataset verification failed: chunks %, rows %', actual_chunks, actual_rows;
  end if;

  update public.iml_dataset_versions
  set status = 'archived'
  where kind = v.kind and status = 'active' and id <> v.id;

  update public.iml_dataset_versions
  set status = 'active', activated_at = now(), error_message = ''
  where id = v.id;

  insert into public.iml_data_sources(
    kind,file_name,row_count,raw_row_count,facilities,loaded_at,
    loaded_by,loaded_by_email,updated_at,active_version_id
  ) values (
    v.kind,v.file_name,v.row_count,v.raw_row_count,v.facilities,now(),
    v.uploaded_by,v.uploaded_by_email,now(),v.id
  )
  on conflict(kind) do update set
    file_name=excluded.file_name,
    row_count=excluded.row_count,
    raw_row_count=excluded.raw_row_count,
    facilities=excluded.facilities,
    loaded_at=excluded.loaded_at,
    loaded_by=excluded.loaded_by,
    loaded_by_email=excluded.loaded_by_email,
    updated_at=excluded.updated_at,
    active_version_id=excluded.active_version_id;
end;
$$;

grant execute on function public.iml_activate_dataset_version(uuid) to authenticated;

-- 7. Safe full-data deletion function.
create or replace function public.iml_delete_all_datasets()
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if not public.iml_is_admin() then raise exception 'Admin permission required'; end if;
  delete from public.iml_data_sources
  where kind in ('production','quality','deviations','targets','packaging_plan');
  delete from public.iml_dataset_versions
  where kind in ('production','quality','deviations','targets','packaging_plan');
  delete from public.iml_data_chunks
  where kind in ('production','quality','deviations','targets','packaging_plan');
end;
$$;

grant execute on function public.iml_delete_all_datasets() to authenticated;

-- 8. Realtime and PostgREST schema refresh.
do $$
begin
  alter publication supabase_realtime add table public.iml_dataset_versions;
exception when duplicate_object then null;
end $$;

notify pgrst, 'reload schema';

-- 9. Verification. Expected result: all three columns are TRUE.
select
  to_regclass('public.iml_dataset_versions') is not null as versions_table_ready,
  to_regclass('public.iml_dataset_chunks') is not null as chunks_table_ready,
  exists (
    select 1 from information_schema.columns
    where table_schema='public'
      and table_name='iml_data_sources'
      and column_name='active_version_id'
  ) as active_version_column_ready;
