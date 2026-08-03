-- IML Control Sprint 11.2.2
-- Public read-only viewer without email/password.
-- Run once in Supabase SQL Editor.
-- IMPORTANT: this intentionally allows the public anon role to read dashboard datasets.

begin;

-- Allow the browser client (anon key) to read only the dashboard data tables.
grant usage on schema public to anon;
grant select on table public.iml_data_sources to anon;
grant select on table public.iml_data_chunks to anon;
grant select on table public.iml_dataset_versions to anon;
grant select on table public.iml_dataset_chunks to anon;
grant select on table public.iml_upload_history to anon;

-- Read-only RLS policies for public viewer mode.
drop policy if exists "public viewer read iml sources" on public.iml_data_sources;
create policy "public viewer read iml sources"
on public.iml_data_sources for select to anon using (true);

drop policy if exists "public viewer read iml chunks" on public.iml_data_chunks;
create policy "public viewer read iml chunks"
on public.iml_data_chunks for select to anon using (true);

drop policy if exists "public viewer read dataset versions" on public.iml_dataset_versions;
create policy "public viewer read dataset versions"
on public.iml_dataset_versions for select to anon using (true);

drop policy if exists "public viewer read dataset chunks" on public.iml_dataset_chunks;
create policy "public viewer read dataset chunks"
on public.iml_dataset_chunks for select to anon using (true);

drop policy if exists "public viewer read upload history" on public.iml_upload_history;
create policy "public viewer read upload history"
on public.iml_upload_history for select to anon using (true);

commit;
