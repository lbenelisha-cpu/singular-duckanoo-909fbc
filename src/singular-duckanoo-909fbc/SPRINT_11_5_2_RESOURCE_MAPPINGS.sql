-- Sprint 11.5.2 — shared resource mapping table
create table if not exists public.iml_resource_mappings (
  id text primary key,
  facility text not null,
  description text not null,
  match_mode text not null default 'exact' check (match_mode in ('exact','contains')),
  resource text not null,
  active boolean not null default true,
  updated_at timestamptz not null default now()
);

alter table public.iml_resource_mappings enable row level security;

drop policy if exists "resource mappings read" on public.iml_resource_mappings;
create policy "resource mappings read" on public.iml_resource_mappings for select to anon, authenticated using (true);

drop policy if exists "resource mappings admin insert" on public.iml_resource_mappings;
create policy "resource mappings admin insert" on public.iml_resource_mappings for insert to authenticated with check (public.iml_is_admin());
drop policy if exists "resource mappings admin update" on public.iml_resource_mappings;
create policy "resource mappings admin update" on public.iml_resource_mappings for update to authenticated using (public.iml_is_admin()) with check (public.iml_is_admin());
drop policy if exists "resource mappings admin delete" on public.iml_resource_mappings;
create policy "resource mappings admin delete" on public.iml_resource_mappings for delete to authenticated using (public.iml_is_admin());

insert into public.iml_resource_mappings (id,facility,description,match_mode,resource,active) values
('1542-liquid-1','1542','LIQUID 1 LITER','exact','LQ 1lt (42)',true),
('1542-liquid-5','1542','LIQUID 5 LITER','exact','LQ 5 lt (42)',true),
('1542-liquid-10','1542','LIQUID 10 LITER','exact','LQ 10/20 lt (42)',true),
('1542-liquid-20','1542','LIQUID 20 LITER','exact','LQ 10/20 lt (42)',true),
('1540-diuron','1540','DIURON','contains','Diuron (40)',true),
('1540-tolurex','1540','TOLUREX','contains','Tolurex (40)',true)
on conflict (id) do update set facility=excluded.facility,description=excluded.description,match_mode=excluded.match_mode,resource=excluded.resource,active=excluded.active,updated_at=now();

select 'Sprint 11.5.2 resource mappings ready' as result;
