-- Sprint 11.2.3 upload permission verification
-- Run in Supabase SQL Editor while logged in as project owner.

-- Confirm the administrator profile exists.
select id, email, role, is_active
from public.profiles
where lower(email) = lower('lbenelisha@gmail.com');

-- Repair the role if required.
update public.profiles
set role = 'admin', is_active = true, updated_at = now()
where lower(email) = lower('lbenelisha@gmail.com');

-- Ensure authenticated users can execute the permission and activation functions.
grant execute on function public.iml_is_admin() to authenticated;
grant execute on function public.iml_activate_dataset_version(uuid) to authenticated;

-- Show recent failed uploads and their exact reason.
select created_at, kind, file_name, row_count, status, error_message, duration_ms
from public.iml_upload_history
order by created_at desc
limit 20;
