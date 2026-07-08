-- ============================================================
-- IM VINA — Auth schema (bổ sung vào project Supabase HIỆN CÓ)
-- Chạy toàn bộ file này 1 lần trong: Supabase Dashboard > SQL Editor
-- An toàn để chạy trên project đang lưu dữ liệu sales/manpower:
--   - Chỉ tạo bảng MỚI `profiles` (không đụng tới bảng dữ liệu hiện có).
--   - Dùng "if not exists" / "drop policy if exists" nên chạy lại nhiều
--     lần không lỗi, không mất dữ liệu.
-- ============================================================

create table if not exists public.profiles (
  id         uuid primary key references auth.users(id) on delete cascade,
  full_name  text not null,
  email      text not null,
  role       text,               -- null = đang chờ duyệt; 'user' | 'editor' | 'admin'
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

drop policy if exists "Users can view own profile" on public.profiles;
create policy "Users can view own profile"
  on public.profiles for select
  using (auth.uid() = id);

drop policy if exists "Admins can view all profiles" on public.profiles;
create policy "Admins can view all profiles"
  on public.profiles for select
  using (
    exists (
      select 1 from public.profiles p
      where p.id = auth.uid() and p.role = 'admin'
    )
  );

drop policy if exists "Users can insert own profile" on public.profiles;
create policy "Users can insert own profile"
  on public.profiles for insert
  with check (auth.uid() = id);

-- Chỉ admin mới gán được role — kiểm tra ngay trong Postgres (security
-- definer), chặn việc client tự UPDATE cột role qua REST API.
create or replace function public.admin_assign_role(target_id uuid, new_role text)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if not exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'admin'
  ) then
    raise exception 'Only admins can assign roles';
  end if;

  if new_role not in ('user', 'editor', 'admin') then
    raise exception 'Invalid role: %', new_role;
  end if;

  update public.profiles set role = new_role where id = target_id;
end;
$$;

grant execute on function public.admin_assign_role(uuid, text) to authenticated;

-- ============================================================
-- BƯỚC BẮT BUỘC SAU KHI CHẠY FILE NÀY (bootstrap admin đầu tiên):
--   1. Vào app, Đăng ký 1 tài khoản bất kỳ.
--   2. Quay lại đây, chạy (đổi email cho đúng):
--
--      update public.profiles set role = 'admin'
--      where email = 'ban@vidu.com';
--
--   3. Đăng nhập lại tài khoản đó -> role='admin' -> thấy nút "Quản trị"
--      (góc phải header) để duyệt/phân quyền cho các tài khoản sau.
-- ============================================================
