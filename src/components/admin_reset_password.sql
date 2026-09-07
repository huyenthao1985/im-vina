-- admin_reset_password.sql
--
-- Cho phép admin đặt lại mật khẩu cho user khác một cách an toàn, không cần
-- lộ service_role key ra client (theo đúng pattern của `admin_assign_role`
-- đã có trong supabase/auth_schema.sql — hàm security definer tự kiểm tra
-- quyền admin ngay trong Postgres).
--
-- LƯU Ý: mật khẩu CŨ của user không bao giờ có thể xem lại được — Supabase
-- Auth chỉ lưu bcrypt hash (một chiều) trong auth.users.encrypted_password.
-- Hàm này chỉ GHI ĐÈ bằng mật khẩu MỚI do admin nhập, không đọc mật khẩu cũ.
--
-- QUAN TRỌNG: mệnh đề kiểm tra quyền admin bên dưới giả định bảng
-- `public.profiles` có cột `role`, và admin được đánh dấu role = 'admin'
-- (giống điều kiện đang dùng để lọc `approved` trong AdminUsersPanel.tsx).
-- Hãy đối chiếu với mệnh đề kiểm tra quyền THẬT trong `admin_assign_role`
-- (trong auth_schema.sql) và sửa lại cho khớp nếu khác.

create extension if not exists pgcrypto;

create or replace function public.admin_reset_password(target_id uuid, new_password text)
returns void
language plpgsql
security definer
set search_path = public, auth, extensions
as $$
begin
  -- Chỉ admin mới được gọi hàm này.
  if not exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'admin'
  ) then
    raise exception 'Not authorized';
  end if;

  if new_password is null or length(new_password) < 6 then
    raise exception 'Password must be at least 6 characters';
  end if;

  update auth.users
  set encrypted_password = extensions.crypt(new_password, extensions.gen_salt('bf')),
      updated_at = now()
  where id = target_id;

  if not found then
    raise exception 'User not found';
  end if;
end;
$$;

-- Chỉ user đã đăng nhập (authenticated) mới được gọi — hàm tự kiểm tra thêm
-- quyền admin ở bên trong, nên vẫn an toàn dù client nào cũng gọi được.
revoke all on function public.admin_reset_password(uuid, text) from public;
grant execute on function public.admin_reset_password(uuid, text) to authenticated;
