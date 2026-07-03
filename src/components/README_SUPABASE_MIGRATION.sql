-- ============================================================================
-- MIGRATION: Fix "mất dữ liệu Sheet 2 (Per Capita) trong Mục 3 sau khi F5"
-- Chạy 1 LẦN DUY NHẤT trong Supabase → SQL Editor → New query → Run.
-- An toàn: KHÔNG xoá dữ liệu cũ, chỉ thêm cột mới + copy dữ liệu vào đó.
-- ============================================================================

-- 1) Thêm cột jsonb để lưu toàn bộ dữ liệu 1 dòng, không phụ thuộc schema cứng
alter table public.sales_data
  add column if not exists payload jsonb;

-- 2) Backfill: gom toàn bộ cột hiện có của mỗi dòng CŨ (trước migration) vào
--    payload, để code mới (đọc thống nhất qua payload) vẫn thấy đủ dữ liệu cũ.
--    to_jsonb(sales_data) tự lấy MỌI cột hiện tại của bảng, không cần bạn liệt
--    kê tên cột thủ công.
update public.sales_data
set payload = to_jsonb(sales_data) - 'id' - 'payload' - 'source_tag'
where payload is null;

-- 3) (Khuyến nghị) Index trên source_tag để lọc theo bucket nhanh hơn
create index if not exists idx_sales_data_source_tag on public.sales_data (source_tag);

-- Xong. Không cần xoá các cột cũ (model/type/date/value/...) — chúng vẫn còn
-- đó nhưng từ nay code chỉ đọc/ghi qua cột `payload`, nên có thêm bao nhiêu
-- cột mới ở bất kỳ sheet Excel nào trong tương lai cũng sẽ không bao giờ làm
-- vỡ insert nữa.
