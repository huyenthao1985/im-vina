-- ============================================================================
-- XỬ LÝ TRIỆT ĐỂ: dọn dữ liệu trùng lặp hiện có (28,808+ dòng do các lần
-- "Save to Cloud" trước đây insert thô + do môi trường dev/production trỏ
-- chung 1 Supabase project) VÀ chặn trùng vĩnh viễn về sau ở tầng database.
--
-- Chạy 1 LẦN DUY NHẤT trong Supabase → SQL Editor → New query → Run.
-- Chạy theo ĐÚNG THỨ TỰ 3 bước bên dưới (không đảo thứ tự).
-- ============================================================================

-- ── BƯỚC 0: Sao lưu trước khi xoá (an toàn, có thể rollback nếu cần) ────────
create table if not exists sales_data_backup_before_dedupe as
  table sales_data;

-- ── BƯỚC 1: Xem trước sẽ xoá bao nhiêu dòng trùng (chỉ SELECT, không xoá gì) ─
-- Khoá nghiệp vụ dùng để xác định "trùng": source_tag, model, origin,
-- customer, type, division, year, month (KHÔNG tính `value` — value là số
-- liệu, được phép ghi đè). Chạy riêng câu này trước để xem con số trước khi
-- xoá thật ở Bước 2.
select
  count(*) as tong_so_dong,
  count(*) filter (where rn > 1) as so_dong_se_bi_xoa
from (
  select
    id,
    row_number() over (
      partition by source_tag, model, origin, customer, type, division, year, month
      order by id desc  -- giữ lại dòng có id LỚN NHẤT (mới nhất) cho mỗi khoá
    ) as rn
  from sales_data
) t;

-- ── BƯỚC 2: Xoá dòng trùng, chỉ giữ lại dòng mới nhất (id lớn nhất) mỗi khoá ─
delete from sales_data
where id in (
  select id from (
    select
      id,
      row_number() over (
        partition by source_tag, model, origin, customer, type, division, year, month
        order by id desc
      ) as rn
    from sales_data
  ) t
  where rn > 1
);

-- ── BƯỚC 3: Thêm UNIQUE CONSTRAINT để Postgres tự chặn trùng vĩnh viễn ──────
-- Từ nay bất kỳ insert/upsert nào (kể cả nếu có script cũ nào đó lỡ chạy lại,
-- hay môi trường dev/local lỡ ghi vào đây) đều KHÔNG THỂ tạo dòng trùng khoá
-- nữa — Postgres sẽ báo lỗi (insert thô) hoặc tự ghi đè (upsert).
alter table sales_data
  add constraint uq_sales_data_business_key
  unique (source_tag, model, origin, customer, type, division, year, month);

-- ── XONG. Kiểm tra lại tổng số dòng sau khi dọn ──────────────────────────────
select
  source_tag,
  count(*) as so_dong
from sales_data
group by source_tag
order by source_tag;

-- ── (Chỉ chạy nếu mọi thứ ổn sau vài ngày, để giải phóng dung lượng) ────────
-- drop table if exists sales_data_backup_before_dedupe;
