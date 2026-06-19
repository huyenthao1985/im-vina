-- Dán toàn bộ mã này vào mục "SQL Editor" trên Supabase và bấm "Run"

CREATE TABLE IF NOT EXISTS public.sales_data (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  model text NOT NULL,
  origin text,
  customer text,
  type text,
  division text NOT NULL, -- 'sales', 'shipment', 'production'
  year integer NOT NULL,
  month text NOT NULL,
  value numeric NOT NULL,
  created_at timestamp with time zone DEFAULT now()
);

-- Tắt Row Level Security (RLS) để cho phép đọc/ghi thoải mái từ Dashboard 
-- (Trong dự án thực tế, bạn nên bật RLS và phân quyền cẩn thận)
ALTER TABLE public.sales_data DISABLE ROW LEVEL SECURITY;
