/**
 * cleanup_supabase.cjs
 * Xóa TOÀN BỘ dữ liệu trong bảng sales_data của Supabase để dọn dẹp
 * 28,808+ dòng bị phình do các lần "Save to Cloud" trước đây.
 * Chạy: node cleanup_supabase.cjs
 */
const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = 'https://pquxjrfyafsaybuzovqy.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo';

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

async function main() {
  console.log('🔍 Đang kiểm tra số lượng dữ liệu hiện tại...');
  
  const { count, error: countError } = await supabase
    .from('sales_data')
    .select('*', { count: 'exact', head: true });
  
  if (countError) {
    console.error('Lỗi kết nối Supabase:', countError.message);
    return;
  }
  
  console.log(`📊 Số dòng hiện tại trong sales_data: ${count}`);
  
  if (!count || count === 0) {
    console.log('✅ Database đã sạch, không cần xóa.');
    return;
  }
  
  console.log('🗑️  Đang xóa toàn bộ dữ liệu...');
  
  // Xóa theo từng bucket để tránh timeout (bảng free tier giới hạn)
  for (const tag of ['Manpower', 'TargetActual', 'Sales']) {
    console.log(`   - Xóa bucket ${tag}...`);
    if (tag === 'Sales') {
      const { error } = await supabase
        .from('sales_data')
        .delete()
        .or('source_tag.eq.Sales,source_tag.is.null');
      if (error) console.error(`   Lỗi khi xóa ${tag}:`, error.message);
      else console.log(`   ✅ Đã xóa ${tag}`);
    } else {
      const { error } = await supabase
        .from('sales_data')
        .delete()
        .eq('source_tag', tag);
      if (error) console.error(`   Lỗi khi xóa ${tag}:`, error.message);
      else console.log(`   ✅ Đã xóa ${tag}`);
    }
  }
  
  // Kiểm tra lại
  const { count: finalCount } = await supabase
    .from('sales_data')
    .select('*', { count: 'exact', head: true });
  
  console.log(`\n✅ Hoàn thành! Số dòng còn lại: ${finalCount ?? 0}`);
  console.log('\n👉 Tiếp theo: Vào dashboard, upload lại các file Excel và bấm "Lưu dữ liệu lên đám mây".');
}

main().catch(console.error);
