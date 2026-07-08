/**
 * cleanup_supabase_v2.cjs
 * Xóa dữ liệu Supabase dùng node-fetch hoặc built-in fetch (Node 18+)
 * Chạy: node cleanup_supabase_v2.cjs
 */

const SUPABASE_URL = 'https://pquxjrfyafsaybuzovqy.supabase.co';
const ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo';

const headers = {
  'apikey': ANON_KEY,
  'Authorization': `Bearer ${ANON_KEY}`,
  'Content-Type': 'application/json',
  'Prefer': 'return=minimal',
};

async function countRows() {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/sales_data?select=id&limit=1`,
    { headers: { ...headers, 'Prefer': 'count=exact' } }
  );
  const count = res.headers.get('content-range');
  console.log('Content-Range header:', count);
  return count;
}

async function deleteBucket(tag) {
  const url = `${SUPABASE_URL}/rest/v1/sales_data?source_tag=eq.${encodeURIComponent(tag)}`;
  const res = await fetch(url, { method: 'DELETE', headers });
  const text = await res.text();
  console.log(`DELETE ${tag}: status=${res.status}, body=${text || '(empty)'}`);
  return res.status;
}

async function deleteNullTag() {
  const url = `${SUPABASE_URL}/rest/v1/sales_data?source_tag=is.null`;
  const res = await fetch(url, { method: 'DELETE', headers });
  const text = await res.text();
  console.log(`DELETE null tags: status=${res.status}, body=${text || '(empty)'}`);
}

async function main() {
  console.log('🔍 Kiểm tra kết nối Supabase...');
  
  // Test GET first
  const testRes = await fetch(
    `${SUPABASE_URL}/rest/v1/sales_data?select=source_tag&limit=5`,
    { headers }
  );
  
  if (!testRes.ok) {
    console.error('❌ Lỗi kết nối:', testRes.status, await testRes.text());
    return;
  }
  
  const sample = await testRes.json();
  console.log('✅ Kết nối OK. Mẫu 5 dòng đầu:', JSON.stringify(sample));
  
  await countRows();
  
  console.log('\n🗑️ Bắt đầu xóa dữ liệu...');
  for (const tag of ['Manpower', 'TargetActual', 'Sales']) {
    await deleteBucket(tag);
  }
  await deleteNullTag();
  
  console.log('\n✅ Hoàn thành! Kiểm tra lại:');
  await countRows();
}

main().catch(err => console.error('Fatal error:', err));
