/**
 * cleanup_manpower_batch.cjs
 * Xóa dữ liệu Manpower theo batch 500 dòng, tránh statement timeout 57014
 * Chạy: node cleanup_manpower_batch.cjs
 */

const SUPABASE_URL = 'https://pquxjrfyafsaybuzovqy.supabase.co';
const ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo';

const headers = {
  'apikey': ANON_KEY,
  'Authorization': `Bearer ${ANON_KEY}`,
  'Content-Type': 'application/json',
};

async function getCount() {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/sales_data?source_tag=eq.Manpower&select=id&limit=1`,
    { headers: { ...headers, 'Prefer': 'count=exact' } }
  );
  const cr = res.headers.get('content-range');
  const total = cr ? parseInt(cr.split('/')[1]) : 0;
  return total;
}

async function getOldestIds(limit) {
  // Lấy các id cũ nhất (ascending) để xóa từ cũ đến mới
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/sales_data?source_tag=eq.Manpower&select=id&order=id.asc&limit=${limit}`,
    { headers }
  );
  if (!res.ok) {
    console.error('Lỗi lấy IDs:', res.status, await res.text());
    return [];
  }
  const rows = await res.json();
  return rows.map(r => r.id);
}

async function deleteByIds(ids) {
  if (ids.length === 0) return 204;
  // Dùng `in` filter để xóa nhiều id cùng lúc
  const idList = `(${ids.join(',')})`;
  const url = `${SUPABASE_URL}/rest/v1/sales_data?id=in.${encodeURIComponent(idList)}`;
  const res = await fetch(url, {
    method: 'DELETE',
    headers: { ...headers, 'Prefer': 'return=minimal' },
  });
  return res.status;
}

async function main() {
  console.log('🔍 Đang đếm dòng Manpower...');
  let total = await getCount();
  console.log(`📊 Hiện tại: ${total.toLocaleString()} dòng Manpower`);

  if (total === 0) {
    console.log('✅ Không có dữ liệu Manpower, xong!');
    return;
  }

  const BATCH = 500;
  let deleted = 0;
  let round = 0;
  const MAX_ROUNDS = 2100; // tối đa 2100 * 500 = 1,050,000 dòng

  while (deleted < total && round < MAX_ROUNDS) {
    round++;
    const ids = await getOldestIds(BATCH);
    if (ids.length === 0) {
      console.log('Không còn ids để xóa.');
      break;
    }

    const status = await deleteByIds(ids);
    if (status === 204 || status === 200) {
      deleted += ids.length;
      const remaining = total - deleted;
      process.stdout.write(`\r  Đã xóa: ${deleted.toLocaleString()}/${total.toLocaleString()} (còn ~${Math.max(0,remaining).toLocaleString()})   `);
    } else {
      console.error(`\nLỗi batch ${round}: status ${status}`);
      // Thử lại với batch nhỏ hơn
      await new Promise(r => setTimeout(r, 2000));
      break;
    }

    // Nghỉ 100ms giữa các batch để tránh rate limit
    await new Promise(r => setTimeout(r, 100));
  }

  console.log('\n\n🔍 Kiểm tra lại...');
  const remaining = await getCount();
  console.log(`✅ Còn lại: ${remaining.toLocaleString()} dòng Manpower`);
  
  if (remaining > 0) {
    console.log(`⚠️  Còn ${remaining} dòng. Chạy lại script để tiếp tục xóa.`);
  } else {
    console.log('🎉 Database đã sạch hoàn toàn!');
  }
}

main().catch(err => console.error('Fatal error:', err));
