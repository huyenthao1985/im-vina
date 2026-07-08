/**
 * cleanup_all_batch.cjs
 * Xóa TOÀN BỘ dữ liệu trong sales_data bằng cách lấy id rồi xóa theo batch
 * Không phụ thuộc vào source_tag filter
 * Chạy: node cleanup_all_batch.cjs
 */

const SUPABASE_URL = 'https://pquxjrfyafsaybuzovqy.supabase.co';
const ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo';

const hdrs = {
  'apikey': ANON_KEY,
  'Authorization': `Bearer ${ANON_KEY}`,
  'Content-Type': 'application/json',
  'Prefer': 'return=minimal',
};

async function fetchIds(limit) {
  // Lấy ID cũ nhất (order by id ascending, không filter source_tag)
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/sales_data?select=id&order=id.asc&limit=${limit}`,
    { headers: { ...hdrs, 'Prefer': '' } }
  );
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`fetchIds failed: ${res.status} ${t}`);
  }
  const rows = await res.json();
  return rows.map(r => r.id);
}

async function deleteIds(ids) {
  if (ids.length === 0) return 0;
  // Supabase REST: id=in.(uuid1,uuid2,...)
  const joined = ids.join(',');
  const url = `${SUPABASE_URL}/rest/v1/sales_data?id=in.(${joined})`;
  const res = await fetch(url, { method: 'DELETE', headers: hdrs });
  if (res.status !== 204 && res.status !== 200) {
    const t = await res.text();
    throw new Error(`deleteIds failed: ${res.status} ${t}`);
  }
  return ids.length;
}

async function main() {
  const BATCH = 200; // nhỏ hơn để tránh timeout
  let totalDeleted = 0;
  let round = 0;
  const MAX_ROUNDS = 10000;

  console.log(`🗑️  Bắt đầu xóa toàn bộ sales_data (batch=${BATCH})...`);

  while (round < MAX_ROUNDS) {
    round++;
    let ids;
    try {
      ids = await fetchIds(BATCH);
    } catch (e) {
      console.error(`\nLỗi lấy IDs (round ${round}):`, e.message);
      await new Promise(r => setTimeout(r, 3000));
      continue;
    }

    if (ids.length === 0) {
      console.log(`\n🎉 Xóa xong! Tổng đã xóa: ${totalDeleted.toLocaleString()} dòng`);
      break;
    }

    try {
      const n = await deleteIds(ids);
      totalDeleted += n;
      process.stdout.write(`\r  Round ${round}: xóa ${n} dòng | Tổng: ${totalDeleted.toLocaleString()}   `);
    } catch (e) {
      console.error(`\nLỗi xóa (round ${round}):`, e.message);
      await new Promise(r => setTimeout(r, 3000));
    }

    // Nghỉ ngắn giữa các batch
    await new Promise(r => setTimeout(r, 200));
  }

  // Verify
  try {
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/sales_data?select=id&limit=1`,
      { headers: { ...hdrs, 'Prefer': 'count=exact' } }
    );
    console.log('\n📊 Kiểm tra cuối: Content-Range =', res.headers.get('content-range'));
  } catch(e) { /* ignore */ }
}

main().catch(err => {
  console.error('\nFatal error:', err);
  process.exit(1);
});
