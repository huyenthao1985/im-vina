const https = require('https');

const SUPABASE_URL = 'https://pquxjrfyafsaybuzovqy.supabase.co';
const ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBxdXhqcmZ5YWZzYXlidXpvdnF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MjYwNzUsImV4cCI6MjA5NTQwMjA3NX0.x3j55_kArTHzDeA1kbelzp73yGQC_H0TcZEwP6pqnAo';

// Test connection - check how many rows have source_tag IS NULL
function makeRequest(path, method, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(SUPABASE_URL + path);
    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: method,
      headers: {
        'Content-Type': 'application/json',
        'apikey': ANON_KEY,
        'Authorization': 'Bearer ' + ANON_KEY,
        'Prefer': 'return=representation',
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, body: JSON.parse(data) });
        } catch {
          resolve({ status: res.statusCode, body: data });
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(JSON.stringify(body));
    req.end();
  });
}

async function main() {
  console.log('=== Kiểm tra kết nối Supabase ===');
  
  // Count rows with source_tag = null
  const checkNull = await makeRequest(
    '/rest/v1/sales_data?source_tag=is.null&select=id&limit=5',
    'GET'
  );
  console.log('Rows with source_tag IS NULL (sample):', checkNull.status, Array.isArray(checkNull.body) ? checkNull.body.length + ' rows returned' : checkNull.body);

  // Count total rows
  const checkTotal = await makeRequest(
    '/rest/v1/sales_data?select=id&limit=1',
    'GET'
  );
  console.log('Connection status:', checkTotal.status === 200 ? '✅ Connected' : '❌ Failed');

  console.log('\n=== Kiểm tra dữ liệu theo source_tag ===');
  const checkSales = await makeRequest('/rest/v1/sales_data?source_tag=eq.Sales&select=id&limit=1', 'GET');
  console.log('Rows with source_tag=Sales (exists):', checkSales.status === 200 ? '✅' : '❌', checkSales.status);
  
  const checkManpower = await makeRequest('/rest/v1/sales_data?source_tag=eq.Manpower&select=id&limit=1', 'GET');
  console.log('Rows with source_tag=Manpower:', checkManpower.status === 200 ? Array.isArray(checkManpower.body) && checkManpower.body.length > 0 ? '✅ HAS DATA' : '⚠️ EMPTY' : '❌ ERROR');
  
  const checkTA = await makeRequest('/rest/v1/sales_data?source_tag=eq.TargetActual&select=id&limit=1', 'GET');
  console.log('Rows with source_tag=TargetActual:', checkTA.status === 200 ? Array.isArray(checkTA.body) && checkTA.body.length > 0 ? '✅ HAS DATA' : '⚠️ EMPTY' : '❌ ERROR');

  console.log('\nNote: The UPDATE and CREATE INDEX SQL commands require service role key.');
  console.log('Please run these 2 statements manually in Supabase SQL Editor:');
  console.log('');
  console.log('  update sales_data');
  console.log('    set source_tag = origin');
  console.log("    where origin in ('Manpower', 'TargetActual')");
  console.log('    and source_tag is null;');
  console.log('');
  console.log('  create index if not exists idx_sales_data_source_tag on sales_data (source_tag);');
}

main().catch(console.error);
