const XLSX = require('./node_modules/xlsx');

try {
  const wb = XLSX.readFile('public/Test 2.xlsx');
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(sheet);
  
  console.log('Total raw rows:', rawRows.length);
  if (rawRows.length === 0) {
    console.log('Empty sheet.');
    process.exit(0);
  }

  // Trim keys
  const rows = rawRows.map(row => {
    const trimmed = {};
    for (const key of Object.keys(row)) {
      trimmed[key.trim()] = row[key];
    }
    return trimmed;
  });

  const keys = Object.keys(rows[0]);
  console.log('Trimmed Keys:', keys);

  // Filter TTL Sales rows
  const ttlSales = rows.filter(r => {
    const m = String(r['MONTH'] || '').trim().toUpperCase();
    const div = String(r['Division'] || '').trim().toUpperCase();
    return m === 'TTL' && div.includes('SALES');
  });

  console.log('TTL Sales Rows Count:', ttlSales.length);
  if (ttlSales.length > 0) {
    console.log('Sample TTL Sales Row:', ttlSales[0]);
    
    // Check if Value is present and non-zero
    const nonZero = ttlSales.filter(r => {
      const val = Number(r['Value']) || 0;
      return val !== 0;
    });
    console.log('Non-zero TTL Sales Rows:', nonZero.length);
    if (nonZero.length > 0) {
      console.log('Sample Non-zero TTL Sales Row:', nonZero[0]);
    } else {
      console.log('ALL TTL Sales Rows have value 0 or NaN! First 5 TTL Sales Rows:', ttlSales.slice(0, 5));
    }
  }

} catch (err) {
  console.error('Error:', err);
}
