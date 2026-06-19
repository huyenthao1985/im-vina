const XLSX = require('./node_modules/xlsx');

try {
  const wb = XLSX.readFile('C:/Users/dell/Desktop/Test 2.xlsx');
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  if (rows.length > 0) {
    console.log('Row keys:', Object.keys(rows[0]));
    console.log('Sample row:', rows[0]);
  } else {
    console.log('No rows found');
  }
} catch (e) {
  console.error(e);
}
