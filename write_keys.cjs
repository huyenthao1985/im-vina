const XLSX = require('./node_modules/xlsx');
const fs = require('fs');

try {
  const wb = XLSX.readFile('C:/Users/dell/Desktop/Test 2.xlsx');
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  if (rows.length > 0) {
    const output = {
      keys: Object.keys(rows[0]),
      sample: rows[0]
    };
    fs.writeFileSync('C:/Users/dell/.gemini/antigravity-ide/scratch/marketing-dashboard/keys.txt', JSON.stringify(output, null, 2));
    console.log('Successfully wrote keys to keys.txt');
  } else {
    fs.writeFileSync('C:/Users/dell/.gemini/antigravity-ide/scratch/marketing-dashboard/keys.txt', 'No rows found');
  }
} catch (e) {
  fs.writeFileSync('C:/Users/dell/.gemini/antigravity-ide/scratch/marketing-dashboard/keys.txt', e.message);
  console.error(e);
}
