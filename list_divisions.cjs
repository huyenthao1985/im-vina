const XLSX = require('./node_modules/xlsx');

try {
  const wb = XLSX.readFile('C:/Users/dell/Desktop/Test 2.xlsx');
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const divs = new Set();
  rows.forEach(r => {
    Object.keys(r).forEach(k => {
      if (k.trim() === 'Division') {
        divs.add(String(r[k]).trim());
      }
    });
  });
  
  console.log('Unique Divisions:', Array.from(divs));
} catch (e) {
  console.error(e);
}
