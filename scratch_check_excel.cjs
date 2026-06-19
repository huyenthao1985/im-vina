const XLSX = require('./node_modules/xlsx');

try {
  const wb = XLSX.readFile('C:/Users/dell/Desktop/Test 2.xlsx');
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  
  const keys = Object.keys(rows[0] || {});
  const findKey = (names) => {
    for (const n of names) {
      const found = keys.find(k => k.toLowerCase() === n.toLowerCase());
      if (found) return found;
    }
    return null;
  };
  
  const keyDivision = findKey(['division']);
  const keyMonth = findKey(['month']);
  const keyValue = findKey(["q'ty/amt", "q`ty/amt", "qty/amt", "value"]);
  
  console.log('Keys found:', { keyDivision, keyMonth, keyValue });
  
  const agg = {};
  rows.forEach(r => {
    const div = String(r[keyDivision] || '').toUpperCase();
    const month = String(r[keyMonth] || '').toUpperCase().trim();
    const val = Number(r[keyValue]) || 0;
    
    let divisionType = null;
    if (div.includes('PROD')) divisionType = 'production';
    else if (div.includes('SHIP')) divisionType = 'shipment';
    else if (div.includes('SALES')) divisionType = 'sales';
    
    if (divisionType) {
      if (!agg[month]) agg[month] = { production: 0, shipment: 0, sales: 0 };
      agg[month][divisionType] += val;
    }
  });
  
  console.log('Aggregated Monthly Data:');
  console.log(JSON.stringify(agg, null, 2));
} catch (e) {
  console.error(e);
}
