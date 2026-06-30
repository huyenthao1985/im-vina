const XLSX = require('./node_modules/xlsx');

function analyzeFile(filePath) {
  console.log('--- Analyzing File:', filePath);
  try {
    const wb = XLSX.readFile(filePath);
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet);
    if (rows.length === 0) {
      console.log('No rows.');
      return;
    }
    console.log('Total rows in file:', rows.length);
    
    // Find keys matching year, month, model, division
    const keys = Object.keys(rows[0]);
    const findKey = (names) => {
      for (const n of names) {
        const found = keys.find(k => k.toLowerCase() === n.toLowerCase());
        if (found) return found;
      }
      return null;
    };
    
    const keyModel = findKey(['model']);
    const keyOrigin = findKey(['출하지', 'origin']);
    const keyDivision = findKey(['division']);
    const keyYear = findKey(['year']);
    const keyMonth = findKey(['month']);
    const keyValue = findKey(["q`ty/amt", "q'ty/amt", 'qty/amt', 'qtyamt', 'value', 'val']);
    
    console.log('Detected mapping:', { keyModel, keyOrigin, keyDivision, keyYear, keyMonth, keyValue });
    
    // Let's filter some rows where Month is TTL
    const ttlRows = rows.filter(r => {
      const month = keyMonth ? String(r[keyMonth] || '').trim().toUpperCase() : '';
      return month === 'TTL';
    });
    
    console.log('TTL rows count:', ttlRows.length);
    if (ttlRows.length > 0) {
      console.log('Sample TTL row:', ttlRows[0]);
      
      const salesTtlRows = ttlRows.filter(r => {
        const divUpper = keyDivision ? String(r[keyDivision] || '').toUpperCase() : '';
        return divUpper.includes('SALES');
      });
      console.log('Sales TTL rows count:', salesTtlRows.length);
      if (salesTtlRows.length > 0) {
        console.log('Sample Sales TTL row:', salesTtlRows[0]);
        const modelPresentInTtl = salesTtlRows.filter(r => keyModel && r[keyModel]).length;
        console.log('Number of Sales TTL rows with model present:', modelPresentInTtl);
      }
    }
  } catch (err) {
    console.error('Error analyzing file:', err);
  }
}

analyzeFile('public/Test 1.xlsx');
analyzeFile('public/Test 2.xlsx');
