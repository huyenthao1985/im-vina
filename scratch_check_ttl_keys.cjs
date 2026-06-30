const XLSX = require('./node_modules/xlsx');

function analyzeFile(filePath) {
  console.log('--- Analyzing File:', filePath);
  try {
    const wb = XLSX.readFile(filePath);
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet);
    if (rawRows.length === 0) {
      console.log('No rows.');
      return;
    }
    
    // Trim keys to match App.tsx behavior
    const rows = rawRows.map(row => {
      const trimmedRow = {};
      Object.keys(row).forEach(key => {
        trimmedRow[key.trim()] = row[key];
      });
      return trimmedRow;
    });

    console.log('Total rows:', rows.length);
    console.log('Trimmed Keys:', Object.keys(rows[0]));
    
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
    
    const ttlRows = rows.filter(r => {
      const month = keyMonth ? String(r[keyMonth] || '').trim().toUpperCase() : '';
      return month === 'TTL';
    });
    
    console.log('TTL rows count:', ttlRows.length);
    if (ttlRows.length > 0) {
      const salesTtlRows = ttlRows.filter(r => {
        const divUpper = keyDivision ? String(r[keyDivision] || '').toUpperCase() : '';
        return divUpper.includes('SALES');
      });
      console.log('Sales TTL rows count:', salesTtlRows.length);
      if (salesTtlRows.length > 0) {
        console.log('Sample Sales TTL row:', salesTtlRows[0]);
        const modelPresentInTtl = salesTtlRows.filter(r => keyModel && String(r[keyModel] || '').trim() !== '').length;
        console.log('Number of Sales TTL rows with model present:', modelPresentInTtl);
        
        // Sum values
        const totalVal = salesTtlRows.reduce((sum, r) => sum + (Number(r[keyValue]) || 0), 0);
        console.log('Total sales value in TTL rows:', totalVal);
      }
    }
  } catch (err) {
    console.error('Error analyzing file:', err);
  }
}

analyzeFile('public/Test 2.xlsx');
