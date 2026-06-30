const XLSX = require('xlsx');

function inspectExcel(filePath) {
  console.log('--- Inspecting:', filePath);
  try {
    const wb = XLSX.readFile(filePath);
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet);
    if (rows.length === 0) {
      console.log('Empty sheet.');
      return;
    }
    
    // Trim keys
    const trimmedRows = rows.map(r => {
      const tr = {};
      Object.keys(r).forEach(k => {
        tr[k.trim()] = r[k];
      });
      return tr;
    });

    const uniqueTypes = new Set();
    const uniqueModels = new Set();
    trimmedRows.forEach(r => {
      if (r.Type) uniqueTypes.add(r.Type);
      if (r.type) uniqueTypes.add(r.type);
      if (r.Model) uniqueModels.add(r.Model);
      if (r.model) uniqueModels.add(r.model);
    });

    console.log('Unique Types:', Array.from(uniqueTypes));
    console.log('Unique Models:', Array.from(uniqueModels));
    console.log('Sample Row:', trimmedRows[0]);
  } catch (err) {
    console.error('Error:', err);
  }
}

inspectExcel('public/Production_Per_Capita_Sample.xlsx');
inspectExcel('public/Manpower_Sample.xlsx');
inspectExcel('public/Test 1.xlsx');
