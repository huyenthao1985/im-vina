const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function checkFile(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    console.log(`File: ${path.basename(filePath)}`);
    console.log(`  Sheets: ${workbook.SheetNames.join(', ')}`);
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (rows.length > 0) {
        console.log(`    Sheet "${sheetName}": ${rows.length} rows. First row:`, rows[0].slice(0, 8));
      }
    });
  } catch (err) {
    console.error(`  Error reading ${filePath}:`, err.message);
  }
}

console.log("=== CHECKING PUBLIC EXCEL FILES ===");
const publicDir = 'c:\\Users\\dell\\.gemini\\antigravity-ide\\scratch\\marketing-dashboard\\public';
fs.readdirSync(publicDir).forEach(file => {
  if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
    checkFile(path.join(publicDir, file));
  }
});

console.log("\n=== CHECKING DOWNLOADS EXCEL FILES ===");
const downloadsDir = 'c:\\Users\\dell\\Downloads';
fs.readdirSync(downloadsDir).forEach(file => {
  if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
    checkFile(path.join(downloadsDir, file));
  }
});
