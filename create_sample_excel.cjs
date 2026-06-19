const XLSX = require('xlsx');
const wb = XLSX.utils.book_new();
const data = [
  ['Ngày', 'Danh mục', 'Doanh thu'],
  [new Date('2026-06-01'), 'Thời trang', 1500000],
  [new Date('2026-06-02'), 'Điện tử', 3200000],
  [new Date('2026-06-03'), 'Thời trang', 1200000],
  [new Date('2026-06-04'), 'Gia dụng', 850000],
  [new Date('2026-06-05'), 'Điện tử', 4500000],
];
const ws = XLSX.utils.aoa_to_sheet(data);
XLSX.utils.book_append_sheet(wb, ws, 'Sales');
XLSX.writeFile(wb, 'C:\\Users\\dell\\.gemini\\antigravity-ide\\scratch\\sample_marketing_data.xlsx');
console.log('Sample Excel file generated at C:\\Users\\dell\\.gemini\\antigravity-ide\\scratch\\sample_marketing_data.xlsx');
