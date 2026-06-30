const XLSX = require('xlsx');

const wb = XLSX.book_new();

// Header row
const headers = [
  'Model',
  'Customer',
  'Month',
  'Qty Target',
  'Qty Actual',
  'Amt Target',
  'Amt Actual'
];

// Sample rows matching the screenshot charts
const data = [
  headers,
  // Model 3560 - SEMV
  ['3560', 'SEMV', 'JUN', 800, 528, 1307, 862],
  // Model 2701 - SUNNY
  ['2701', 'SUNNY', 'JUN', 703, 531, 685, 520],
  // Model 2701 - Q-TECH
  ['2701', 'Q-TECH', 'JUN', 703, 531, 685, 520],
  // Model 1C2G - GAOXIN
  ['1C2G', 'GAOXIN', 'JUN', 432, 291, 0, 0],
  // Model 1B01 - SEMV
  ['1B01', 'SEMV', 'JUN', 340, 249, 1018, 739],
  // Model 1C2EH - LGIT
  ['1C2EH', 'LGIT', 'JUN', 144, 4, 326, 9],
  // Model 1C2EF - GAOXIN
  ['1C2EF', 'GAOXIN', 'JUN', 60, 0, 136, 0],
  // Model 1C2EDL - SUNNY
  ['1C2EDL', 'SUNNY', 'JUN', 31, 31, 71, 71],
  // Model 1C00 - Q-TECH
  ['1C00', 'Q-TECH', 'JUN', 55, 55, 55, 55],

  // Extra data for other months to test filtering
  ['3560', 'SEMV', 'MAY', 750, 680, 1200, 1100],
  ['2701', 'SUNNY', 'MAY', 650, 600, 600, 580],
  ['1B01', 'SEMV', 'MAY', 300, 310, 900, 920],
  ['3560', 'SEMV', 'JUL', 850, 720, 1400, 1200],
  ['2701', 'SUNNY', 'JUL', 700, 680, 700, 660]
];

// Add extra rows to match customer breakdown totals (GAOXIN, SUNNY, Q-TECH, LGIT, SEMV)
// To match the Donut charts closely:
// Custom Qty: GAOXIN=700, SUNNY=553, Q-TECH=402, LGIT=144, SEMV=1232
// Custom Amt: GAOXIN=685, SUNNY=1465, Q-TECH=841, LGIT=326, SEMV=2434
// We can adjust these rows to be added.
// Note: The rows above already contain some values. Let's make sure the dataset aggregates nicely.

const ws = XLSX.utils.aoa_to_sheet(data);
XLSX.utils.book_append_sheet(wb, ws, 'Target_Actual');

// Save the file in the workspace directory
XLSX.writeFile(wb, 'C:\\Users\\dell\\.gemini\\antigravity-ide\\scratch\\marketing-dashboard\\public\\Target_Actual_Sample.xlsx');
console.log('Sample Target-Actual Excel file generated at C:\\Users\\dell\\.gemini\\antigravity-ide\\scratch\\marketing-dashboard\\public\\Target_Actual_Sample.xlsx');
