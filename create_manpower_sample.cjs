const XLSX = require('./node_modules/xlsx');

const wb = XLSX.utils.book_new();

const headers = ['Model', 'Type', 'Date', 'Value'];

const data = [
  headers,
  // Standard YR24 row
  ['TTL', 'TTL ManPower AVG', 'YR24', 180],
  
  // Model SO1B01
  ['SO1B01', 'TTL ManPower AVG', 'JAN', 50],
  ['SO1B01', 'TTL ManPower AVG', 'FEB', 52],
  ['SO1B01', 'TTL ManPower AVG', 'MAR', 55],
  ['SO1B01', 'TTL ManPower AVG', 'APR', 58],
  ['SO1B01', 'TTL ManPower AVG', 'MAY', 60],
  ['SO1B01', 'TTL ManPower AVG', 'JUN', 62],
  ['SO1B01', 'TTL ManPower AVG', 'W22', 59],
  ['SO1B01', 'TTL ManPower AVG', 'W23', 60],
  ['SO1B01', 'TTL ManPower AVG', 'W24', 61],
  ['SO1B01', 'TTL ManPower AVG', 'W25', 62],
  ['SO1B01', 'TTL ManPower AVG', 'W26', 63],
  ['SO1B01', 'TTL ManPower AVG', 'W27', 64],
  ['SO1B01', 'TTL ManPower AVG', 'W28', 65],
  ['SO1B01', 'TTL ManPower AVG', '06/22', 61],
  ['SO1B01', 'TTL ManPower AVG', '06/23', 62],
  ['SO1B01', 'TTL ManPower AVG', '06/24', 63],
  ['SO1B01', 'TTL ManPower AVG', '06/25', 62],
  ['SO1B01', 'TTL ManPower AVG', '06/26', 63],
  ['SO1B01', 'TTL ManPower AVG', '06/27', 64],
  ['SO1B01', 'TTL ManPower AVG', '06/28', 65],

  // Model SO1C2EF
  ['SO1C2EF', 'TTL ManPower AVG', 'JAN', 30],
  ['SO1C2EF', 'TTL ManPower AVG', 'FEB', 31],
  ['SO1C2EF', 'TTL ManPower AVG', 'MAR', 32],
  ['SO1C2EF', 'TTL ManPower AVG', 'APR', 33],
  ['SO1C2EF', 'TTL ManPower AVG', 'MAY', 35],
  ['SO1C2EF', 'TTL ManPower AVG', 'JUN', 36],
  ['SO1C2EF', 'TTL ManPower AVG', 'W22', 34],
  ['SO1C2EF', 'TTL ManPower AVG', 'W23', 35],
  ['SO1C2EF', 'TTL ManPower AVG', 'W24', 35],
  ['SO1C2EF', 'TTL ManPower AVG', 'W25', 36],
  ['SO1C2EF', 'TTL ManPower AVG', 'W26', 36],
  ['SO1C2EF', 'TTL ManPower AVG', 'W27', 37],
  ['SO1C2EF', 'TTL ManPower AVG', 'W28', 38],
  ['SO1C2EF', 'TTL ManPower AVG', '06/22', 35],
  ['SO1C2EF', 'TTL ManPower AVG', '06/23', 36],
  ['SO1C2EF', 'TTL ManPower AVG', '06/24', 36],
  ['SO1C2EF', 'TTL ManPower AVG', '06/25', 36],
  ['SO1C2EF', 'TTL ManPower AVG', '06/26', 37],
  ['SO1C2EF', 'TTL ManPower AVG', '06/27', 38],
  ['SO1C2EF', 'TTL ManPower AVG', '06/28', 38],

  // Model SO1C2G
  ['SO1C2G', 'TTL ManPower AVG', 'JAN', 40],
  ['SO1C2G', 'TTL ManPower AVG', 'FEB', 41],
  ['SO1C2G', 'TTL ManPower AVG', 'MAR', 42],
  ['SO1C2G', 'TTL ManPower AVG', 'APR', 40],
  ['SO1C2G', 'TTL ManPower AVG', 'MAY', 43],
  ['SO1C2G', 'TTL ManPower AVG', 'JUN', 45],
  ['SO1C2G', 'TTL ManPower AVG', 'W22', 42],
  ['SO1C2G', 'TTL ManPower AVG', 'W23', 43],
  ['SO1C2G', 'TTL ManPower AVG', 'W24', 44],
  ['SO1C2G', 'TTL ManPower AVG', 'W25', 45],
  ['SO1C2G', 'TTL ManPower AVG', 'W26', 44],
  ['SO1C2G', 'TTL ManPower AVG', 'W27', 46],
  ['SO1C2G', 'TTL ManPower AVG', 'W28', 47],
  ['SO1C2G', 'TTL ManPower AVG', '06/22', 43],
  ['SO1C2G', 'TTL ManPower AVG', '06/23', 44],
  ['SO1C2G', 'TTL ManPower AVG', '06/24', 44],
  ['SO1C2G', 'TTL ManPower AVG', '06/25', 45],
  ['SO1C2G', 'TTL ManPower AVG', '06/26', 46],
  ['SO1C2G', 'TTL ManPower AVG', '06/27', 46],
  ['SO1C2G', 'TTL ManPower AVG', '06/28', 47],

  // Model SO3560
  ['SO3560', 'TTL ManPower AVG', 'JAN', 20],
  ['SO3560', 'TTL ManPower AVG', 'FEB', 22],
  ['SO3560', 'TTL ManPower AVG', 'MAR', 21],
  ['SO3560', 'TTL ManPower AVG', 'APR', 23],
  ['SO3560', 'TTL ManPower AVG', 'MAY', 24],
  ['SO3560', 'TTL ManPower AVG', 'JUN', 25],
  ['SO3560', 'TTL ManPower AVG', 'W22', 23],
  ['SO3560', 'TTL ManPower AVG', 'W23', 24],
  ['SO3560', 'TTL ManPower AVG', 'W24', 24],
  ['SO3560', 'TTL ManPower AVG', 'W25', 25],
  ['SO3560', 'TTL ManPower AVG', 'W26', 26],
  ['SO3560', 'TTL ManPower AVG', 'W27', 25],
  ['SO3560', 'TTL ManPower AVG', 'W28', 27],
  ['SO3560', 'TTL ManPower AVG', '06/22', 25],
  ['SO3560', 'TTL ManPower AVG', '06/23', 25],
  ['SO3560', 'TTL ManPower AVG', '06/24', 26],
  ['SO3560', 'TTL ManPower AVG', '06/25', 26],
  ['SO3560', 'TTL ManPower AVG', '06/26', 26],
  ['SO3560', 'TTL ManPower AVG', '06/27', 27],
  ['SO3560', 'TTL ManPower AVG', '06/28', 27],

  // Total (TTL) rows
  ['TTL', 'TTL ManPower AVG', 'JAN', 140],
  ['TTL', 'TTL ManPower AVG', 'FEB', 146],
  ['TTL', 'TTL ManPower AVG', 'MAR', 150],
  ['TTL', 'TTL ManPower AVG', 'APR', 154],
  ['TTL', 'TTL ManPower AVG', 'MAY', 162],
  ['TTL', 'TTL ManPower AVG', 'JUN', 168],
  ['TTL', 'TTL ManPower AVG', 'W22', 148],
  ['TTL', 'TTL ManPower AVG', 'W23', 152],
  ['TTL', 'TTL ManPower AVG', 'W24', 154],
  ['TTL', 'TTL ManPower AVG', 'W25', 158],
  ['TTL', 'TTL ManPower AVG', 'W26', 159],
  ['TTL', 'TTL ManPower AVG', 'W27', 164],
  ['TTL', 'TTL ManPower AVG', 'W28', 167],
  ['TTL', 'TTL ManPower AVG', '06/22', 154],
  ['TTL', 'TTL ManPower AVG', '06/23', 157],
  ['TTL', 'TTL ManPower AVG', '06/24', 159],
  ['TTL', 'TTL ManPower AVG', '06/25', 159],
  ['TTL', 'TTL ManPower AVG', '06/26', 162],
  ['TTL', 'TTL ManPower AVG', '06/27', 165],
  ['TTL', 'TTL ManPower AVG', '06/28', 167]
];

const ws = XLSX.utils.aoa_to_sheet(data);
XLSX.utils.book_append_sheet(wb, ws, 'Manpower');

// Write to public and downloads
XLSX.writeFile(wb, 'c:\\Users\\dell\\.gemini\\antigravity-ide\\scratch\\marketing-dashboard\\public\\Manpower_Sample.xlsx');
XLSX.writeFile(wb, 'c:\\Users\\dell\\Downloads\\Manpower_Sample.xlsx');

console.log("Generated Manpower_Sample.xlsx successfully in public and downloads!");
