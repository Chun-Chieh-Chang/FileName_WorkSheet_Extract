const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const filePath = 'RawData/2025/零組件入庫-2025/Tubing-2025/Tubing-2025-01/Tubing-2501.xlsx';
const wb = XLSX.readFile(filePath);

console.log('Worksheet names in Tubing-2501.xlsx:', wb.SheetNames);

wb.SheetNames.forEach(sheetName => {
  if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例') return;
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
  console.log(`\n=== Sheet: ${sheetName} ===`);
  
  // Print some header cells to see date positions
  for (let r = 0; r < Math.min(10, json.length); r++) {
    const row = json[r];
    if (row && row.some(v => v !== '')) {
      console.log(`  Row ${r}:`, row.slice(0, 18).map((v, c) => `${XLSX.utils.encode_col(c)}${r+1}: "${v}"`).filter(v => !v.includes('""')));
    }
  }
});
