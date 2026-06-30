const XLSX = require('xlsx');

const filePath = 'DataExtract/2025品檢報表統計.xlsx';
const wb = XLSX.readFile(filePath);

const sheetName = '零組件入庫品檢(QC10007-R03)';
const ws = wb.Sheets[sheetName];
if (!ws) {
  console.log(`Sheet "${sheetName}" not found!`);
  process.exit(1);
}

const json = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
console.log('Sheet data:');
json.forEach(row => {
  console.log(row.slice(0, 10).map(v => String(v).trim()));
});
