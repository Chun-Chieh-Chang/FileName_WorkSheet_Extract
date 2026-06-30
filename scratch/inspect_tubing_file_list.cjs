const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function walkDir(dirPath, callback) {
  let entries;
  try { entries = fs.readdirSync(dirPath); } catch(e) { return; }
  entries.forEach(entry => {
    const full = path.join(dirPath, entry);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      walkDir(full, callback);
    } else if (entry.toLowerCase().endsWith('.xlsx') && !entry.startsWith('~$')) {
      callback(full, entry);
    }
  });
}

const targetDir = 'RawData/2025/零組件入庫-2025/Tubing-2025';
walkDir(targetDir, (filePath, fileName) => {
  const wb = XLSX.readFile(filePath);
  console.log(`${fileName} -> Sheets:`, wb.SheetNames);
});
