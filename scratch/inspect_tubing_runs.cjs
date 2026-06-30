const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Mock LETTER_MONTH and extractRawMonth to check
const LETTER_MONTH = { A:1, B:2, C:3, D:4, E:5, F:6, G:7, H:8, I:9, J:10, K:11, L:12 };

function findDateInSheet(ws, actualQC) {
  // Mock findDateInSheet from etl_pipeline.cjs
  return null; // Let's see what happens if it returns null
}

function extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC) {
  var dateInfo = findDateInSheet(ws, actualQC);
  if (dateInfo) {
    if (dateInfo.year === year) return dateInfo.month;
  }
  var y = String(year);
  var n, mn;
  n = fileName.match(/(\d{4})[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    var yr = parseInt(n[1], 10);
    if (yr === year) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }
  n = fileName.match(/(\d{4})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/i);
  if (n) {
    var yr = parseInt(n[1], 10);
    if (yr === year) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }
  n = fileName.match(/(\d{2})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/);
  if (n) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }
  n = fileName.match(/(?:^|[^\d])(\d{2})(\d{2})\.xlsx$/i);
  if (n && n[1] === String(year).slice(-2)) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }
  n = fileName.match(/[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    mn = parseInt(n[1], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }
  n = sheetName.match(/(\d{2})(\d{2})\d{2}/);
  if (n) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }
  n = sheetName.match(/(\d{1,2})月/);
  if (n) {
    mn = parseInt(n[1], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }
  if (relPath) {
    n = relPath.match(/[-_](\d{1,2})$/);
    if (n) {
      mn = parseInt(n[1], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
    n = relPath.match(/(\d{1,2})月/);
    if (n) {
      mn = parseInt(n[1], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }
  return null;
}

const targetDir = 'RawData/2025/零組件入庫-2025/Tubing-2025';

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

walkDir(targetDir, (filePath, fileName) => {
  const relPath = path.relative('RawData/2025', filePath);
  const wb = XLSX.readFile(filePath);
  wb.SheetNames.forEach(sheetName => {
    if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例') return;
    if (sheetName.trim().indexOf('出貨') === 0) return;
    
    // Check cell O4 first
    const ws = wb.Sheets[sheetName];
    const cell = ws['O4'];
    let o4Val = cell ? (cell.w || cell.v || '') : '';
    
    const month = extractRawMonth(ws, fileName, sheetName, 2025, relPath, [], 'QC10007-R03');
    console.log(`${fileName} -> ${sheetName} : O4="${o4Val}" | month = ${month}`);
  });
});
