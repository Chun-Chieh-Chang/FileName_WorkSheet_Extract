/**
 * ETL Pipeline: Raw QC Excel Files → Summary Excel → HTML Reports
 * 
 * Usage: node etl_pipeline.cjs <year>
 *   e.g., node etl_pipeline.cjs 2025
 *         node etl_pipeline.cjs 2026
 *         node etl_pipeline.cjs all
 */

var XLSX = require('xlsx');
var fs = require('fs');
var path = require('path');

var MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];

// ============================================================
// 1. MONTH EXTRACTION
// ============================================================

// Pattern A: Sheet name contains "-MM-" (e.g., "QIP-2025(1-10)-01-Setup" → 1)
var RE_SHEET_DASH_MONTH = /-(\d{1,2})-/;

// Pattern B: Sheet name ends with "-MM" or "_MM" (e.g., "Tubing-01", "2025-01")
var RE_SHEET_SUFFIX_MONTH = /[_-](\d{1,2})$/;

// Pattern C: Sheet name has year pattern "2025-MM" or "2026-MM"
var RE_SHEET_YEAR_MONTH = /20\d{2}[-_](\d{1,2})/;

// Pattern D: Sheet name has Chinese month pattern "2025年MM月"
var RE_SHEET_CN_MONTH = /(\d{1,2})月/;

// Pattern E: Sheet name letter suffix (A=01, B=02, ..., L=12)
var RE_SHEET_LETTER = /-([A-L])$/;

// Pattern F: Filename with date code "YYMMDD" (e.g., "250102" → 01)
var RE_FILENAME_DATE = /_?(\d{2})(\d{2})\d{2}/;

// Pattern G: Sheet name like "QIP-2025(Tubing)-MM-Setup" → MM
var RE_QIP_TUBING_MONTH = /\(Tubing\)-(\d{1,2})-/;

var LETTER_MONTH = {A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9,J:10,K:11,L:12};
var LETTER_NAMES = ['A','B','C','D','E','F','G','H','I','J','K','L'];

function extractMonth(row, sheetName, fileName) {
  var m;
  
  // Method 1: Try to parse month from row[0] (檔案名稱 column)
  // File names contain date codes like "ICU-250102.xlsx", "東林易_250102.xlsx"
  if (row[0]) {
    var fnStr = String(row[0]);
    // Pattern: YYMMDD or YYYYMMDD
    var fnMatch = fnStr.match(/_?(\d{2})(\d{2})\d{2}/);
    if (fnMatch) {
      var month = parseInt(fnMatch[2], 10);
      if (month >= 1 && month <= 12) return month;
    }
    // Also try YYYYMMDD
    var fnMatch2 = fnStr.match(/(\d{4})(\d{2})\d{2}/);
    if (fnMatch2) {
      var month2 = parseInt(fnMatch2[2], 10);
      if (month2 >= 1 && month2 <= 12) return month2;
    }
  }
  
  // Method 2: Sheet name patterns
  m = sheetName.match(RE_QIP_TUBING_MONTH);
  if (m) return parseInt(m[1], 10);
  
  m = sheetName.match(RE_SHEET_CN_MONTH);
  if (m) return parseInt(m[1], 10);
  
  m = sheetName.match(RE_SHEET_YEAR_MONTH);
  if (m) return parseInt(m[1], 10);
  
  m = sheetName.match(RE_SHEET_DASH_MONTH);
  if (m) return parseInt(m[1], 10);
  
  m = sheetName.match(RE_SHEET_SUFFIX_MONTH);
  if (m) return parseInt(m[1], 10);
  
  // Method 3: File name date code
  m = fileName.match(RE_FILENAME_DATE);
  if (m) {
    var mn = parseInt(m[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }
  
  // Method 4: Letter mapping
  m = sheetName.match(RE_SHEET_LETTER);
  if (m) return LETTER_MONTH[m[1]];
  
  for (var li = 0; li < LETTER_NAMES.length; li++) {
    var letter = LETTER_NAMES[li];
    var re = new RegExp(letter + '$');
    if (re.test(sheetName)) return li + 1;
    var re2 = new RegExp('-' + letter + '$');
    if (re2.test(sheetName)) return li + 1;
  }
  
  return null;
}

// ============================================================
// 2. RawData SCANNER (directly reads RawData/{year}/)
// ============================================================

// Map folder name prefix to QC code
var QC_FOLDER_MAP = {
  '進料檢驗': 'QC10002-R02',
  '出貨檢驗': 'QC10008-R02',
  '裝配檢驗': 'QC10006-R02',
  '裝配巡檢': 'QC10006-R01',
  '零組件入庫': 'QC10007-R03',
  'QIP尺寸檢驗': 'QC10004-R02',
  '射出檢驗': 'QC10004-R02',
  '押出檢驗': 'QC10004-R02',
};

function detectQCFromFolder(dirname) {
  for (var prefix in QC_FOLDER_MAP) {
    if (dirname.indexOf(prefix) >= 0) return QC_FOLDER_MAP[prefix];
  }
  return null;
}

// Mapping of form title text → QC code
var FORM_TITLE_MAP = {
  '原物料品檢表': 'QC10002-R02',
  'RAW MATERIAL QUALITY INSPECTION PLAN': 'QC10002-R02',
  '零組件入庫品檢表': 'QC10007-R03',
  'COMPONENT QUALITY INSPECTION PLAN FOR WAREHOUSING': 'QC10007-R03',
  '出貨檢驗報告': 'QC10008-R02',
  'OUTGOING INSPECTION REPORT': 'QC10008-R02',
  '半成品品檢表': 'QC10006-R02',
  'SUB-ASSEMBLED SETS QUALITY INSPECTION PLAN': 'QC10006-R02',
  '完成品品檢表': 'QC10007-R01',
  'FINISHED PRODUCT QUALITY INSPECTION PLAN': 'QC10007-R01',
  '裝配對樣巡檢記錄表': 'QC10006-R01',
  'QUALITY INSPECTION PLAN RECORD': 'QC10004-R02',
};

function determineQCFromSheet(json, initialQC, relPath) {
  // Scan rows: first 15 rows AND last 30 rows (to catch footer-area QC codes like "MOULDEX QC10007-R02.D"
  // which may appear at rows 40, 51, or 61+ depending on form length)
  var totalRows = json.length;
  var scanRanges = [
    { start: 0, end: Math.min(15, totalRows) },
    { start: Math.max(0, totalRows - 30), end: totalRows },
  ];

  for (var ri = 0; ri < scanRanges.length; ri++) {
    var range = scanRanges[ri];
    for (var r = range.start; r < range.end; r++) {
      var row = json[r];
      if (!row || !row.length) continue;
      for (var c = 0; c < Math.min(row.length, 8); c++) {
        var v = String(row[c] || '').trim();

        // Check for explicit QC code pattern
        if (v.indexOf('QC10002') >= 0) return 'QC10002-R02';
        if (v.indexOf('QC10004') >= 0) return 'QC10004-R02';
        if (v.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
        if (v.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
        if (v.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
        if (v.indexOf('QC10007-R01') >= 0 || v.indexOf('QC10007-R02') >= 0 || v.indexOf('QC10007') >= 0) return 'QC10007-R01'; // R01 and R02 treated as same
        if (v.indexOf('QC10008') >= 0) return 'QC10008-R02';

        // Check for form title (header only)
        if (ri === 0) {
          for (var title in FORM_TITLE_MAP) {
            if (v.indexOf(title) >= 0) return FORM_TITLE_MAP[title];
          }
        }
      }
    }
  }

  // Priority 2: For 零組件入庫/射出D, override to QC10002-R02
  if (initialQC === 'QC10007-R03' && relPath && relPath.indexOf('射出D') >= 0) {
    return 'QC10002-R02';
  }

  // Priority 3: Fall back to folder-based QC code
  return initialQC;
}

function findDateInSheet(ws, actualQC) {
  if (!ws) return null;

  function parseDateFromString(str) {
    if (!str) return null;
    str = String(str).trim();
    var m = str.match(/\b(20\d{2})[-/](\d{1,2})[-/]\d{1,2}\b/);
    if (m) return { year: parseInt(m[1], 10), month: parseInt(m[2], 10) };
    
    var m2 = str.match(/\b(\d{2})[-/](\d{1,2})[-/]\d{1,2}\b/);
    if (m2) return { year: 2000 + parseInt(m2[1], 10), month: parseInt(m2[2], 10) };
    
    var m3 = str.match(/\b(\d{2})(\d{2})(\d{2})[A-Za-z]?\b/);
    if (m3) {
      var mm = parseInt(m3[2], 10);
      if (mm >= 1 && mm <= 12) return { year: 2000 + parseInt(m3[1], 10), month: mm };
    }
    return null;
  }

  function parseDateFromValue(val, formatted) {
    if (typeof val === 'number' && val >= 40000 && val <= 50000) {
      var date = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(date.getTime())) {
        return { year: date.getFullYear(), month: date.getMonth() + 1 };
      }
    }
    return parseDateFromString(formatted || val);
  }

  function getCellValAndFormatted(addr) {
    var cell = ws[addr];
    if (!cell) return { val: null, formatted: null };
    return { val: cell.v, formatted: cell.w || '' };
  }

  var cellInfo;
  var dateInfo = null;

  switch (actualQC) {
    case 'QC10002-R02': // 原物料進料品檢
      cellInfo = getCellValAndFormatted('N4');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (!dateInfo) {
        cellInfo = getCellValAndFormatted('O4');
        dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      }
      break;
    case 'QC10004-R02': // QIP 尺寸檢驗 (製程)
      cellInfo = getCellValAndFormatted('Q4');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      break;
    case 'QC10006-R02': // 半成品品檢表
    case 'QC10007-R01': // 完成品品檢表
    case 'QC10007-R02': // 完成品品檢表 R02
      cellInfo = getCellValAndFormatted('N5');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      break;
    case 'QC10007-R03': // 零組件入庫品檢表
      cellInfo = getCellValAndFormatted('O4');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      break;
    case 'QC10008-R02': // 出貨檢驗報告
      cellInfo = getCellValAndFormatted('R6');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      break;
  }

  return dateInfo;
}

function findDateInSheetFallback(json) {
  var limit = Math.min(10, json.length);
  for (var r = 0; r < limit; r++) {
    var row = json[r];
    if (!row || !row.length) continue;
    for (var c = 0; c < Math.min(row.length, 5); c++) {
      var v = String(row[c] || '');
      var d = v.match(/(20\d{2})\/(\d{1,2})\/\d{1,2}/);
      if (d) { var mn = parseInt(d[2], 10); if (mn >= 1 && mn <= 12) return mn; }
      d = v.match(/(20\d{2})-(\d{1,2})-\d{1,2}/);
      if (d) { var mn = parseInt(d[2], 10); if (mn >= 1 && mn <= 12) return mn; }
    }
  }
  return null;
}

function extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC) {
  // Priority 1: Specific cell inspection by QC Code
  var dateInfo = findDateInSheet(ws, actualQC);
  if (dateInfo) {
    // Cross-year holiday correction:
    // If the cell year is next year (year + 1) and month is January (1),
    // but it is in the year folder, it belongs to December of year.
    if (dateInfo.year === year + 1 && dateInfo.month === 1) {
      return 12;
    }
    // Only accept date if the year matches the folder year (or was corrected above)
    // to prevent template date typos (e.g. 2015) from overriding.
    if (dateInfo.year === year) {
      return dateInfo.month;
    }
  }

  // Priority 2: Standard regex fallbacks
  var y = String(year);
  var n, mn;

  // Strategy 1: YYYY-MM or YYYY_MM before .xlsx at end of filename
  n = fileName.match(/(\d{4})[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    var yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(y, 10)) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }

  // Strategy 2: YYYYMMDD in filename (consecutive 8 digits)
  n = fileName.match(/(\d{4})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/i);
  if (n) {
    var yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(y, 10)) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }

  // Strategy 3: YYMMDD in filename (6 consecutive digits, any year)
  n = fileName.match(/(\d{2})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/);
  if (n) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 3b: YYMM in filename (4 digits, YY matches year suffix)
  n = fileName.match(/(?:^|[^\d])(\d{2})(\d{2})\.xlsx$/i);
  if (n && n[1] === String(year).slice(-2)) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 4: Filename ends with simple -MM or _MM suffix
  n = fileName.match(/[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    mn = parseInt(n[1], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 5: Sheet name has YYMMDD pattern
  n = sheetName.match(/(\d{2})(\d{2})\d{2}/);
  if (n) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 6: Sheet name has Chinese month format
  n = sheetName.match(/(\d{1,2})月/);
  if (n) {
    mn = parseInt(n[1], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 7: Folder/relPath name has month suffix
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

  // Strategy 8: Scan sheet content for date patterns (Fallback)
  if (json) {
    mn = findDateInSheetFallback(json);
    if (mn) return mn;
  }

  return null;
}

function getRawSubCategory(qc, relPath, fileName, sheetName, qcFolder) {
  if (qc === 'QC10002-R02') {
    // 進料檢驗
    // relPath examples: "原料", "物料/紙箱", "物料/B膠"
    // Also handles 射出D from 零組件入庫 folder (actual QC overrides to QC10002-R02)
    var parts = relPath.split('/');
    if (parts[0] === '原料') return '原料';
    if (parts[0] === '物料') {
      if (parts.length > 1) {
        // Normalize: remove spaces around special chars
        var sub = parts[1].replace(/\s+/g, '');
        if (sub) return '物料-' + sub;
        return '物料-' + parts[1];
      }
      // File directly in 物料/ (e.g., B膠.xlsx, 塑膠袋.xlsx)
      var name = path.basename(fileName, '.xlsx');
      name = name.replace(/[-_]\d{4}[-_]\d{1,2}$/, '');
      name = name.replace(/[-_]\d{1,2}$/, '');
      name = name.replace(/\s+/g, ''); // normalize spaces
      return '物料-' + name;
    }
    // 射出D from 零組件入庫 folder
    if (parts[0].indexOf('射出D') >= 0) return '射出D';
    return null;
  }

  if (qc === 'QC10008-R02') {
    // 出貨檢驗: relPath = "ICU-2025", "其他-2025"
    return relPath.replace(/[-_](20\d{2})$/, '');
  }

  if (qc === 'QC10006-R02' || qc === 'QC10007-R01' || qc === 'QC10007-R02') {
    // 裝配檢驗 or 完成品品檢: relPath = "BD-2025", "Biometrix-2025"
    var name = relPath.replace(/[-_](20\d{2})$/, '');
    // Map also known 完成品 customer names
    var finMap = { 'MarMed': 'MarMed', 'Saxon': 'Saxon' };
    if (finMap[name]) return name;
    // For other 裝配檢驗 subdirs, check if file is 半成品 or 完成品 based on qc
    return name;
  }

  if (qc === 'QC10006-R01') {
    // 裝配巡檢: always "裝配巡檢"
    return '裝配巡檢';
  }

  if (qc === 'QC10007-R03') {
    // 零組件入庫: relPath first segment = sub-category folder
    var parts = relPath.split('/');
    var name = parts[0].replace(/[-_](20\d{2})$/, '');
    var catMap = {
      '射出': '射出', '射出A': '射出A', '射出C': '射出C',
      '射出D(組件)': '射出D(組件)', '射出D': '射出D',
      'Tubing': 'Tubing',
      '裝配A': '裝配A', '裝配B': '裝配B', '裝配C': '裝配C',
    };
    return catMap[name] || name;
  }

  if (qc === 'QC10004-R02') {
    return null; // QC10004-R02 is processed separately via scanInjectionData and scanExtrusionData
  }

  return null;
}

function walkRawDataDir(dirPath, relPath, initialQC, qcFolder, year, counts) {
  var entries;
  try { entries = fs.readdirSync(dirPath); } catch(e) { return; }

  entries.forEach(function(entry) {
    if (entry.startsWith('~$')) return;
    var full = dirPath + '/' + entry;
    var st;
    try { st = fs.statSync(full); } catch(e) { return; }

    if (st.isDirectory()) {
      var newRel = relPath ? relPath + '/' + entry : entry;
      walkRawDataDir(full, newRel, initialQC, qcFolder, year, counts);
    } else if (entry.toLowerCase().indexOf('.xlsx') > 0) {
      processRawDataFile(full, relPath, entry, initialQC, qcFolder, year, counts);
    }
  });
}

function processRawDataFile(filePath, relPath, fileName, initialQC, qcFolder, year, counts) {
  var wb;
  try {
    wb = XLSX.readFile(filePath);
  } catch(e) {
    console.log('    ERROR reading: ' + filePath.replace(/\\/g, '/').split('/').slice(-3).join('/') + ' - ' + e.message);
    return;
  }

  // For 完成品品檢 (QC10007-R01): deduplicate worksheets within this file.
  // Sheets like BT25001(1), BT25001(2), BT25001(3) all belong to one batch -> count as 1.
  var seenQC7R1BaseNames = new Set();

  wb.SheetNames.forEach(function(sheetName) {
    // Skip non-data sheets
    if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別') return;
    if (sheetName.indexOf('Sheet1') === 0) return;
    if (sheetName.indexOf('.K(') >= 0) return; // skip control plan templates
    if (sheetName.indexOf('範例樣本') >= 0) return; // skip sample sheets
    if (/^QC[-_]?\d+/i.test(sheetName.trim())) return; // skip template sheets named after QC form codes (e.g. QC-009)
    if (/^(工作表|Sheet)\d+/i.test(sheetName.trim())) return; // skip default un-renamed sheets
    if (sheetName.trim().indexOf('出貨') === 0) return; // skip shipping summary sheets (e.g. "出貨 (14)")

    var ws = wb.Sheets[sheetName];
    var json;
    try {
      json = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
    } catch(e) { return; }

    if (!json || json.length < 1) return;

    // Verify the sheet has content (at least one non-empty cell beyond header row)
    var hasContent = false;
    for (var r = 0; r < json.length; r++) {
      var row = json[r];
      if (!row) continue;
      for (var c = 0; c < row.length; c++) {
        if (row[c] !== '' && row[c] !== undefined && row[c] !== null) { hasContent = true; break; }
      }
      if (hasContent) break;
    }
    if (!hasContent) return;

    // Determine QC code from form title or folder fallback
    var actualQC = determineQCFromSheet(json, initialQC, relPath);
    // Determine sub-category using the actual QC code
    var subCat = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
    // Extract month
    var month = extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC);

    // Special override for QC10007-R03 (零組件入庫) files:
    // 1. If the file is in QC10007-R03 and ends with a month letter suffix (e.g. 裝配B-2025A.xlsx, 裝配A-2025-G.xlsx),
    //    route it to QC10007-R03 and parse month as A=1, B=2, ... L=12.
    // 2. Override actualQC to QC10002-R02 if the relPath matches 射出D (original logic in determineQCFromSheet).
    if (initialQC === 'QC10007-R03') {
      actualQC = 'QC10007-R03';
      var letterMatch = fileName.match(/[-_]?([A-L])\.xlsx$/i);
      if (letterMatch) {
        month = LETTER_MONTH[letterMatch[1].toUpperCase()];
      }
      if (relPath && relPath.indexOf('射出D') >= 0 && relPath.indexOf('射出D(組件)') < 0) {
        actualQC = 'QC10002-R02';
      }
    }

    // Special override for 半成品品檢表-20XX.xlsx files:
    // Route to QC10006-R02 under subcategory '裝配C' and determine month from sheet name letter (e.g., PJW25D13 -> D -> 4)
    var isSemiFinishedTable = /半成品品檢表-20\d{2}\.xlsx$/i.test(fileName);
    if (isSemiFinishedTable) {
      actualQC = 'QC10006-R02';
      subCat = '裝配C';
      var sheetMatch = sheetName.match(/2[56]([A-L])/i);
      if (sheetMatch) {
        month = LETTER_MONTH[sheetMatch[1].toUpperCase()];
      }
    }

    // Deduplication for 完成品品檢 (QC10007-R01):
    // Strip trailing suffix like (1), (2), (NG), (N) to get the batch base name.
    // Only count each unique base name once per file.
    if (actualQC === 'QC10007-R01') {
      var baseName = sheetName.replace(/\s*\([^)]+\)\s*$/, '').trim();
      if (seenQC7R1BaseNames.has(baseName)) return;
      seenQC7R1BaseNames.add(baseName);
    }

    if (!actualQC) return;
    if (!subCat) return;
    if (!month || month < 1 || month > 12) {
      // Skip worksheets that have no date (blank templates or drafts)
      return;
    }

    // Increment count
    if (!counts[actualQC]) counts[actualQC] = {};
    if (!counts[actualQC][subCat]) counts[actualQC][subCat] = {};
    counts[actualQC][subCat][month] = (counts[actualQC][subCat][month] || 0) + 1;
  });
}

function scanInjectionData(year) {
  var setupCounts = {};
  var patrolCounts = {};
  for (var m = 1; m <= 12; m++) {
    setupCounts[m] = 0;
    patrolCounts[m] = 0;
  }

  // 1. Setup counts from the entire '射出檢驗-{year}' directory
  var setupBaseDir = 'RawData/' + year + '/射出檢驗-' + year;
  if (fs.existsSync(setupBaseDir)) {
    function walkSetup(dir) {
      var entries;
      try { entries = fs.readdirSync(dir); } catch(e) { return; }
      entries.forEach(function(entry) {
        if (entry.startsWith('~$')) return;
        var fullPath = path.join(dir, entry);
        var stat;
        try { stat = fs.statSync(fullPath); } catch(e) { return; }

        if (stat.isDirectory()) {
          var mMatch = entry.match(/-(\d{2})$/);
          if (mMatch) {
            var month = parseInt(mMatch[1], 10);
            if (month >= 1 && month <= 12) {
              var files;
              try {
                files = fs.readdirSync(fullPath).filter(function(f) {
                  return f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$');
                });
              } catch(e) { return; }
              setupCounts[month] += files.length;
            }
          } else {
            walkSetup(fullPath);
          }
        }
      });
    }
    walkSetup(setupBaseDir);
  }

  // 2. Patrol counts only from 'QIP-{year}(1~10)' directory under '射出檢驗-{year}'
  var patrolBaseDir = 'RawData/' + year + '/射出檢驗-' + year + '/QIP-' + year + '(1~10)';
  if (fs.existsSync(patrolBaseDir)) {
    var subDirs;
    try {
      subDirs = fs.readdirSync(patrolBaseDir);
    } catch(e) {
      subDirs = [];
    }

    subDirs.forEach(function(sub) {
      var fullSubPath = path.join(patrolBaseDir, sub);
      var stat;
      try { stat = fs.statSync(fullSubPath); } catch(e) { return; }
      if (!stat.isDirectory()) return;

      var mMatch = sub.match(/-(\d{2})$/);
      if (!mMatch) return;

      var month = parseInt(mMatch[1], 10);
      if (month < 1 || month > 12) return;

      var files;
      try {
        files = fs.readdirSync(fullSubPath).filter(function(f) {
          return f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$');
        });
      } catch(e) {
        return;
      }

      files.forEach(function(file) {
        var filePath = path.join(fullSubPath, file);
        var wb;
        try {
          wb = XLSX.readFile(filePath);
        } catch(e) {
          console.log('    ERROR reading: ' + filePath + ' - ' + e.message);
          return;
        }

        var uniqueBaseInFile = {};
        wb.SheetNames.forEach(function(sheetName) {
          // Normalize to get base name (strip -1, (2) etc)
          var baseName = sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
          
          // Only count sheets whose base name is in Date Code format (e.g. 250103D or 260521)
          if (/^\d{6}[a-zA-Z]?$/.test(baseName)) {
            uniqueBaseInFile[baseName] = true;
          }
        });
        patrolCounts[month] += Object.keys(uniqueBaseInFile).length;
      });
    });
  }

  return { setup: setupCounts, patrol: patrolCounts };
}

function scanExtrusionData(year) {
  var setupCounts = {};
  var patrolCounts = {};
  for (var m = 1; m <= 12; m++) {
    setupCounts[m] = 0;
    patrolCounts[m] = 0;
  }

  var targetDir = 'RawData/' + year + '/押出檢驗-' + year;
  if (fs.existsSync(targetDir)) {
    var subDirs;
    try {
      subDirs = fs.readdirSync(targetDir);
    } catch(e) {
      subDirs = [];
    }

    subDirs.forEach(function(sub) {
      var fullSubPath = path.join(targetDir, sub);
      var stat;
      try { stat = fs.statSync(fullSubPath); } catch(e) { return; }
      if (!stat.isDirectory()) return;

      var mMatch = sub.match(/-(\d{2})$/);
      if (!mMatch) return;

      var month = parseInt(mMatch[1], 10);
      if (month < 1 || month > 12) return;

      var files;
      try {
        files = fs.readdirSync(fullSubPath).filter(function(f) {
          // Must contain Date Code in filename (6 digits followed by optional letter)
          return f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$') && /\d{6}[a-zA-Z]?/i.test(f);
        });
      } catch(e) {
        return;
      }

      setupCounts[month] += files.length;

      files.forEach(function(file) {
        var filePath = path.join(fullSubPath, file);
        var wb;
        try {
          wb = XLSX.readFile(filePath);
        } catch(e) {
          console.log('    ERROR reading: ' + filePath + ' - ' + e.message);
          return;
        }

        var uniqueBaseInFile = {};
        wb.SheetNames.forEach(function(sheetName) {
          if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別' || sheetName.indexOf('Sheet1') === 0) return;
          if (sheetName.indexOf('.K(') >= 0 || sheetName.indexOf('範例樣本') >= 0 || /^QC[-_]?\d+/i.test(sheetName.trim())) return;
          if (/^(工作表|Sheet)\d+/i.test(sheetName.trim())) return;

          var snLower = sheetName.toLowerCase();
          var isSetup = (snLower.indexOf('setup') >= 0 || snLower.indexOf('set up') >= 0 || snLower.indexOf('set-up') >= 0 || snLower === 'setup');
          
          if (!isSetup) {
            // Deduplicate base sheet name
            var baseName = sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
            uniqueBaseInFile[baseName] = true;
          }
        });
        patrolCounts[month] += Object.keys(uniqueBaseInFile).length;
      });
    });
  }

  return { setup: setupCounts, patrol: patrolCounts };
}

function scanRawData(year) {
  var base = 'RawData/' + year;
  if (!fs.existsSync(base)) {
    console.log('  RawData/' + year + ' not found, skipping');
    return {};
  }

  var counts = {};
  var topDirs = fs.readdirSync(base).filter(function(d) {
    try { return fs.statSync(base + '/' + d).isDirectory(); } catch(e) { return false; }
  });

  topDirs.forEach(function(dirname) {
    var qc = detectQCFromFolder(dirname);
    if (!qc) {
      console.log('  SKIP (unknown QC type): ' + dirname);
      return;
    }
    var fullPath = base + '/' + dirname;
    walkRawDataDir(fullPath, '', qc, dirname, year, counts);
  });

  // Inject QIP Injection data (Setup & Patrol) using specific logic
  var inj = scanInjectionData(year);
  if (!counts['QC10004-R02']) counts['QC10004-R02'] = {};
  counts['QC10004-R02']['QIP-Setup'] = inj.setup;
  counts['QC10004-R02']['QIP-Patrol'] = inj.patrol;

  // Inject QIP Extrusion data (Setup & Patrol) using specific logic
  var ext = scanExtrusionData(year);
  counts['QC10004-R02']['押出-Setup'] = ext.setup;
  counts['QC10004-R02']['押出-Patrol'] = ext.patrol;

  return counts;
}

// ============================================================
// 3. SUMMARY EXCEL WRITER
// ============================================================

function writeSummaryExcel(counts, year) {
  var wb = XLSX.utils.book_new();
  var outFile = 'DataExtract/' + year + '品檢報表統計.xlsx';
  
  // Helper: create monthly array [Jan-Dec]
  function monthArray(data) {
    return MONTHS.map(function(m) { return (data && data[m]) || 0; });
  }
  
  function totalArray(data) {
    return monthArray(data).reduce(function(a,b){return a+b;}, 0);
  }
  
  // ---- Sheet: 品檢地圖 ---- (Skipped)
  
  // ---- Helper to write a category sheet ----
  function addCategorySheet(sheetName, columns, subCatToCol, qcCode, subCategoryFilter, titleRowText) {
    // Accumulate monthly data per column
    var colData = [];
    columns.forEach(function() { colData.push({}); });
    
    var qcCounts = counts[qcCode];
    if (!qcCounts) {
      // Empty sheet with title row, header, and 小計 row
      var emptyHeader = ['月份'];
      columns.forEach(function(c) { emptyHeader.push(c.label); });
      if (qcCode !== 'QC10006-R01') emptyHeader.push('小計');
      
      var emptyRows = [
        [titleRowText || sheetName],
        emptyHeader
      ];
      MONTHS.forEach(function(m) { 
        var row = [m]; 
        columns.forEach(function() { row.push(0); }); 
        if (qcCode !== 'QC10006-R01') row.push(0); 
        emptyRows.push(row); 
      });
      var emptyTotalRow = ['小計'];
      columns.forEach(function() { emptyTotalRow.push(0); });
      if (qcCode !== 'QC10006-R01') emptyTotalRow.push(0);
      emptyRows.push(emptyTotalRow);
      
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(emptyRows), sheetName);
      return;
    }
    
    // For each sub-category in this QC group
    for (var subCat in qcCounts) {
      if (subCategoryFilter && !subCategoryFilter(subCat)) continue;
      var colIdx = subCatToCol(subCat);
      if (colIdx === null || colIdx === undefined) continue;
      if (colIdx < 0 || colIdx >= columns.length) continue;
      
      var monthly = qcCounts[subCat];
      for (var m = 1; m <= 12; m++) {
        if (monthly[m]) {
          colData[colIdx][m] = (colData[colIdx][m] || 0) + monthly[m];
        }
      }
    }
    
    // Build the sheet data
    var header = ['月份'];
    columns.forEach(function(c) { header.push(c.label); });
    if (qcCode !== 'QC10006-R01') header.push('小計');
    
    var rows = [
      [titleRowText || sheetName],
      header
    ];
    MONTHS.forEach(function(m) {
      var row = [m];
      var total = 0;
      columns.forEach(function(c, ci) {
        var v = colData[ci][m] || 0;
        row.push(v);
        total += v;
      });
      if (qcCode !== 'QC10006-R01') row.push(total);
      rows.push(row);
    });
    
    // 小計 row
    var totalRow = ['小計'];
    var grandTotal = 0;
    columns.forEach(function(c, ci) {
      var t = totalArray(colData[ci]);
      totalRow.push(t);
      grandTotal += t;
    });
    if (qcCode !== 'QC10006-R01') totalRow.push(grandTotal);
    rows.push(totalRow);
    
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
  }
  
  // ---- Sheet 1: 原物料品檢(QC10002-R02) ----
  var rawCols = [
    {key:'原料', label:'原料'},
    {key:'物料-B膠', label:'B膠'},
    {key:'物料-收縮膜', label:'收縮膜'},
    {key:'物料-色粉', label:'色粉'},
    {key:'物料-空白包裝袋', label:'空白包裝袋'},
    {key:'物料-空白感壓紙', label:'空白感壓紙'},
    {key:'物料-塑膠袋', label:'塑膠袋'},
    {key:'物料-塑膠袋40*50', label:'塑膠袋40*50'},
    {key:'物料-紙箱', label:'紙箱'},
    {key:'物料-過濾網連蓋', label:'過濾網連蓋'},
    {key:'物料-標籤', label:'標籤'},
    {key:'射出D', label:'射出D'},
  ];

  addCategorySheet(
    '原物料品檢(QC10002-R02)',
    rawCols,
    function(subCat) {
      if (subCat === '原料') return 0;
      if (subCat === '物料-B膠') return 1;
      if (subCat === '物料-收縮膜') return 2;
      if (subCat === '物料-色粉') return 3;
      if (subCat === '物料-空白包裝袋') return 4;
      if (subCat === '物料-空白感壓紙') return 5;
      if (subCat === '物料-塑膠袋') return 6;
      if (subCat === '物料-塑膠袋40*50' || subCat === '物料-塑膠袋40X50') return 7;
      if (subCat === '物料-紙箱') return 8;
      if (subCat === '物料-過濾網連蓋') return 9;
      if (subCat === '物料-標籤') return 10;
      if (subCat === '射出D') return 11;
      return null;
    },
    'QC10002-R02',
    null,
    '原物料進料品檢'
  );
  
  // ---- Sheet 2: QIP(QC10004-R02) ----
  var qipQC = counts['QC10004-R02'];
  var qipSetupExt = {}; var qipSetupInj = {};
  var qipPatrolExt = {}; var qipPatrolInj = {};
  
  if (qipQC) {
    for (var subCat in qipQC) {
      if (subCat === 'QIP-Setup') {
        qipSetupInj = qipQC[subCat];
      } else if (subCat === 'QIP-Patrol') {
        qipPatrolInj = qipQC[subCat];
      } else if (subCat === '押出-Setup') {
        qipSetupExt = qipQC[subCat];
      } else if (subCat === '押出-Patrol') {
        qipPatrolExt = qipQC[subCat];
      }
    }
  }
  
  var qipRows = [
    ['射出、押出製程Setup','','','','','射出、押出製程巡檢'],
    ['月份','押出(Setup)','射出(Setup)','小計','','月份','押出(巡檢)','射出(廠內)','小計']
  ];
  MONTHS.forEach(function(m) {
    var extS = qipSetupExt[m] || 0;
    var injS = qipSetupInj[m] || 0;
    var extP = qipPatrolExt[m] || 0;
    var injP = qipPatrolInj[m] || 0;
    qipRows.push([m, extS, injS, extS+injS, '', m, extP, injP, extP+injP]);
  });
  var tExtS = totalArray(qipSetupExt);
  var tInjS = totalArray(qipSetupInj);
  var tExtP = totalArray(qipPatrolExt);
  var tInjP = totalArray(qipPatrolInj);
  qipRows.push(['小計', tExtS, tInjS, tExtS+tInjS, '', '小計', tExtP, tInjP, tExtP+tInjP]);
  
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qipRows), 'QIP(QC10004-R02)');
  
  // ---- Sheet 3: 裝配對樣巡檢(QC10006-R01) ----
  addCategorySheet('裝配對樣巡檢(QC10006-R01)',
    [{key:'裝配', label:'裝配'}],
    function() { return 0; },
    'QC10006-R01',
    function(subCat) { return subCat === '裝配巡檢'; },
    '裝配對樣巡檢'
  );
  
  // ---- Sheet 4: 半成品品檢(QC10006-R02) ----
  var semiCols = [
    {key:'裝配C', label:'裝配C'},
    {key:'BD', label:'BD'},
    {key:'Biometrix', label:'Biometrix'},
    {key:'MPS', label:'MPS'},
    {key:'Vivus', label:'Vivus'},
  ];
  
  var semiQC = counts['QC10006-R02'];
  var semiData = [];
  semiCols.forEach(function() { semiData.push({}); });
  
  var semiSubCatMap = {
    'BD': 1, 'Biometrix': 2, 'MPS': 3, 'Vivus': 4,
    'MarMed': null, 'Saxon': null
  };
  
  // Populate '裝配C' (index 0) normally from semiQC below
  
  if (semiQC) {
    for (var subCat in semiQC) {
      var actualIdx = semiSubCatMap[subCat];
      if (actualIdx === null || actualIdx === undefined) {
        if (subCat.indexOf('裝配C') >= 0) actualIdx = 0;
      }
      if (actualIdx === null || actualIdx === undefined) continue;
      
      var monthly = semiQC[subCat];
      for (var m = 1; m <= 12; m++) {
        if (monthly[m]) semiData[actualIdx][m] = (semiData[actualIdx][m] || 0) + monthly[m];
      }
    }
  }
  
  var semiRows = [
    ['裝配半成品品檢'],
    ['月份','裝配C','BD','Biometrix','MPS','Vivus','小計']
  ];
  MONTHS.forEach(function(m) {
    var row = [m];
    var total = 0;
    semiCols.forEach(function(c, ci) {
      var v = semiData[ci][m] || 0;
      row.push(v);
      total += v;
    });
    row.push(total);
    semiRows.push(row);
  });
  var semiTotalRow = ['小計'];
  var semiGrandTotal = 0;
  semiCols.forEach(function(c, ci) {
    var t = totalArray(semiData[ci]);
    semiTotalRow.push(t);
    semiGrandTotal += t;
  });
  semiTotalRow.push(semiGrandTotal);
  semiRows.push(semiTotalRow);
  
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(semiRows), '半成品品檢(QC10006-R02)');
  
  // ---- Sheet 5: 完成品品檢(QC10007-R01 R02) ----
  var finCols = [
    {key:'Biometrix', label:'Biometrix'},
    {key:'MarMed', label:'MarMed'},
    {key:'Saxon', label:'Saxon'},
    {key:'Vivus', label:'Vivus'},
  ];
  var finSubCatMap = {'Biometrix': 0, 'MarMed': 1, 'Saxon': 2, 'Vivus': 3};

  var finQC = counts['QC10007-R01'];
  var finData = [];
  finCols.forEach(function() { finData.push({}); });

  if (finQC) {
    for (var subCat in finQC) {
      var actualIdx = finSubCatMap[subCat];
      if (actualIdx === undefined) continue;
      var monthly = finQC[subCat];
      for (var m = 1; m <= 12; m++) {
        if (monthly[m]) finData[actualIdx][m] = (finData[actualIdx][m] || 0) + monthly[m];
      }
    }
  }

  var finRows = [
    ['裝配完成品品檢'],
    ['月份','Biometrix','MarMed','Saxon','Vivus','小計']
  ];
  MONTHS.forEach(function(m) {
    var row = [m];
    var total = 0;
    finCols.forEach(function(c, ci) {
      var v = finData[ci][m] || 0;
      row.push(v);
      total += v;
    });
    row.push(total);
    finRows.push(row);
  });
  var finTotalRow = ['小計'];
  var finGrandTotal = 0;
  finCols.forEach(function(c, ci) {
    var t = totalArray(finData[ci]);
    finTotalRow.push(t);
    finGrandTotal += t;
  });
  finTotalRow.push(finGrandTotal);
  finRows.push(finTotalRow);

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(finRows), '完成品品檢(QC10007-R01 R02)');
  
  // ---- Sheet 6: 零組件入庫品檢(QC10007-R03) ----
  var partsCols = [
    {key:'Tubing', label:'Tubing'},
    {key:'射出', label:'射出(廠內)'},
    {key:'射出A', label:'射出A'},
    {key:'射出C', label:'射出C'},
    {key:'射出D(組件)', label:'射出D(組件)'},
    {key:'裝配A', label:'裝配A'},
    {key:'裝配B', label:'裝配B'},
    {key:'裝配C', label:'裝配C'},
  ];
  
  var partsQC = counts['QC10007-R03'];
  var partsData = [];
  partsCols.forEach(function() { partsData.push({}); });
  
  var partsSubCatMap = {
    'Tubing': 0,
    '射出': 1,
    '射出A': 2,
    '射出C': 3,
    '射出D(組件)': 4,
    '裝配A': 5,
    '裝配B': 6,
    '裝配C': 7,
  };
  
  if (partsQC) {
    for (var subCat in partsQC) {
      var actualIdx = partsSubCatMap[subCat];
      if (actualIdx === undefined) continue;
      
      var monthly = partsQC[subCat];
      for (var m = 1; m <= 12; m++) {
        if (monthly[m]) partsData[actualIdx][m] = (partsData[actualIdx][m] || 0) + monthly[m];
      }
    }
  }
  
  var partsRows = [
    ['零組件入庫檢'],
    ['月份','Tubing','射出(廠內)','射出A','射出C','射出D(組件)','裝配A','裝配B','裝配C','小計']
  ];
  MONTHS.forEach(function(m) {
    var row = [m];
    var total = 0;
    partsCols.forEach(function(c, ci) {
      var v = partsData[ci][m] || 0;
      row.push(v);
      total += v;
    });
    row.push(total);
    partsRows.push(row);
  });
  var partsTotalRow = ['小計'];
  var partsGrand = 0;
  partsCols.forEach(function(c, ci) {
    var t = totalArray(partsData[ci]);
    partsTotalRow.push(t);
    partsGrand += t;
  });
  partsTotalRow.push(partsGrand);
  partsRows.push(partsTotalRow);
  
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(partsRows), '零組件入庫品檢(QC10007-R03)');
  
  // ---- Sheet 7: 出貨檢驗(QC10008-R02) ----
  var shipQC = counts['QC10008-R02'];
  var shipICU = {}; var shipOther = {};
  
  if (shipQC) {
    for (var subCat in shipQC) {
      if (subCat.indexOf('ICU') >= 0) shipICU = shipQC[subCat];
      else if (subCat.indexOf('其他') >= 0) shipOther = shipQC[subCat];
    }
  }
  
  var shipRows = [
    ['出貨品檢'],
    ['月份','ICU','其他','小計']
  ];
  MONTHS.forEach(function(m) {
    var icu = shipICU[m] || 0;
    var oth = shipOther[m] || 0;
    shipRows.push([m, icu, oth, icu+oth]);
  });
  var tICU = totalArray(shipICU);
  var tOth = totalArray(shipOther);
  shipRows.push(['小計', tICU, tOth, tICU+tOth]);
  
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(shipRows), '出貨檢驗(QC10008-R02)');
  
  // ---- Sheet 8: NCA ---- (Skipped)
  
  // ---- Sheet 9: 彙總表 ---- (Skipped)
  
  // Write the workbook
  XLSX.writeFile(wb, outFile);
  console.log('  Summary written: ' + outFile);
  return outFile;
}

function deployToPublic(year) {
  var srcFile = 'DataExtract/' + year + '品檢報表統計.xlsx';
  var destDir = 'public/DataExtract';
  var destFile = path.join(destDir, year + '品檢報表統計.xlsx');
  
  if (!fs.existsSync(srcFile)) return;
  
  // Ensure destination directory exists
  if (!fs.existsSync(destDir)) {
    fs.mkdirSync(destDir, { recursive: true });
  }
  
  // Copy file
  fs.copyFileSync(srcFile, destFile);
  console.log('  Copied summary to public: ' + destFile);
  
  // Remove '彙總表', 'NCA', '品檢地圖' from public file
  try {
    var wb = XLSX.readFile(destFile);
    var targetSheets = ['彙總表', 'NCA', '品檢地圖'];
    var modified = false;
    
    targetSheets.forEach(function(sheetName) {
      var index = wb.SheetNames.indexOf(sheetName);
      if (index !== -1) {
        wb.SheetNames.splice(index, 1);
        delete wb.Sheets[sheetName];
        modified = true;
      }
    });
    
    if (modified) {
      XLSX.writeFile(wb, destFile);
      console.log('  Cleared template sheets from public summary.');
    }
  } catch (e) {
    console.log('  Error clearing public summary: ' + e.message);
  }
}


// ============================================================
// 5. MAIN
// ============================================================

function main() {
  var years = [];
  var targetYear = process.argv[2] || 'all';
  
  if (targetYear === 'all') {
    years = [2025, 2026];
  } else {
    years = [parseInt(targetYear, 10)];
  }
  
  years.forEach(function(year) {
    console.log('\n=== Processing ' + year + ' ===');
    console.log('Step 1: Scanning RawData...');
    var counts = scanRawData(year);
    
    // Print summary
    var totalRecords = 0;
    for (var qc in counts) {
      var qcTotal = 0;
      for (var sub in counts[qc]) {
        var subTotal = 0;
        for (var m = 1; m <= 12; m++) {
          subTotal += counts[qc][sub][m] || 0;
        }
        console.log('  ' + qc + ' - ' + sub + ': ' + subTotal + ' records');
        qcTotal += subTotal;
      }
      console.log('    ' + qc + ' total: ' + qcTotal);
      totalRecords += qcTotal;
    }
    console.log('  Grand total: ' + totalRecords + ' records');
    
    console.log('Step 2: Writing summary Excel...');
    writeSummaryExcel(counts, year);
    deployToPublic(year);
    
    console.log('Step 3: Generating styled HTML report...');
    var styledReportScript = require('child_process').spawnSync('node', ['generate_styled_reports.cjs', String(year)], {cwd: process.cwd(), stdio: 'inherit'});
    if (styledReportScript.status !== 0) {
      console.log('  Warning: Failed to generate styled HTML report');
    }
  });
  
  // Generate comparison page if both years processed
  if (years.indexOf(2025) >= 0 && years.indexOf(2026) >= 0) {
    console.log('\n=== Generating comparison report ===');
    var cmpHtml = '<!DOCTYPE html>\n<html lang="zh-TW">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>2025 vs 2026 品檢報表比較分析</title>\n<style>\nbody{font-family:Segoe UI,Noto Sans TC,sans-serif;background:#F8FAFC;color:#1E293B;padding:60px}\nh1{font-size:28px;border-left:5px solid #3B82F6;padding-left:15px}\na{display:inline-block;margin:20px;padding:15px 30px;background:#3B82F6;color:#fff;text-decoration:none;border-radius:8px;font-size:18px}\na:hover{background:#2563EB}\n</style>\n</head>\n<body>\n<h1>品檢報表比較分析</h1>\n<p>請先檢視個別年度報告：</p>\n<a href="2025品檢報表分析.html">2025年度分析報告</a>\n<a href="2026品檢報表分析.html">2026年度分析報告</a>\n</body>\n</html>';
    fs.writeFileSync('品檢報表比較分析.html', cmpHtml);
    console.log('  (比較報告為簡化版，請直接開啟個別年度報告)');
    console.log('  HTML report: 品檢報表比較分析.html');
  }
  
  console.log('\nAll done!');
}

main();
