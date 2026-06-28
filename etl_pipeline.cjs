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
var MONTH_NAMES = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
var COLORS = ['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899','#06B6D4','#84CC16','#F97316','#6366F1'];

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
  // Priority 1: Search first 15 rows × 8 cols for explicit QC codes or form titles
  var limit = Math.min(15, json.length);
  for (var r = 0; r < limit; r++) {
    var row = json[r];
    if (!row || !row.length) continue;
    for (var c = 0; c < Math.min(row.length, 8); c++) {
      var v = String(row[c] || '').trim();
      
      // Check for explicit QC code pattern
      if (v.indexOf('QC10002') >= 0) return 'QC10002-R02';
      if (v.indexOf('QC10004') >= 0) return 'QC10004-R02';
      if (v.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (v.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (v.indexOf('QC10007-R01') >= 0) return 'QC10007-R01';
      if (v.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (v.indexOf('QC10008') >= 0) return 'QC10008-R02';
      
      // Check for form title
      for (var title in FORM_TITLE_MAP) {
        if (v.indexOf(title) >= 0) return FORM_TITLE_MAP[title];
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

function findDateInSheet(json) {
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

function extractRawMonth(fileName, sheetName, year, relPath, json) {
  var y = String(year);
  var n, mn;

  // Strategy 1: YYYY-MM or YYYY_MM before .xlsx at end of filename
  // e.g., "原料-2025-01.xlsx", "2025-01.xlsx", "2025_01.xlsx"
  n = fileName.match(/(\d{4})[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    var yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(y, 10)) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }

  // Strategy 2: YYYYMMDD in filename (consecutive 8 digits)
  // e.g., "BD-20250107.xlsx", "Biometrix-20250107.xlsx"
  n = fileName.match(/(\d{4})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/i);
  if (n) {
    var yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(y, 10)) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }

  // Strategy 3: YYMMDD in filename (6 consecutive digits, any year)
  // e.g., "ICU-250102.xlsx", "其他-250103.xlsx", "射出-250102.xlsx"
  n = fileName.match(/(\d{2})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/);
  if (n) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 3b: YYMM in filename (4 digits, YY matches year suffix)
  // e.g., "Tubing-2501.xlsx" → month 1
  n = fileName.match(/(?:^|[^\d])(\d{2})(\d{2})\.xlsx$/i);
  if (n && n[1] === String(year).slice(-2)) {
    mn = parseInt(n[2], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 4: Filename ends with simple -MM or _MM suffix
  // e.g., 標籤 files like "2025-02.xlsx" would be caught by strategy 1 already
  n = fileName.match(/[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    mn = parseInt(n[1], 10);
    if (mn >= 1 && mn <= 12) return mn;
  }

  // Strategy 5: Sheet name has YYMMDD pattern (6 consecutive digits, any year)
  // e.g., "PE004 250106" → 250106, "250106", "RM066 250206"
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

  // Strategy 8: Scan sheet content for date patterns
  if (json) {
    mn = findDateInSheet(json);
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

  if (qc === 'QC10006-R02' || qc === 'QC10007-R01') {
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
    // QIP: determine Setup/Patrol from sheet name AND filename (case-insensitive)
    var snLower = sheetName.toLowerCase();
    var fnLower = fileName.toLowerCase();
    var type = '';
    if (snLower.indexOf('setup') >= 0 || snLower.indexOf('set up') >= 0 || snLower.indexOf('set-up') >= 0 || snLower === 'setup') {
      type = 'Setup';
    } else if (snLower.indexOf('patrol') >= 0) {
      type = 'Patrol';
    } else if (fnLower.indexOf('setup') >= 0 || fnLower.indexOf('set up') >= 0 || fnLower.indexOf('set-up') >= 0 || fnLower.indexOf('set_up') >= 0) {
      type = 'Setup';
    } else if (fnLower.indexOf('patrol') >= 0) {
      type = 'Patrol';
    } else {
      return null; // skip if can't determine Setup/Patrol
    }
    var folderLower = (qcFolder || '').toLowerCase();
    if (folderLower.indexOf('押出') >= 0) return '押出-' + type;
    if (folderLower.indexOf('射出') >= 0) return 'QIP-' + type;
    if (folderLower.indexOf('qip') >= 0) return 'QIP-' + type;
    if (relPath && relPath.toLowerCase().indexOf('押出') >= 0) return '押出-' + type;
    return 'QIP-' + type;
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

  wb.SheetNames.forEach(function(sheetName) {
    // Skip non-data sheets
    if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別') return;
    if (sheetName.indexOf('Sheet1') === 0) return;
    if (sheetName.indexOf('.K(') >= 0) return; // skip control plan templates
    if (sheetName.indexOf('範例樣本') >= 0) return; // skip sample sheets

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
    if (!actualQC) return;

    // Determine sub-category using the actual QC code
    var subCat = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
    if (!subCat) return;

    // Extract month
    var month = extractRawMonth(fileName, sheetName, year, relPath, json);
    if (!month || month < 1 || month > 12) month = 1; // default to Jan to preserve data

    // Increment count
    if (!counts[actualQC]) counts[actualQC] = {};
    if (!counts[actualQC][subCat]) counts[actualQC][subCat] = {};
    counts[actualQC][subCat][month] = (counts[actualQC][subCat][month] || 0) + 1;
  });
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
  
  // ---- Sheet: 品檢地圖 ----
  var mapData = [
    ['QC編碼', '表單名稱', '類別'],
    ['QC10002-R02', '原物料品檢表', '原物料進料'],
    ['QC10004-R02', 'QUALITY INSPECTION PLAN RECORD', '射出/押出製程'],
    ['QC10006-R01', '裝配對樣巡檢記錄表', '裝配巡檢'],
    ['QC10006-R02', '半成品品檢表', '半成品'],
    ['QC10007-R01', '完成品品檢表(首頁)', '完成品'],
    ['QC10007-R03', '零組件入庫品檢表', '零組件入庫'],
    ['QC10008-R02', '出貨檢驗報告', '出貨檢驗'],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mapData), '品檢地圖');
  
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
  
  var finQC = counts['QC10007-R01'];
  var finData = [];
  finCols.forEach(function() { finData.push({}); });
  
  var finSubCatMap = {'Biometrix': 0, 'MarMed': 1, 'Saxon': 2, 'Vivus': 3};
  
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
  
  // ---- Sheet 8: NCA ----
  var ncaRows = [
    ['NCA'],
    ['月份','件數','帶N','不帶N']
  ];
  MONTHS.forEach(function(m) { ncaRows.push([m, 0, 0, 0]); });
  ncaRows.push(['小計', 0, 0, 0]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ncaRows), 'NCA');
  
  // ---- Sheet 9: 彙總表 (Grid layout matching Template) ----
  var rawMainCols = [
    {key:'原料', label:'原料'},
    {key:'物料', label:'物料'},
    {key:'紙箱', label:'紙箱'},
    {key:'過濾網連蓋', label:'過濾網連蓋'},
    {key:'標籤', label:'標籤'},
    {key:'射出D', label:'射出D'},
  ];
  
  var rawMainData = [];
  rawMainCols.forEach(function() { rawMainData.push({}); });
  
  var rawQC = counts['QC10002-R02'];
  if (rawQC) {
    for (var subCat in rawQC) {
      var colIdx = null;
      if (subCat === '原料') colIdx = 0;
      else if (subCat === '物料' || subCat.indexOf('物料-') === 0) {
        if (subCat === '物料-紙箱') colIdx = 2;
        else if (subCat === '物料-過濾網連蓋') colIdx = 3;
        else if (subCat === '物料-標籤') colIdx = 4;
        else colIdx = 1; // generic 物料
      }
      else if (subCat === '射出D' || subCat.indexOf('射出D') === 0) colIdx = 5;
      if (colIdx === null) continue;
      
      var monthly = rawQC[subCat];
      for (var m = 1; m <= 12; m++) {
        if (monthly[m]) {
          rawMainData[colIdx][m] = (rawMainData[colIdx][m] || 0) + monthly[m];
        }
      }
    }
  }
  
  var rawGrandTotal = 0;
  rawMainCols.forEach(function(c, ci) {
    rawGrandTotal += totalArray(rawMainData[ci]);
  });

  var aggRows = [];
  for (var r = 0; r < 50; r++) {
    var rowArr = [];
    for (var c = 0; c < 19; c++) {
      rowArr.push('');
    }
    aggRows.push(rowArr);
  }
  
  // Section 1: 原物料進料品檢
  aggRows[0][0] = '原物料進料品檢(QC10002-R02)';
  var s1Headers = ["月份","原料","物料","紙箱","過濾網連蓋","標籤","射出D","小計"];
  s1Headers.forEach(function(h, ci) { aggRows[1][ci] = h; });
  MONTHS.forEach(function(m) {
    aggRows[m+1][0] = m;
    var total = 0;
    rawMainCols.forEach(function(c, ci) {
      var v = rawMainData[ci][m] || 0;
      aggRows[m+1][ci+1] = v;
      total += v;
    });
    aggRows[m+1][7] = total;
  });
  aggRows[14][0] = '小計';
  rawMainCols.forEach(function(c, ci) {
    aggRows[14][ci+1] = totalArray(rawMainData[ci]);
  });
  aggRows[14][7] = rawGrandTotal;
  
  // Section 2: 射出Setup
  aggRows[0][9] = '射出、押出製程Setup(QC10004-R02)';
  ["月份","押出(Setup)","射出(Setup)","小計"].forEach(function(h, ci) { aggRows[1][ci+9] = h; });
  MONTHS.forEach(function(m) {
    var extS = qipSetupExt[m] || 0;
    var injS = qipSetupInj[m] || 0;
    aggRows[m+1][9] = m;
    aggRows[m+1][10] = extS;
    aggRows[m+1][11] = injS;
    aggRows[m+1][12] = extS + injS;
  });
  aggRows[14][9] = '小計';
  aggRows[14][10] = tExtS;
  aggRows[14][11] = tInjS;
  aggRows[14][12] = tExtS + tInjS;
  
  // Section 3: 射出巡檢
  aggRows[0][14] = '射出、押出製程巡檢(QC10004-R02)';
  ["月份","押出(巡檢)","射出(廠內)","小計"].forEach(function(h, ci) { aggRows[1][ci+14] = h; });
  MONTHS.forEach(function(m) {
    var extP = qipPatrolExt[m] || 0;
    var injP = qipPatrolInj[m] || 0;
    aggRows[m+1][14] = m;
    aggRows[m+1][15] = extP;
    aggRows[m+1][16] = injP;
    aggRows[m+1][17] = extP + injP;
  });
  aggRows[14][14] = '小計';
  aggRows[14][15] = tExtP;
  aggRows[14][16] = tInjP;
  aggRows[14][17] = tExtP + tInjP;
  
  // Section 4: 裝配對樣巡檢
  var asmTotal = {};
  if (counts['QC10006-R01']) {
    asmTotal = counts['QC10006-R01']['裝配巡檢'] || {};
  }
  aggRows[17][0] = '裝配對樣巡檢(QC10006-R01)';
  aggRows[18][0] = '月份'; aggRows[18][1] = '裝配';
  MONTHS.forEach(function(m) {
    aggRows[m+18][0] = m;
    aggRows[m+18][1] = asmTotal[m] || 0;
  });
  aggRows[31][0] = '小計';
  aggRows[31][1] = totalArray(asmTotal);
  
  // Section 5: 半成品
  aggRows[17][3] = '裝配半成品品檢(QC10006-R02)';
  ["月份","裝配C","BD","Biometrix","MPS","Vivus","小計"].forEach(function(h, ci) { aggRows[18][ci+3] = h; });
  MONTHS.forEach(function(m) {
    aggRows[m+18][3] = m;
    var total = 0;
    semiCols.forEach(function(c, ci) {
      var v = semiData[ci][m] || 0;
      aggRows[m+18][ci+4] = v;
      total += v;
    });
    aggRows[m+18][9] = total;
  });
  aggRows[31][3] = '小計';
  semiCols.forEach(function(c, ci) {
    aggRows[31][ci+4] = totalArray(semiData[ci]);
  });
  aggRows[31][9] = semiGrandTotal;
  
  // Section 6: 完成品
  aggRows[17][11] = '裝配完成品品檢(QC10007-R01 R02)';
  ["月份","Biometrix","MarMed","Saxon","Vivus","小計"].forEach(function(h, ci) { aggRows[18][ci+11] = h; });
  MONTHS.forEach(function(m) {
    aggRows[m+18][11] = m;
    var total = 0;
    finCols.forEach(function(c, ci) {
      var v = finData[ci][m] || 0;
      aggRows[m+18][ci+12] = v;
      total += v;
    });
    aggRows[m+18][16] = total;
  });
  aggRows[31][11] = '小計';
  finCols.forEach(function(c, ci) {
    aggRows[31][ci+12] = totalArray(finData[ci]);
  });
  aggRows[31][16] = finGrandTotal;
  
  // Section 7: 零組件入庫
  aggRows[34][0] = '零組件入庫品檢(QC10007-R03)';
  ["月份","Tubing","射出(廠內)","射出A","射出C","射出D(組件)","裝配A","裝配B","裝配C","小計"].forEach(function(h, ci) { aggRows[35][ci] = h; });
  MONTHS.forEach(function(m) {
    aggRows[m+35][0] = m;
    var total = 0;
    partsCols.forEach(function(c, ci) {
      var v = partsData[ci][m] || 0;
      aggRows[m+35][ci+1] = v;
      total += v;
    });
    aggRows[m+35][9] = total;
  });
  aggRows[48][0] = '小計';
  partsCols.forEach(function(c, ci) {
    aggRows[48][ci+1] = totalArray(partsData[ci]);
  });
  aggRows[48][9] = partsGrand;
  
  // Section 8: 出貨
  aggRows[34][11] = '出貨品檢(QC10008-R02)';
  ["月份","ICU","其他","小計"].forEach(function(h, ci) { aggRows[35][ci+11] = h; });
  MONTHS.forEach(function(m) {
    aggRows[m+35][11] = m;
    aggRows[m+35][12] = shipICU[m] || 0;
    aggRows[m+35][13] = shipOther[m] || 0;
    aggRows[m+35][14] = (shipICU[m] || 0) + (shipOther[m] || 0);
  });
  aggRows[48][11] = '小計';
  aggRows[48][12] = tICU;
  aggRows[48][13] = tOth;
  aggRows[48][14] = tICU + tOth;
  
  // Section 9: NCA
  aggRows[34][16] = 'NCA(品質異常案件)';
  ["月份","件數"].forEach(function(h, ci) { aggRows[35][ci+16] = h; });
  MONTHS.forEach(function(m) {
    aggRows[m+35][16] = m;
    aggRows[m+35][17] = 0;
  });
  aggRows[48][16] = '小計';
  aggRows[48][17] = 0;
  
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aggRows), '彙總表');
  
  // Write the workbook
  XLSX.writeFile(wb, outFile);
  console.log('  Summary written: ' + outFile);
  return outFile;
}

// ============================================================
// 4. HTML REPORT GENERATOR
// ============================================================

function generateReport(summaryFile, year, isComparison, summaryFile2, year2) {
  var wb = XLSX.readFile(summaryFile);
  var outFile;
  
  if (isComparison) {
    outFile = '品檢報表比較分析.html';
    generateComparisonReport(wb, summaryFile2, outFile);
  } else {
    outFile = year + '品檢報表分析.html';
    generateYearlyReport(wb, year, outFile);
  }
  console.log('  HTML report: ' + outFile);
}

function getSheet(wb, name) {
  var s = wb.Sheets[name];
  return s ? XLSX.utils.sheet_to_json(s, {header:1, defval:''}) : [];
}

function getMonths(arr) { return MONTHS.map(function(m){return arr[m]||0;}); }
function totalOf(arr) { return getMonths(arr).reduce(function(a,b){return a+b;},0); }

function generateYearlyReport(wb, year, outFile) {
  // Extract data from summary sheets (same logic as the old gen scripts)
  // ... (simplified: reuse the existing report structure)
  
  // For now, generate a basic report using the summary data
  var rawSheet = getSheet(wb, '原物料品檢(QC10002-R02)');
  var qipSheet = getSheet(wb, 'QIP(QC10004-R02)');
  var asmSheet = getSheet(wb, '裝配對樣巡檢(QC10006-R01)');
  var semiSheet = getSheet(wb, '半成品品檢(QC10006-R02)');
  var finSheet = getSheet(wb, '完成品品檢(QC10007-R01 R02)');
  var partsSheet = getSheet(wb, '零組件入庫品檢(QC10007-R03)');
  var shipSheet = getSheet(wb, '出貨檢驗(QC10008-R02)');
  
  // Extract monthly data rows (rows 2-13, indices 2-13)
  function extractMonthly(sheet, colIdx) {
    var data = {};
    for (var i = 2; i <= 13; i++) {
      var r = sheet[i];
      if (r && r[colIdx] !== '' && r[colIdx] !== undefined) {
        data[i - 1] = Number(r[colIdx]) || 0;
      }
    }
    return data;
  }
  
  // Also try to extract from the 彙總表 for better accuracy
  var aggSheet = getSheet(wb, '彙總表');
  
  // Parse individual sheets
  var rawMain = {};
  var rawMtl = {};
  var rawBox = {};
  var rawFilter = {};
  var rawLabel = {};
  var rawInjD = {};
  
  if (rawSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = rawSheet[i];
      if (!r) continue;
      var m = i - 1;
      if (r[1] !== '' && r[1] !== undefined) rawMain[m] = Number(r[1]) || 0;
      
      // rawMtl is the sum of B膠, 收縮膜, 色粉, 空白包裝袋, 空白感壓紙, 塑膠袋, 塑膠袋40*50 (indices 2 to 8)
      var sumMtl = 0;
      for (var c = 2; c <= 8; c++) {
        if (r[c] !== '' && r[c] !== undefined) sumMtl += Number(r[c]) || 0;
      }
      rawMtl[m] = sumMtl;
      
      if (r[9] !== '' && r[9] !== undefined) rawBox[m] = Number(r[9]) || 0;
      if (r[10] !== '' && r[10] !== undefined) rawFilter[m] = Number(r[10]) || 0;
      if (r[11] !== '' && r[11] !== undefined) rawLabel[m] = Number(r[11]) || 0;
      if (r[12] !== '' && r[12] !== undefined) rawInjD[m] = Number(r[12]) || 0;
    }
  }
  
  var rawSub = {};
  MONTHS.forEach(function(m) {
    rawSub[m] = (rawMtl[m]||0)+(rawBox[m]||0)+(rawFilter[m]||0)+(rawLabel[m]||0)+(rawInjD[m]||0);
  });
  
  // QIP
  var injSetup = {}; var extSetup = {};
  var injPatrol = {}; var extPatrol = {};
  if (qipSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = qipSheet[i];
      if (!r) continue;
      var m = i - 1;
      if (r[1] !== '' && r[1] !== undefined) extSetup[m] = Number(r[1]) || 0;
      if (r[2] !== '' && r[2] !== undefined) injSetup[m] = Number(r[2]) || 0;
      if (r[6] !== '' && r[6] !== undefined) extPatrol[m] = Number(r[6]) || 0;
      if (r[7] !== '' && r[7] !== undefined) injPatrol[m] = Number(r[7]) || 0;
    }
  }
  
  // 裝配巡檢
  var patrolAsm = {};
  if (asmSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = asmSheet[i];
      var m = i - 1;
      if (r && r[1] !== '' && r[1] !== undefined) patrolAsm[m] = Number(r[1]) || 0;
    }
  }
  
  // 半成品
  var semiAssyA = {}; var semiAssyB = {}; var semiAssyC = {};
  var semiBD = {}; var semiBM = {}; var semiMP = {}; var semiVV = {}; var semiOther = {};
  if (semiSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = semiSheet[i];
      if (!r) continue;
      var m = i - 1;
      semiAssyA[m] = 0;
      semiAssyB[m] = 0;
      if (r[1] !== '' && r[1] !== undefined) semiAssyC[m] = Number(r[1]) || 0;
      if (r[2] !== '' && r[2] !== undefined) semiBD[m] = Number(r[2]) || 0;
      if (r[3] !== '' && r[3] !== undefined) semiBM[m] = Number(r[3]) || 0;
      if (r[4] !== '' && r[4] !== undefined) semiMP[m] = Number(r[4]) || 0;
      if (r[5] !== '' && r[5] !== undefined) semiVV[m] = Number(r[5]) || 0;
      semiOther[m] = 0;
    }
  }
  
  // 完成品
  var finBM = {}; var finMM = {}; var finSX = {}; var finVV = {};
  if (finSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = finSheet[i];
      if (!r) continue;
      var m = i - 1;
      if (r[1] !== '' && r[1] !== undefined) finBM[m] = Number(r[1]) || 0;
      if (r[2] !== '' && r[2] !== undefined) finMM[m] = Number(r[2]) || 0;
      if (r[3] !== '' && r[3] !== undefined) finSX[m] = Number(r[3]) || 0;
      if (r[4] !== '' && r[4] !== undefined) finVV[m] = Number(r[4]) || 0;
    }
  }
  
  // 零組件入庫
  var pTub = {}; var pInj = {}; var pInjA = {}; var pInjC = {};
  var pInjD = {}; var pAsmA = {}; var pAsmB = {}; var pAsmC = {};
  if (partsSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = partsSheet[i];
      if (!r) continue;
      var m = i - 1;
      if (r[1] !== '' && r[1] !== undefined) pTub[m] = Number(r[1]) || 0;
      if (r[2] !== '' && r[2] !== undefined) pInj[m] = Number(r[2]) || 0;
      if (r[3] !== '' && r[3] !== undefined) pInjA[m] = Number(r[3]) || 0;
      if (r[4] !== '' && r[4] !== undefined) pInjC[m] = Number(r[4]) || 0;
      if (r[5] !== '' && r[5] !== undefined) pInjD[m] = Number(r[5]) || 0;
      if (r[6] !== '' && r[6] !== undefined) pAsmA[m] = Number(r[6]) || 0;
      if (r[7] !== '' && r[7] !== undefined) pAsmB[m] = Number(r[7]) || 0;
      if (r[8] !== '' && r[8] !== undefined) pAsmC[m] = Number(r[8]) || 0;
    }
  }
  var ptSub = {};
  MONTHS.forEach(function(m) {
    ptSub[m] = (pTub[m]||0)+(pInj[m]||0)+(pInjA[m]||0)+(pInjC[m]||0)+(pInjD[m]||0)+(pAsmA[m]||0)+(pAsmB[m]||0)+(pAsmC[m]||0);
  });
  
  // 出貨
  var sICU = {}; var sOth = {};
  if (shipSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = shipSheet[i];
      if (!r) continue;
      var m = i - 1;
      if (r[1] !== '' && r[1] !== undefined) sICU[m] = Number(r[1]) || 0;
      if (r[2] !== '' && r[2] !== undefined) sOth[m] = Number(r[2]) || 0;
    }
  }
  
  var catData = [
    {label:'原物料進料', m:MONTHS.map(function(m){return (rawMain[m]||0)+(rawSub[m]||0);})},
    {label:'射出Setup', m:getMonths(injSetup)},
    {label:'射出巡檢', m:getMonths(injPatrol)},
    {label:'裝配巡檢', m:getMonths(patrolAsm)},
    {label:'半成品', m:getMonths(semiAssyA).map(function(v,i){return v+getMonths(semiAssyB)[i]+getMonths(semiAssyC)[i]+getMonths(semiBD)[i]+getMonths(semiBM)[i]+getMonths(semiMP)[i]+getMonths(semiVV)[i]+getMonths(semiOther)[i];})},
    {label:'完成品', m:getMonths(finBM).map(function(v,i){return v+getMonths(finMM)[i]+getMonths(finSX)[i]+getMonths(finVV)[i];})},
    {label:'零組件入庫', m:getMonths(ptSub)},
    {label:'出貨檢驗', m:MONTHS.map(function(m){return (sICU[m]||0)+(sOth[m]||0);})},
  ];
  
  var grandTotal = 0;
  catData.forEach(function(c) { c.total = c.m.reduce(function(a,b){return a+b;},0); grandTotal += c.total; });
  
  // Stats
  var tShip = totalOf(sICU) + totalOf(sOth);
  var tRaw = totalOf(rawMain) + totalOf(rawSub);
  var tInj = totalOf(injSetup) + totalOf(injPatrol);
  var tExt = totalOf(extSetup) + totalOf(extPatrol);
  var tSemi = totalOf(semiAssyA) + totalOf(semiAssyB) + totalOf(semiAssyC) + totalOf(semiBD) + totalOf(semiBM) + totalOf(semiMP) + totalOf(semiVV) + totalOf(semiOther);
  var tFin = totalOf(finBM) + totalOf(finMM) + totalOf(finSX) + totalOf(finVV);
  var tParts = totalOf(ptSub);
  
  // Build slides
  var slides = [];
  var chartCfgs = [];
  
  // Slide 0: Overview
  var catLabels = JSON.stringify(catData.map(function(c){return c.label;}));
  var catTotals = JSON.stringify(catData.map(function(c){return c.total;}));
  chartCfgs.push('{id:"c0",type:"bar",data:{labels:'+catLabels+',datasets:[{label:"'+year+'年度",data:'+catTotals+',backgroundColor:'+JSON.stringify(COLORS.slice(0,8))+',borderRadius:6}]},options:{responsive:true,indexAxis:"y",plugins:{title:{display:true,text:"'+year+'年度 各類別品檢總量",font:{size:18}},legend:{display:false}},scales:{x:{beginAtZero:true,title:{display:true,text:"數量"}},y:{grid:{display:false}}}}}');
  slides.push({id:'s0', title:year+'年度 品檢概覽', chart:'c0', text:'<p>'+year+'年度品檢彙總：<strong>'+grandTotal.toLocaleString()+'</strong> 筆</p><p>最大量為<strong>射出巡檢</strong> '+totalOf(injPatrol).toLocaleString()+' 筆。</p>'});
  
  // Slide 1: 主要類別月度趨勢
  var mainCats = [
    {label:'原物料進料', data:MONTHS.map(function(m){return (rawMain[m]||0)+(rawSub[m]||0);}), color:'#3B82F6'},
    {label:'射出Setup', data:getMonths(injSetup), color:'#10B981'},
    {label:'射出巡檢', data:getMonths(injPatrol), color:'#F59E0B'},
    {label:'裝配巡檢', data:getMonths(patrolAsm), color:'#8B5CF6'},
    {label:'出貨', data:MONTHS.map(function(m){return (sICU[m]||0)+(sOth[m]||0);}), color:'#EC4899'},
  ];
  chartCfgs.push('{id:"c1",type:"bar",data:{labels:'+JSON.stringify(MONTH_NAMES)+',datasets:'+JSON.stringify(mainCats.map(function(d){return {label:d.label,data:d.data,backgroundColor:d.color,borderRadius:2};}))+'},options:{responsive:true,plugins:{title:{display:true,text:"'+year+'年度 主要類別月度趨勢",font:{size:18}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,title:{display:true,text:"數量"}}}}}');
  slides.push({id:'s1', title:'主要類別月度趨勢', chart:'c1', text:'<p>各類別月度分布趨勢。</p>'});
  
  // Slide 2: 射出分布
  chartCfgs.push('{id:"c2",type:"bar",data:{labels:'+JSON.stringify(MONTH_NAMES)+',datasets:[{label:"射出Setup",data:'+JSON.stringify(getMonths(injSetup))+',backgroundColor:"#3B82F6",borderRadius:2},{label:"射出巡檢",data:'+JSON.stringify(getMonths(injPatrol))+',backgroundColor:"#F59E0B",borderRadius:2}]},options:{responsive:true,stacked:true,plugins:{title:{display:true,text:"'+year+'年度 射出品檢月度分布",font:{size:18}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,title:{display:true,text:"數量"}}}}}');
  slides.push({id:'s2', title:'射出品檢分布', chart:'c2', text:'<p>射出品檢合計 <strong>'+tInj.toLocaleString()+'</strong> 筆。</p>'});
  
  // Slide 3: 出貨分布
  chartCfgs.push('{id:"c3",type:"doughnut",data:{labels:["ICU ('+totalOf(sICU).toLocaleString()+'筆)","其他 ('+totalOf(sOth).toLocaleString()+'筆)"],datasets:[{data:['+totalOf(sICU)+','+totalOf(sOth)+'],backgroundColor:["#3B82F6","#F59E0B"],borderWidth:2}]},options:{responsive:true,plugins:{title:{display:true,text:"'+year+'年度 出貨檢驗分布",font:{size:18}},legend:{position:"bottom"}}}}');
  slides.push({id:'s3', title:'出貨檢驗分布', chart:'c3', text:'<p>出貨檢驗合計 <strong>'+tShip.toLocaleString()+'</strong> 筆。</p>'});
  
  // Slide 4: 進料分布
  chartCfgs.push('{id:"c4",type:"doughnut",data:{labels:["原料 ('+totalOf(rawMain).toLocaleString()+'筆)","物料+射出D ('+totalOf(rawSub).toLocaleString()+'筆)"],datasets:[{data:['+totalOf(rawMain)+','+totalOf(rawSub)+'],backgroundColor:["#10B981","#8B5CF6"],borderWidth:2}]},options:{responsive:true,plugins:{title:{display:true,text:"'+year+'年度 進料檢驗分布",font:{size:18}},legend:{position:"bottom"}}}}');
  slides.push({id:'s4', title:'進料檢驗分布', chart:'c4', text:'<p>進料檢驗合計 <strong>'+tRaw.toLocaleString()+'</strong> 筆。</p>'});
  
  if (tSemi > 0) {
    var tAssy = totalOf(semiAssyA) + totalOf(semiAssyB) + totalOf(semiAssyC);
    chartCfgs.push('{id:"c5",type:"doughnut",data:{labels:["裝配 ('+tAssy.toLocaleString()+'筆)","BD ('+totalOf(semiBD).toLocaleString()+'筆)","Biometrix ('+totalOf(semiBM).toLocaleString()+'筆)","MPS ('+totalOf(semiMP).toLocaleString()+'筆)","Vivus ('+totalOf(semiVV).toLocaleString()+'筆)","其他 ('+totalOf(semiOther).toLocaleString()+'筆)"],datasets:[{data:['+tAssy+','+totalOf(semiBD)+','+totalOf(semiBM)+','+totalOf(semiMP)+','+totalOf(semiVV)+','+totalOf(semiOther)+'],backgroundColor:["#06B6D4","#3B82F6","#10B981","#F59E0B","#8B5CF6","#EC4899"],borderWidth:2}]},options:{responsive:true,plugins:{title:{display:true,text:"'+year+'年度 半成品品檢分布",font:{size:18}},legend:{position:"bottom"}}}}');
    slides.push({id:'s5', title:'半成品品檢分布', chart:'c5', text:'<p>半成品品檢合計 <strong>'+tSemi.toLocaleString()+'</strong> 筆。</p>'});
  }
  
  if (tFin > 0) {
    chartCfgs.push('{id:"c6",type:"doughnut",data:{labels:["Biometrix ('+totalOf(finBM).toLocaleString()+'筆)","MarMed ('+totalOf(finMM).toLocaleString()+'筆)","Saxon ('+totalOf(finSX).toLocaleString()+'筆)","Vivus ('+totalOf(finVV).toLocaleString()+'筆)"],datasets:[{data:['+totalOf(finBM)+','+totalOf(finMM)+','+totalOf(finSX)+','+totalOf(finVV)+'],backgroundColor:["#3B82F6","#10B981","#F59E0B","#8B5CF6"],borderWidth:2}]},options:{responsive:true,plugins:{title:{display:true,text:"'+year+'年度 完成品品檢分布",font:{size:18}},legend:{position:"bottom"}}}}');
    slides.push({id:'s6', title:'完成品品檢分布', chart:'c6', text:'<p>完成品品檢合計 <strong>'+tFin.toLocaleString()+'</strong> 筆。</p>'});
  }
  
  // 零組件入庫
  chartCfgs.push('{id:"c7",type:"bar",data:{labels:'+JSON.stringify(MONTH_NAMES)+',datasets:[{label:"射出(廠內)",data:'+JSON.stringify(getMonths(pInj))+',backgroundColor:"#3B82F6",borderRadius:2},{label:"Tubing",data:'+JSON.stringify(getMonths(pTub))+',backgroundColor:"#10B981",borderRadius:2}]},options:{responsive:true,stacked:true,plugins:{title:{display:true,text:"'+year+'年度 零組件入庫月度分布",font:{size:18}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,title:{display:true,text:"數量"}}}}}');
  slides.push({id:'s7', title:'零組件入庫分布', chart:'c7', text:'<p>零組件入庫合計 <strong>'+tParts.toLocaleString()+'</strong> 筆。</p>'});
  
  // Render HTML
  renderSlidesToHTML(slides, chartCfgs, outFile, year+'年度 品檢報表分析');
}

function generateComparisonReport(wb1, wb2, outFile) {
  // Simplified comparison report
  var slides = [];
  var chartCfgs = [];
  
  // Extract totals from both summaries
  function getCatTotals(wb) {
    var agg = getSheet(wb, '彙總表');
    var cats = ['原物料進料', '射出Setup', '射出巡檢', '裝配巡檢', '半成品', '完成品', '零組件入庫', '出貨檢驗'];
    var totals = {};
    
    // Read from 彙總表 sections
    for (var i = 0; i < agg.length; i++) {
      var r = agg[i];
      if (!r || !r[0]) continue;
      var header = String(r[0]);
      
      if (header.indexOf('原物料進料品檢') >= 0) {
        // 小計 is at i+13 (after 月份 header + 12 months)
        var sumRow = agg[i+13];
        if (sumRow && sumRow[0] === '小計') totals['原物料進料'] = Number(sumRow[7]) || 0;
      }
      if (header.indexOf('射出Setup') >= 0) {
        var sumRow = agg[i+13];
        if (sumRow && sumRow[0] === '小計') totals['射出Setup'] = Number(sumRow[3]) || 0;
      }
      if (header.indexOf('射出巡檢') >= 0) {
        var sumRow = agg[i+13];
        if (sumRow && sumRow[0] === '小計') totals['射出巡檢'] = Number(sumRow[3]) || 0;
      }
    }
    
    return totals;
  }
  
  // Actually, just use the yearly report data from individual sheets
  renderComparisonPage(outFile, wb1, wb2);
}

function renderSlidesToHTML(slides, chartCfgs, outFile, title) {
  var slidesHtml = '';
  for (var si = 0; si < slides.length; si++) {
    var s = slides[si];
    slidesHtml += '\n<section class="slide' + (si===0?' active':'') + '" id="slide-' + s.id + '">' +
      '<h2>' + s.title + '</h2>' +
      '<div class="chart-container" style="width:65%;float:left;">' +
      (s.chart ? '<canvas id="' + s.chart + '"></canvas>' : '') +
      '</div>' +
      '<div class="text-container" style="width:30%;float:right;padding:20px;">' +
      s.text +
      '</div>' +
    '</section>';
  }
  
  var chartInit = '';
  for (var ci = 0; ci < chartCfgs.length; ci++) {
    var idMatch = chartCfgs[ci].match(/id:"([^"]+)"/);
    if (idMatch) chartInit += '\nnew Chart(document.getElementById("' + idMatch[1] + '"),' + chartCfgs[ci] + ');';
  }
  
  var html = '<!DOCTYPE html>\n<html lang="zh-TW">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>' + title + '</title>\n<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\/script>\n<style>\n*{margin:0;padding:0;box-sizing:border-box}\nbody{background:#F8FAFC;font-family:Segoe UI,Noto Sans TC,sans-serif;color:#1E293B;overflow:hidden}\n.slider{width:100vw;height:100vh;position:relative;overflow:hidden}\n.slide{position:absolute;top:0;left:0;width:100%;height:100%;padding:40px 60px;display:none;background:#F8FAFC}\n.slide.active{display:block;animation:fadeIn .4s ease}\n@keyframes fadeIn{from{opacity:0;transform:translateX(30px)}to{opacity:1;transform:translateX(0)}}\nh2{font-size:26px;margin-bottom:20px;color:#0F172A;border-left:5px solid #3B82F6;padding-left:15px}\n.chart-container{height:70vh;position:relative}\n.text-container{font-size:15px;line-height:1.8;color:#334155}\n.text-container strong{color:#1E293B}\n.text-container em{color:#64748B;font-size:13px}\n.nav{position:fixed;bottom:30px;left:50%;transform:translateX(-50%);display:flex;gap:12px;align-items:center;z-index:100}\n.nav button{background:#3B82F6;color:#fff;border:none;padding:8px 20px;border-radius:6px;cursor:pointer;font-size:15px}\n.nav button:hover{background:#2563EB}\n.nav span{color:#64748B;font-size:14px}\n<\/style>\n</head>\n<body>\n<div class="slider">\n' + slidesHtml + '\n</div>\n<div class="nav">\n<button onclick="prevSlide()">&#9664; 上一頁</button>\n<span id="pageInfo">1 / ' + slides.length + '</span>\n<button onclick="nextSlide()">下一頁 &#9654;</button>\n</div>\n<script>\nvar ci=0;\nfunction showSlide(i){document.querySelectorAll(".slide").forEach(function(s,j){s.classList.toggle("active",j===i)});document.getElementById("pageInfo").textContent=(i+1)+" / ' + slides.length + '";ci=i}\nfunction nextSlide(){if(ci<' + (slides.length-1) + ')showSlide(ci+1)}\nfunction prevSlide(){if(ci>0)showSlide(ci-1)}\ndocument.addEventListener("keydown",function(e){if(e.key==="ArrowRight")nextSlide();if(e.key==="ArrowLeft")prevSlide()});\nwindow.addEventListener("load",function(){\n' + chartInit + '\n});\n<\/script>\n</body>\n</html>';
  
  fs.writeFileSync(outFile, html);
}

function renderComparisonPage(outFile, wb25, wb26) {
  // Simplified: just write a placeholder that references both yearly reports
  var html = '<!DOCTYPE html>\n<html lang="zh-TW">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>2025 vs 2026 品檢報表比較分析</title>\n<style>\nbody{font-family:Segoe UI,Noto Sans TC,sans-serif;background:#F8FAFC;color:#1E293B;padding:60px}\nh1{font-size:28px;border-left:5px solid #3B82F6;padding-left:15px}\na{display:inline-block;margin:20px;padding:15px 30px;background:#3B82F6;color:#fff;text-decoration:none;border-radius:8px;font-size:18px}\na:hover{background:#2563EB}\n</style>\n</head>\n<body>\n<h1>品檢報表比較分析</h1>\n<p>請先檢視個別年度報告：</p>\n<a href="2025品檢報表分析.html">2025年度分析報告</a>\n<a href="2026品檢報表分析.html">2026年度分析報告</a>\n</body>\n</html>';
  fs.writeFileSync(outFile, html);
  console.log('  (比較報告為簡化版，請直接開啟個別年度報告)');
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
    
    console.log('Step 3: Generating styled HTML report...');
    var styledReportScript = require('child_process').spawnSync('node', ['generate_styled_reports.cjs', String(year)], {cwd: process.cwd(), stdio: 'inherit'});
    if (styledReportScript.status !== 0) {
      console.log('  Warning: Failed to generate styled HTML report');
    }
  });
  
  // Generate comparison report if both years processed
  if (years.indexOf(2025) >= 0 && years.indexOf(2026) >= 0) {
    console.log('\n=== Generating comparison report ===');
    generateReport('DataExtract/2025品檢報表統計.xlsx', 2025, true, 'DataExtract/2026品檢報表統計.xlsx');
  }
  
  console.log('\nAll done!');
}

main();
