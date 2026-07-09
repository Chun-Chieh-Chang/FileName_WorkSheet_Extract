import * as XLSX from 'xlsx';

const MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
const LETTER_MONTH = { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, I: 9, J: 10, K: 11, L: 12 };

const FOLDER_QC_MAP = {
  '半成品品檢表': 'QC10006-R02',  // Keyword-driven: matches filename OR folder path (PRIORITY over 零組件入庫)
  '原物料品檢': 'QC10002-R02',
  '進料檢驗': 'QC10002-R02',
  '出貨檢驗': 'QC10008-R02',
  '裝配檢驗': 'QC10006-R02',
  '裝配巡檢': 'QC10006-R01',
  '零組件入庫': 'QC10007-R03',
  '完成品品檢': 'QC10007-R01',
  'QIP尺寸檢驗': 'QC10004-R02',
  '射出檢驗': 'QC10004-R02',
  '押出檢驗': 'QC10004-R02'
};

const FORM_TITLE_MAP = {
  '原物料/配件進料品檢表': 'QC10002-R02',
  '進料檢驗紀錄表': 'QC10002-R02',
  '裝配對樣巡檢記錄表': 'QC10006-R01',
  '半成品檢驗記錄表': 'QC10006-R02',
  '半成品巡檢品檢表': 'QC10006-R02',
  'SUB-ASSEMBLED SETS QUALITY INSPECTION PLAN': 'QC10006-R02',
  '完成品裝配品檢紀錄表': 'QC10007-R01',
  '完成品裝配品檢記錄表': 'QC10007-R01',
  'FINISHED SETS QUALITY INSPECTION PLAN': 'QC10007-R01',
  '零組件入庫品檢表': 'QC10007-R03',
  '出貨品檢記錄表': 'QC10008-R02',
  'OUT-GOING QUALITY INSPECTION PLAN': 'QC10008-R02',
  '出貨品檢報告': 'QC10008-R02'
};

export function normalizeScientificNotation(val) {
  if (val === undefined || val === null) return "";
  const str = String(val).trim();
  
  // 1. 處理被誤讀為浮點數的科學記號，例如 260101E-2 變成 2601.01，260101E-3 變成 2601.001
  const decimalMatch = str.match(/^(\d{4})\.(\d+)$/);
  if (decimalMatch) {
    const decimalPart = decimalMatch[2];
    const exp = decimalPart.length;
    const numValue = parseFloat(str);
    const multiplier = Math.pow(10, exp);
    const reconstructedBase = Math.round(numValue * multiplier);
    if (String(reconstructedBase).length === 6) {
      return `${reconstructedBase}E-${exp}`;
    }
  }
  
  // 2. 處理標準科學記號字串，例如 2.60101e+5 或 2.60101E+5
  const sciMatch = str.match(/^(\d)\.(\d+)e\+?(\d+)$/i);
  if (sciMatch) {
    const numValue = parseFloat(str);
    const rounded = Math.round(numValue);
    if (String(rounded).length === 6) {
      return String(rounded);
    }
  }
  
  return str;
}

export function detectQCFromFolder(dirname) {
  const m = dirname.match(/QC\d{5}-R\d{2}/i);
  if (m) return m[0].toUpperCase();
  const keys = Object.keys(FOLDER_QC_MAP);
  for (let i = 0; i < keys.length; i++) {
    if (dirname.indexOf(keys[i]) >= 0) return FOLDER_QC_MAP[keys[i]];
  }
  return null;
}

export function parseDateFromString(str) {
  if (!str) return null;
  str = normalizeScientificNotation(String(str).trim());
  let m = str.match(/\b(20\d{2})[-/](\d{1,2})[-/]\d{1,2}\b/);
  if (m) return { year: parseInt(m[1], 10), month: parseInt(m[2], 10) };
  
  let m2 = str.match(/\b(\d{2})[-/](\d{1,2})[-/]\d{1,2}\b/);
  if (m2) return { year: 2000 + parseInt(m2[1], 10), month: parseInt(m2[2], 10) };
  
  let m3 = str.match(/\b(\d{2})(\d{2})(\d{2})[A-Za-z]?\b/);
  if (m3) {
    const mm = parseInt(m3[2], 10);
    if (mm >= 1 && mm <= 12) return { year: 2000 + parseInt(m3[1], 10), month: mm };
  }
  return null;
}

export function parseDateFromValue(val, formatted) {
  if (val instanceof Date) {
    return { year: val.getFullYear(), month: val.getMonth() + 1 };
  }
  if (typeof val === 'number' && val >= 40000 && val <= 50000) {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(date.getTime())) {
      return { year: date.getFullYear(), month: date.getMonth() + 1 };
    }
  }
  return parseDateFromString(formatted || val);
}

export function findDateInSheet(ws, qc) {
  if (!ws) return null;

  const getCellValAndFormatted = (addr) => {
    const cell = ws[addr];
    if (!cell) return { val: null, formatted: null };
    return { val: cell.v, formatted: cell.w || '' };
  };

  let cellInfo;
  let dateInfo = null;

  switch (qc) {
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

export function findDateInSheetFallback(json) {
  const limit = Math.min(20, json.length);
  for (let r = 0; r < limit; r++) {
    const row = json[r];
    if (!row || !row.length) continue;
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || '');
      let d = v.match(/(20\d{2})[-/](\d{1,2})[-/]\d{1,2}/);
      if (d) { const mn = parseInt(d[2], 10); if (mn >= 1 && mn <= 12) return mn; }
      d = v.match(/(\d{2})[-/](\d{1,2})[-/]\d{1,2}/);
      if (d) { const mn = parseInt(d[2], 10); if (mn >= 1 && mn <= 12) return mn; }
      d = v.match(/(\d{2})(\d{2})(\d{2})/);
      if (d) { const mn = parseInt(d[2], 10); if (mn >= 1 && mn <= 12) return mn; }
      d = v.match(/(\d{1,2})月/);
      if (d) { const mn = parseInt(d[1], 10); if (mn >= 1 && mn <= 12) return mn; }
    }
  }
  return null;
}

export function determineQCFromSheet(json, initialQC, relPath) {
  if (initialQC === 'QC10006-R01') return 'QC10006-R01';
  if (initialQC === 'QC10004-R02') return 'QC10004-R02';

  const scanLimit = Math.min(15, json.length);
  for (let ri = 0; ri < scanLimit; ri++) {
    const row = json[ri];
    if (!row) continue;
    
    // OPTIMIZATION: Check column A first (index 0), return immediately if QC code found
    const colA = String(row[0] || '').trim();
    if (colA) {
      if (colA.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (colA.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (colA.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (colA.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (colA.indexOf('QC10007-R01') >= 0 || colA.indexOf('QC10007-R02') >= 0) return 'QC10007-R01';
      if (colA.indexOf('QC10008') >= 0) return 'QC10008-R02';
      
      // Check for form title in column A (header only)
      if (ri === 0) {
        const titleKeys = Object.keys(FORM_TITLE_MAP);
        for (let k = 0; k < titleKeys.length; k++) {
          if (colA.indexOf(titleKeys[k]) >= 0) return FORM_TITLE_MAP[titleKeys[k]];
        }
      }
    }
    
    // If column A doesn't have QC code, check remaining columns (B-H, indices 1-7)
    for (let ci = 1; ci < row.length && ci < 8; ci++) {
      const val = String(row[ci] || '').trim();
      if (!val) continue;

      if (val.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (val.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (val.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (val.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (val.indexOf('QC10007-R01') >= 0 || val.indexOf('QC10007-R02') >= 0) return 'QC10007-R01';
      if (val.indexOf('QC10008') >= 0) return 'QC10008-R02';

      if (ri === 0 || ri === 1 || ri === 2) {
        const titleKeys = Object.keys(FORM_TITLE_MAP);
        for (let k = 0; k < titleKeys.length; k++) {
          if (val.indexOf(titleKeys[k]) >= 0) return FORM_TITLE_MAP[titleKeys[k]];
        }
      }
    }
  }

  // Scan footer (last 30 rows) for merged or hidden sheet footprints
  const footerStart = Math.max(0, json.length - 30);
  for (let ri = footerStart; ri < json.length; ri++) {
    const row = json[ri];
    if (!row) continue;
    
    // OPTIMIZATION: Check column A first
    const colA = String(row[0] || '').trim();
    if (colA) {
      if (colA.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (colA.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (colA.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (colA.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (colA.indexOf('QC10007-R01') >= 0 || colA.indexOf('QC10007-R02') >= 0) return 'QC10007-R01';
      if (colA.indexOf('QC10008') >= 0) return 'QC10008-R02';
    }
    
    // Check remaining columns
    for (let ci = 1; ci < row.length && ci < 8; ci++) {
      const val = String(row[ci] || '').trim();
      if (!val) continue;
      if (val.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (val.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (val.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (val.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (val.indexOf('QC10007-R01') >= 0 || val.indexOf('QC10007-R02') >= 0) return 'QC10007-R01';
      if (val.indexOf('QC10008') >= 0) return 'QC10008-R02';
    }
  }

  return initialQC;
}

export function getRawSubCategory(qc, relPath, fileName, sheetName, qcFolder) {
  if (qc === 'QC10002-R02') {
    const parts = relPath.split('/');
    const p0 = parts[0].replace(/[-_]\d{4}$/, '').replace(/\s+/g, '');
    if (p0 === '原料') return '原料';
    if (p0 === '物料') {
      if (parts.length > 1) {
        const sub = parts[1].replace(/[-_]\d{4}$/, '').replace(/\s+/g, '');
        if (sub) return '物料-' + sub;
        return '物料-' + parts[1];
      }
      let name = fileName.replace(/\.xlsx$/i, '');
      name = name.replace(/[-_]\d{4}[-_]\d{1,2}$/, '');
      name = name.replace(/[-_]\d{4}$/, '');
      name = name.replace(/[-_]\d{1,2}$/, '');
      name = name.replace(/\s+/g, '');
      return '物料-' + name;
    }
    if (parts[0].indexOf('射出D') >= 0) return '射出D';
    return null;
  }

  if (qc === 'QC10008-R02') {
    return relPath.replace(/[-_](20\d{2})$/, '');
  }

  if (qc === 'QC10006-R02' || qc === 'QC10007-R01' || qc === 'QC10007-R02') {
    return relPath.replace(/[-_](20\d{2})$/, '');
  }

  if (qc === 'QC10006-R01') {
    return '裝配巡檢';
  }

  if (qc === 'QC10007-R03') {
    const parts = relPath.split('/');
    const name = parts[0].replace(/[-_](20\d{2})$/, '');
    const catMap = {
      '射出': '射出', '射出A': '射出A', '射出C': '射出C',
      '射出D(組件)': '射出D(組件)', '射出D': '射出D',
      'Tubing': 'Tubing',
      '裝配A': '裝配A', '裝配B': '裝配B', '裝配C': '裝配C'
    };
    return catMap[name] || name;
  }

  if (qc === 'QC10004-R02') {
    return null; // QC10004-R02 is processed separately via scanInjectionData and scanExtrusionData
  }

  return null;
}

export function extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC) {
  const dateInfo = findDateInSheet(ws, actualQC);
  if (dateInfo) {
    if (dateInfo.year === year + 1 && dateInfo.month === 1) return 12;
    if (dateInfo.year === year) return dateInfo.month;
  }

  const y = String(year);
  let n, mn;
  n = fileName.match(/(\d{4})[-_](\d{1,2})\.xlsx$/i);
  if (n) {
    const yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(y, 10)) {
      mn = parseInt(n[2], 10);
      if (mn >= 1 && mn <= 12) return mn;
    }
  }
  n = fileName.match(/(\d{4})(\d{2})\d{2}(?=[^\/\\]*\.xlsx)/i);
  if (n) {
    const yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(y, 10)) {
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
  if (json) {
    mn = findDateInSheetFallback(json);
    if (mn) return mn;
  }
  // Strategy 9: Letter suffix A-L in filename (e.g., 裝配C-2021A.xlsx → A=1月)
  n = fileName.match(/[-_](\d{4})([A-L])\.xlsx$/i);
  if (n) {
    const yr = parseInt(n[1], 10);
    if (yr === year || yr === parseInt(String(year).slice(-2), 10)) {
      const letter = n[2].toUpperCase();
      if (LETTER_MONTH[letter]) return LETTER_MONTH[letter];
    }
  }
  return null;
}

export const runETLInBrowser = async (filesList, year, onProgress) => {
  const counts = {};
  const activeFiles = Array.from(filesList).filter(f => 
    f.name.endsWith('.xlsx') && !f.name.startsWith('~$')
  );

  let processedCount = 0;
  
  // QIP Injection and Extrusion variables
  const qipInjFiles = [];
  const qipExtFiles = [];
  const generalFiles = [];

  activeFiles.forEach(file => {
    const normalizedPath = file.webkitRelativePath.replace(/\\/g, '/');
    const pathLower = normalizedPath.toLowerCase();
    
    const isExtrusion = pathLower.includes('押出檢驗-' + year);
    const isInjection = pathLower.includes('射出檢驗-' + year);
    
    if (isInjection) {
      qipInjFiles.push(file);
    } else if (isExtrusion) {
      qipExtFiles.push(file);
    } else {
      generalFiles.push(file);
    }
  });

  const updateProgress = (filename) => {
    processedCount++;
    if (onProgress) {
      onProgress(processedCount, activeFiles.length, filename);
    }
  };

  // 1. Process General files
  for (let file of generalFiles) {
    updateProgress(file.name);
    
    // Determine initialQC and relPath
    const pathParts = file.webkitRelativePath.split('/');
    // Extract qcFolder and relPath
    let qcFolder = "";
    let initialQC = null;
    let folderIdx = -1;

    for (let i = 0; i < pathParts.length; i++) {
      const qc = detectQCFromFolder(pathParts[i]);
      if (qc) {
        initialQC = qc;
        qcFolder = pathParts[i];
        folderIdx = i;
        break;
      }
    }

    if (!initialQC) continue;

    const relPath = pathParts.slice(folderIdx + 1, pathParts.length - 1).join('/');
    const fileName = file.name;

    // Skip files with "空白" in the name (only if it is a blank template file with isolated '空白')
    const isBlankFile = (() => {
      let idx = fileName.indexOf('空白');
      if (idx === -1) return false;
      
      const letterOrCjk = /[a-zA-Z\u4e00-\u9fa5]/;
      
      while (idx !== -1) {
        const prevChar = idx > 0 ? fileName[idx - 1] : '';
        const nextChar = idx + 2 < fileName.length ? fileName[idx + 2] : '';
        
        const isAdjacent = letterOrCjk.test(prevChar) || letterOrCjk.test(nextChar);
        if (isAdjacent) {
          return false;
        }
        idx = fileName.indexOf('空白', idx + 1);
      }
      return true;
    })();
    if (isBlankFile) continue;

    try {
      const data = await file.arrayBuffer();
      
      // Pre-read sheet names to skip parsing non-target sheets
      const wbHeader = XLSX.read(data, { type: 'array', bookSheets: true });
      const targetSheets = wbHeader.SheetNames.filter(sheetName => {
        if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別') return false;
        if (sheetName.indexOf('Sheet') >= 0) return false; // skip any sheet with "Sheet" in name
        if (/^QC[-_]?\d+/i.test(sheetName.trim())) return false; // skip template sheets named QC-xxx
        if (sheetName.trim().indexOf('出貨') === 0) return false;
        return true;
      });

      if (targetSheets.length === 0) continue;

      const wb = XLSX.read(data, { 
        type: 'array',
        sheets: targetSheets,
        cellFormula: false,
        cellHTML: false,
        cellStyles: false,
        cellDates: true
      });
      const seenQC7R1BaseNames = new Set();

      targetSheets.forEach(sheetName => {
        const ws = wb.Sheets[sheetName];
        if (!ws) return;
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        let actualQC = determineQCFromSheet(json, initialQC, relPath);
        let subCat = null;
        let month = null;

        // Overrides and blank guards
        // Keyword-driven: matches filename OR folder path (e.g., 半成品品檢表2023(限組件用)/xxx.xlsx)
        const isSemiFinishedTable = /半成品品檢表/i.test(fileName) || /半成品品檢表/i.test(relPath || '');
        if (isSemiFinishedTable) {
          actualQC = 'QC10006-R02';
          subCat = '裝配C';
          // Flexible month extraction for 半成品品檢表:
          // 1. Sheet name letter pattern: PJW25D13 → D=4月 (covers 2015-2027+)
          const sheetMatch = sheetName.match(/(\d{2})([A-L])/i);
          if (sheetMatch) {
            const yr = parseInt(sheetMatch[1], 10);
            if (yr >= 15 && yr <= 99) {
              month = LETTER_MONTH[sheetMatch[2].toUpperCase()];
            }
          }
          // 2. Filename pattern: 半成品品檢表2023-01.xlsx → 1月
          if (!month) {
            const fnMonthMatch = fileName.match(/半成品品檢表\d{4}[-_](\d{1,2})/i);
            if (fnMonthMatch) {
              month = parseInt(fnMonthMatch[1], 10);
            }
          }
          // 3. Sheet name YYMMDD pattern: 230115 → 1月 (matches any 2-digit year prefix)
          if (!month) {
            const sheetYymmdd = sheetName.match(/(\d{2})(\d{2})(\d{2})/);
            if (sheetYymmdd) {
              const m = parseInt(sheetYymmdd[2], 10);
              if (m >= 1 && m <= 12) month = m;
            }
          }
        } else if (initialQC === 'QC10007-R03') {
          actualQC = 'QC10007-R03';
          
          const tempSub = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
          const isAssemblyParts = tempSub === '裝配A' || tempSub === '裝配B' || tempSub === '裝配C' || tempSub === '射出D(組件)';
          
          if (isAssemblyParts) {
            const letterMatch = fileName.match(/[-_]?([A-L])\.xlsx$/i);
            if (letterMatch) {
              month = LETTER_MONTH[letterMatch[1].toUpperCase()];
            }
          } else if (tempSub === 'Tubing' || tempSub === '射出' || tempSub === '射出A' || tempSub === '射出C') {
            // Tubing and 射出 subcategories month extraction directly from folder path
            if (relPath) {
              const folderMatch = relPath.match(/[-_](\d{1,2})$/) || relPath.match(/[-_](\d{1,2})[\\/]/);
              if (folderMatch) {
                const mVal = parseInt(folderMatch[1], 10);
                if (mVal >= 1 && mVal <= 12) month = mVal;
              } else {
                const monthChineseMatch = relPath.match(/(\d{1,2})月/);
                if (monthChineseMatch) {
                  const mVal = parseInt(monthChineseMatch[1], 10);
                  if (mVal >= 1 && mVal <= 12) month = mVal;
                }
              }
            }
          } else {
            // Only apply letter suffix matching for non-Tubing subcategories
            if (relPath && relPath.indexOf('Tubing') < 0) {
              const letterMatch = fileName.match(/[-_]?([A-L])\.xlsx$/i);
              if (letterMatch) {
                const derivedMonth = LETTER_MONTH[letterMatch[1].toUpperCase()];
                // Verify the derived month matches actual file date content
                // If not, fall through to extractRawMonth() for proper detection
                const fileDate = findDateInSheet(ws, actualQC);
                if (fileDate && fileDate.month === derivedMonth) {
                  month = derivedMonth;
                }
                // If no date in cell or mismatch, let extractRawMonth() handle it
              }
            }
          }
          if (relPath && relPath.indexOf('射出D') >= 0 && relPath.indexOf('射出D(組件)') < 0) {
            actualQC = 'QC10002-R02';
          }

          // Blank template guard for QC10007-R03 (零組件入庫品檢):
          const isExemptedInjection = tempSub === '射出' || tempSub === '射出A' || tempSub === '射出C' || tempSub === '射出D(組件)';
          if (actualQC === 'QC10007-R03' && !isExemptedInjection && json && json.length > 3) {
            const _lotRow = json[3];
            const _lotVal = (_lotRow && _lotRow.length > 6) ? _lotRow[6] : '';
            const _lotIsBlank = (_lotVal === '' || _lotVal === null || _lotVal === undefined ||
                              _lotVal === 0 || String(_lotVal).trim() === '' || String(_lotVal).trim() === '0');
            if (_lotIsBlank) {
              return; // Skip blank template worksheet
            }
          }
        }

        // If not set by override, compute them now
        if (subCat === null) {
          subCat = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
        }
        if (month === null) {
          const isAssemblyParts = actualQC === 'QC10007-R03' && (subCat === '裝配A' || subCat === '裝配B' || subCat === '裝配C' || subCat === '射出D(組件)');
          const isBypassedParts = actualQC === 'QC10007-R03' && (subCat === 'Tubing' || subCat === '射出' || subCat === '射出A' || subCat === '射出C');
          if (isAssemblyParts || isBypassedParts) {
            // Do not call extractRawMonth, keep overridden month from filename suffix or folder path
          } else {
            month = extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC);
          }
        }

        if (actualQC === 'QC10007-R01') {
          const baseName = sheetName.replace(/\s*\([^)]+\)\s*$/, '').trim();
          if (seenQC7R1BaseNames.has(baseName)) return;
          seenQC7R1BaseNames.add(baseName);
        }

        if (!actualQC || !subCat || !month || month < 1 || month > 12) return;

        if (!counts[actualQC]) counts[actualQC] = {};
        if (!counts[actualQC][subCat]) counts[actualQC][subCat] = {};
        counts[actualQC][subCat][month] = (counts[actualQC][subCat][month] || 0) + 1;
      });
    } catch (e) {
      console.error(`Error reading ${file.name}:`, e);
    }
  }

  // 2. Process QIP Injection data
  const injSetupCounts = {};
  const injPatrolCounts = {};
  MONTHS.forEach(m => { injSetupCounts[m] = 0; injPatrolCounts[m] = 0; });

  for (let file of qipInjFiles) {
    updateProgress(file.name);
    const normalizedPath = file.webkitRelativePath.replace(/\\/g, '/');
    const parts = normalizedPath.split('/');
    if (parts.length < 2) continue;
    
    const parentDir = parts[parts.length - 2];
    const mMatch = parentDir.match(/-(\d{2})$/);
    if (!mMatch) continue;
    const month = parseInt(mMatch[1], 10);
    if (month < 1 || month > 12) continue;
    
    // Every file in the folder is counted as Setup
    injSetupCounts[month]++;
    
    // If it is in the patrol folder, also parse sheets for patrol counts
    const pathLower = normalizedPath.toLowerCase();
    const isPatrol = pathLower.includes('qip-' + year + '(1~10)') || pathLower.includes('qip-' + year + '(1-10)');
    if (isPatrol) {
      try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array', bookSheets: true });
        const uniqueBaseInFile = {};
        wb.SheetNames.forEach(sheetName => {
          const normalizedSheetName = normalizeScientificNotation(sheetName);
          const baseName = normalizedSheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
          if (/^\d{6}[a-zA-Z]?$/.test(baseName)) {
            uniqueBaseInFile[baseName] = true;
          }
        });
        injPatrolCounts[month] += Object.keys(uniqueBaseInFile).length;
      } catch (e) {
        console.error(`Error reading QIP Patrol ${file.name}:`, e);
      }
    }
  }
  if (!counts['QC10004-R02']) counts['QC10004-R02'] = {};
  counts['QC10004-R02']['QIP-Setup'] = injSetupCounts;
  counts['QC10004-R02']['QIP-Patrol'] = injPatrolCounts;

  // 3. Process QIP Extrusion data
  const extSetupCounts = {};
  const extPatrolCounts = {};
  MONTHS.forEach(m => { extSetupCounts[m] = 0; extPatrolCounts[m] = 0; });

  for (let file of qipExtFiles) {
    updateProgress(file.name);
    const normalizedPath = file.webkitRelativePath.replace(/\\/g, '/');
    const parts = normalizedPath.split('/');
    if (parts.length < 2) continue;
    
    const parentDir = parts[parts.length - 2];
    const mMatch = parentDir.match(/-(\d{2})$/);
    if (!mMatch) continue;
    const month = parseInt(mMatch[1], 10);
    if (month < 1 || month > 12) continue;

    const isDateCodeFile = /\d{6}[a-zA-Z]?/i.test(file.name);
    if (!isDateCodeFile) continue;

    extSetupCounts[month]++;

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array', bookSheets: true });
      const uniqueBaseInFile = {};
      wb.SheetNames.forEach(sheetName => {
        const normalizedSheetName = normalizeScientificNotation(sheetName);
        if (normalizedSheetName === 'DATE' || normalizedSheetName === '空白' || normalizedSheetName === '範例' || normalizedSheetName === '客戶別' || normalizedSheetName.indexOf('Sheet1') === 0) return;
        if (normalizedSheetName.indexOf('.K(') >= 0 || normalizedSheetName.indexOf('範例樣本') >= 0 || /^QC[-_]?\d+/i.test(normalizedSheetName.trim())) return;
        if (/^(工作表|Sheet)\d+/i.test(normalizedSheetName.trim())) return;
        if (/^(工作表|Sheet)/i.test(normalizedSheetName.trim())) return; // skip any sheet named "工作表" or "Sheet"
        
        const snLower = normalizedSheetName.toLowerCase();
        const isSetup = (snLower.indexOf('setup') >= 0 || snLower.indexOf('set up') >= 0 || snLower.indexOf('set-up') >= 0 || snLower === 'setup');
        
        if (!isSetup) {
          const baseName = normalizedSheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
          uniqueBaseInFile[baseName] = true;
        }
      });
      extPatrolCounts[month] += Object.keys(uniqueBaseInFile).length;
    } catch (e) {
      console.error(`Error reading Extrusion ${file.name}:`, e);
    }
  }
  counts['QC10004-R02']['押出-Setup'] = extSetupCounts;
  counts['QC10004-R02']['押出-Patrol'] = extPatrolCounts;

  return counts;
};

export const exportSummaryExcelInBrowser = (counts, year) => {
  const wb = XLSX.utils.book_new();

  const monthArray = (data) => MONTHS.map(m => (data && data[m]) || 0);
  const totalArray = (data) => monthArray(data).reduce((a, b) => a + b, 0);

  const addCategorySheet = (sheetName, columns, subCatToCol, qcCode, subCategoryFilter, titleRowText) => {
    const colData = columns.map(() => ({}));
    const qcCounts = counts[qcCode];
    if (qcCounts) {
      for (let subCat in qcCounts) {
        if (subCategoryFilter && !subCategoryFilter(subCat)) continue;
        const colIdx = subCatToCol(subCat);
        if (colIdx === null || colIdx === undefined) continue;

        const monthly = qcCounts[subCat];
        for (let m = 1; m <= 12; m++) {
          if (monthly[m]) {
            colData[colIdx][m] = (colData[colIdx][m] || 0) + monthly[m];
          }
        }
      }
    }

    const rows = [
      [titleRowText || sheetName],
      ['月份', ...columns.map(c => c.label)]
    ];
    if (qcCode !== 'QC10006-R01') {
      rows[1].push('小計');
    }

    MONTHS.forEach(m => {
      const row = [`${m}月`];
      let total = 0;
      columns.forEach((c, ci) => {
        const v = colData[ci][m] || 0;
        row.push(v);
        total += v;
      });
      if (qcCode !== 'QC10006-R01') {
        row.push(total);
      }
      rows.push(row);
    });

    const totalRow = ['小計'];
    let grandTotal = 0;
    columns.forEach((c, ci) => {
      const t = totalArray(colData[ci]);
      totalRow.push(t);
      grandTotal += t;
    });
    if (qcCode !== 'QC10006-R01') {
      totalRow.push(grandTotal);
    }
    rows.push(totalRow);

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
  };

  // 1. 原物料品檢 (QC10002-R02)
  const rawCols = [
    {key:'原料', label:'原料'}, {key:'物料-B膠', label:'物料-B膠'}, {key:'物料-收縮膜', label:'物料-收縮膜'},
    {key:'物料-色粉', label:'物料-色粉'}, {key:'物料-空白包裝袋', label:'物料-空白包裝袋'}, {key:'物料-空白感壓紙', label:'物料-空白感壓紙'},
    {key:'物料-塑膠袋', label:'物料-塑膠袋'}, {key:'物料-塑膠袋40X50', label:'物料-塑膠袋40X50'}, {key:'物料-紙箱', label:'物料-紙箱'},
    {key:'物料-過濾網連蓋', label:'物料-過濾網連蓋'}, {key:'物料-標籤', label:'物料-標籤'}, {key:'射出D', label:'射出D'}
  ];
  const rawSubCatMap = (subCat) => {
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
  };
  addCategorySheet(
    '原物料品檢(QC10002-R02)',
    rawCols,
    rawSubCatMap,
    'QC10002-R02',
    null,
    '原物料/配件進料品檢'
  );

  // 2. QIP (QC10004-R02)
  const qipCols = [
    {key:'QIP-Setup', label:'Setup(射出)'}, {key:'QIP-Patrol', label:'巡檢(射出)'},
    {key:'押出-Setup', label:'Setup(押出)'}, {key:'押出-Patrol', label:'巡檢(押出)'}
  ];
  const qipSubCatMap = {
    'QIP-Setup': 0, 'QIP-Patrol': 1, '押出-Setup': 2, '押出-Patrol': 3
  };
  addCategorySheet(
    'QIP(QC10004-R02)',
    qipCols,
    (sub) => qipSubCatMap[sub],
    'QC10004-R02',
    null,
    'QIP品檢'
  );

  // 3. 裝配對樣巡檢 (QC10006-R01)
  addCategorySheet(
    '裝配對樣巡檢(QC10006-R01)',
    [{key:'裝配巡檢', label:'裝配巡檢'}],
    () => 0,
    'QC10006-R01',
    null,
    '裝配對樣巡檢'
  );

  // 4. 半成品品檢 (QC10006-R02)
  const semiCols = [
    {key:'裝配C', label:'裝配C'}, {key:'BD', label:'BD'}, {key:'Biometrix', label:'Biometrix'},
    {key:'MPS', label:'MPS'}, {key:'Vivus', label:'Vivus'}
  ];
  const semiSubCatMap = { 'BD': 1, 'Biometrix': 2, 'MPS': 3, 'Vivus': 4 };
  addCategorySheet(
    '半成品品檢(QC10006-R02)',
    semiCols,
    (sub) => {
      const idx = semiSubCatMap[sub];
      if (idx === undefined && sub.indexOf('裝配C') >= 0) return 0;
      return idx;
    },
    'QC10006-R02',
    null,
    '裝配半成品品檢'
  );

  // 5. 完成品品檢 (QC10007-R01 R02)
  const finCols = [
    {key:'Biometrix', label:'Biometrix'}, {key:'MarMed', label:'MarMed'},
    {key:'Saxon', label:'Saxon'}, {key:'Vivus', label:'Vivus'}
  ];
  const finSubCatMap = { 'Biometrix': 0, 'MarMed': 1, 'Saxon': 2, 'Vivus': 3 };
  addCategorySheet(
    '完成品品檢(QC10007-R01 R02)',
    finCols,
    (sub) => finSubCatMap[sub],
    'QC10007-R01',
    null,
    '裝配完成品品檢'
  );

  // 6. 零組件入庫品檢 (QC10007-R03)
  const partsCols = [
    {key:'Tubing', label:'Tubing'}, {key:'射出', label:'射出(廠內)'}, {key:'射出A', label:'射出A'},
    {key:'射出C', label:'射出C'}, {key:'射出D(組件)', label:'射出D(組件)'}, {key:'裝配A', label:'裝配A'},
    {key:'裝配B', label:'裝配B'}, {key:'裝配C', label:'裝配C'}
  ];
  const partsSubCatMap = {
    'Tubing': 0, '射出': 1, '射出A': 2, '射出C': 3, '射出D(組件)': 4,
    '裝配A': 5, '裝配B': 6, '裝配C': 7
  };
  addCategorySheet(
    '零組件入庫品檢(QC10007-R03)',
    partsCols,
    (sub) => partsSubCatMap[sub],
    'QC10007-R03',
    null,
    '零組件入庫檢'
  );

  // 7. 出貨檢驗 (QC10008-R02)
  const shipCols = [
    {key:'ICU', label:'ICU'}, {key:'其他', label:'其他'}
  ];
  const shipSubCatMap = { 'ICU': 0, '其他': 1 };
  addCategorySheet(
    '出貨檢驗(QC10008-R02)',
    shipCols,
    (sub) => shipSubCatMap[sub],
    'QC10008-R02',
    null,
    '出貨檢驗'
  );

  XLSX.writeFile(wb, `${year}品檢報表統計.xlsx`);
};
