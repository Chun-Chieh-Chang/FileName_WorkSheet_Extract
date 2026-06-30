import * as XLSX from 'xlsx';

const MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
const LETTER_MONTH = { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, I: 9, J: 10, K: 11, L: 12 };

const FOLDER_QC_MAP = {
  '原物料品檢': 'QC10002-R02',
  '裝配巡檢': 'QC10006-R01',
  '裝配檢驗': 'QC10006-R02',
  '完成品品檢': 'QC10007-R01',
  '零組件入庫': 'QC10007-R03',
  '出貨檢驗': 'QC10008-R02'
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

function detectQCFromFolder(dirname) {
  const m = dirname.match(/QC\d{5}-R\d{2}/i);
  if (m) return m[0].toUpperCase();
  const keys = Object.keys(FOLDER_QC_MAP);
  for (let i = 0; i < keys.length; i++) {
    if (dirname.indexOf(keys[i]) >= 0) return FOLDER_QC_MAP[keys[i]];
  }
  return null;
}

function parseDateFromValue(val, formatted) {
  if (!val) return null;
  
  if (val instanceof Date) {
    return { year: val.getFullYear(), month: val.getMonth() + 1 };
  }
  
  if (typeof val === 'number') {
    if (val > 20000 && val < 60000) {
      const d = new Date((val - 25569) * 86400 * 1000);
      return { year: d.getFullYear(), month: d.getMonth() + 1 };
    }
  }

  const s = String(formatted || val || '').trim();
  const mYMD = s.match(/(20\d{2})[-/年](\d{1,2})[-/月](\d{1,2})/);
  if (mYMD) {
    return { year: parseInt(mYMD[1], 10), month: parseInt(mYMD[2], 10) };
  }

  const mYY = s.match(/(2\d)[-/年](\d{1,2})[-/月](\d{1,2})/);
  if (mYY) {
    return { year: 2000 + parseInt(mYY[1], 10), month: parseInt(mYY[2], 10) };
  }

  return null;
}

function findDateInSheet(ws, qc) {
  if (!ws) return null;

  const getCellValAndFormatted = (addr) => {
    const cell = ws[addr];
    if (!cell) return { val: null, formatted: null };
    return { val: cell.v, formatted: cell.w };
  };

  let cellInfo, dateInfo;
  switch (qc) {
    case 'QC10002-R02':
      cellInfo = getCellValAndFormatted('J3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      cellInfo = getCellValAndFormatted('K3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      break;

    case 'QC10006-R02':
      cellInfo = getCellValAndFormatted('J3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      break;

    case 'QC10007-R01':
      cellInfo = getCellValAndFormatted('J3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      break;

    case 'QC10007-R03':
      cellInfo = getCellValAndFormatted('O4');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      break;

    case 'QC10008-R02':
      cellInfo = getCellValAndFormatted('I3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      cellInfo = getCellValAndFormatted('H3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      cellInfo = getCellValAndFormatted('G3');
      dateInfo = parseDateFromValue(cellInfo.val, cellInfo.formatted);
      if (dateInfo) return dateInfo;
      break;
  }

  // Scan first 15 rows for any date
  const cellKeys = Object.keys(ws).filter(k => !k.startsWith('!'));
  for (let key of cellKeys) {
    const parseRow = parseInt(key.replace(/^[A-Z]+/i, ''), 10);
    if (parseRow <= 15) {
      const cell = ws[key];
      if (cell && cell.v) {
        const info = parseDateFromValue(cell.v, cell.w);
        if (info && info.year >= 2024 && info.year <= 2028) {
          return info;
        }
      }
    }
  }

  return null;
}

function determineQCFromSheet(json, initialQC, relPath) {
  if (initialQC === 'QC10006-R01') return 'QC10006-R01';
  if (initialQC === 'QC10004-R02') return 'QC10004-R02';

  const scanLimit = Math.min(15, json.length);
  for (let ri = 0; ri < scanLimit; ri++) {
    const row = json[ri];
    if (!row) continue;
    for (let ci = 0; ci < row.length; ci++) {
      const val = String(row[ci] || '').trim();
      if (!val) continue;

      if (val.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (val.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (val.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (val.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (val.indexOf('QC10007-R01') >= 0 || val.indexOf('QC10007-R02') >= 0 || val.indexOf('QC10007') >= 0) return 'QC10007-R01';
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
    for (let ci = 0; ci < row.length; ci++) {
      const val = String(row[ci] || '').trim();
      if (!val) continue;
      if (val.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (val.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (val.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (val.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (val.indexOf('QC10007-R01') >= 0 || val.indexOf('QC10007-R02') >= 0 || val.indexOf('QC10007') >= 0) return 'QC10007-R01';
      if (val.indexOf('QC10008') >= 0) return 'QC10008-R02';
    }
  }

  return initialQC;
}

function getRawSubCategory(qc, relPath, fileName, sheetName, qcFolder) {
  if (qc === 'QC10002-R02') {
    if (relPath && relPath.indexOf('原料') >= 0) return '原料';
    const matMap = {
      'B膠': '物料-B膠', '塑膠袋': '物料-塑膠袋', '塑膠袋40': '物料-塑膠袋40X50',
      '收縮膜': '物料-收縮膜', '標籤': '物料-標籤', '空白包裝袋': '物料-空白包裝袋',
      '空白感壓紙': '物料-空白感壓紙', '紙箱': '物料-紙箱', '色粉': '物料-色粉',
      '過濾網': '物料-過濾網連蓋'
    };
    const keys = Object.keys(matMap);
    for (let i = 0; i < keys.length; i++) {
      if (relPath && relPath.indexOf(keys[i]) >= 0) return matMap[keys[i]];
    }
    if (relPath && relPath.indexOf('射出D') >= 0) return '射出D';
    return '原料';
  }

  if (qc === 'QC10008-R02') {
    if (sheetName.indexOf('ICU') >= 0) return 'ICU';
    return '其他';
  }

  if (qc === 'QC10006-R02' || qc === 'QC10007-R01') {
    const name = relPath.split('/')[0].replace(/[-_](20\d{2})$/, '');
    return name;
  }

  if (qc === 'QC10006-R01') return '裝配巡檢';

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

  return null;
}

function extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC) {
  if (relPath && relPath.indexOf('Tubing') >= 0) {
    const folderMatch = relPath.match(/Tubing-\d{4}-(\d{1,2})/);
    if (folderMatch) {
      return parseInt(folderMatch[1], 10);
    }
  }

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
    const pathLower = file.webkitRelativePath.toLowerCase();
    if (pathLower.includes('射出檢驗') && (pathLower.includes('qip') || pathLower.includes('setup') || pathLower.includes('patrol'))) {
      qipInjFiles.push(file);
    } else if (pathLower.includes('押出檢驗')) {
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

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const seenQC7R1BaseNames = new Set();

      wb.SheetNames.forEach(sheetName => {
        if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別') return;
        if (sheetName.trim().indexOf('出貨') === 0) return;

        const ws = wb.Sheets[sheetName];
        if (!ws) return;
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        let actualQC = determineQCFromSheet(json, initialQC, relPath);
        let subCat = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
        let month = extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC);

        // Special overrides
        if (initialQC === 'QC10007-R03') {
          actualQC = 'QC10007-R03';
          const letterMatch = fileName.match(/(?:202[56]|2[56])[-_]?([A-L])\.xlsx$/i);
          if (letterMatch) {
            month = LETTER_MONTH[letterMatch[1].toUpperCase()];
          }
          if (relPath && relPath.indexOf('射出D') >= 0 && relPath.indexOf('射出D(組件)') < 0) {
            actualQC = 'QC10002-R02';
          }
        }

        const isSemiFinishedTable = /半成品品檢表-20\d{2}\.xlsx$/i.test(fileName);
        if (isSemiFinishedTable) {
          actualQC = 'QC10006-R02';
          subCat = '裝配C';
          const sheetMatch = sheetName.match(/2[56]([A-L])/i);
          if (sheetMatch) {
            month = LETTER_MONTH[sheetMatch[1].toUpperCase()];
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
    const pathLower = file.webkitRelativePath.toLowerCase();
    
    // Extract month from folder name ending in -MM (like QIP-2025(1-10)-02)
    const mMatch = file.webkitRelativePath.match(/-(\d{2})[/\\]/);
    if (!mMatch) continue;
    const month = parseInt(mMatch[1], 10);
    if (month < 1 || month > 12) continue;

    // Is it under QIP Patrol folder?
    const isPatrol = pathLower.includes('qip-') && (pathLower.includes('(1~10)') || pathLower.includes('(1-10)'));
    
    if (!isPatrol) {
      // Setup count
      injSetupCounts[month]++;
    } else {
      // Patrol count
      try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const uniqueBaseInFile = {};
        wb.SheetNames.forEach(sheetName => {
          const baseName = sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
          if (/^\d{6}[a-zA-Z]?$/.test(baseName)) {
            uniqueBaseInFile[baseName] = true;
          }
        });
        injPatrolCounts[month] += Object.keys(uniqueBaseInFile).length;
      } catch (e) {
        console.error(e);
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
    
    const mMatch = file.webkitRelativePath.match(/-(\d{2})[/\\]/);
    if (!mMatch) continue;
    const month = parseInt(mMatch[1], 10);
    if (month < 1 || month > 12) continue;

    const isDateCodeFile = /\d{6}[a-zA-Z]?/i.test(file.name);
    if (!isDateCodeFile) continue;

    extSetupCounts[month]++;

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const uniqueBaseInFile = {};
      wb.SheetNames.forEach(sheetName => {
        if (sheetName.toUpperCase().includes('SETUP')) return;
        const baseName = sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
        if (/^\d{6}[a-zA-Z]?$/.test(baseName)) {
          uniqueBaseInFile[baseName] = true;
        }
      });
      extPatrolCounts[month] += Object.keys(uniqueBaseInFile).length;
    } catch (e) {
      console.error(e);
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
  const rawSubCatMap = {
    '原料': 0, '物料-B膠': 1, '物料-收縮膜': 2, '物料-色粉': 3, '物料-空白包裝袋': 4,
    '物料-空白感壓紙': 5, '物料-塑膠袋': 6, '物料-塑膠袋40X50': 7, '物料-紙箱': 8,
    '物料-過濾網連蓋': 9, '物料-標籤': 10, '射出D': 11
  };
  addCategorySheet(
    '原物料品檢(QC10002-R02)',
    rawCols,
    (sub) => rawSubCatMap[sub],
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

  // 8. 品檢地圖
  const mapRows = [
    ['品檢地圖'],
    ['表單代碼', '工作表名稱'],
    ['QC10002-R02', '原物料品檢(QC10002-R02)'],
    ['QC10004-R02', 'QIP(QC10004-R02)'],
    ['QC10006-R01', '裝配對樣巡檢(QC10006-R01)'],
    ['QC10006-R02', '半成品品檢(QC10006-R02)'],
    ['QC10007-R01 R02', '完成品品檢(QC10007-R01 R02)'],
    ['QC10007-R03', '零組件入庫品檢(QC10007-R03)'],
    ['QC10008-R02', '出貨檢驗(QC10008-R02)'],
    ['彙總表', '彙總表']
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mapRows), '品檢地圖');

  // 9. 彙總表 (McKinsey 3x3 Grid Summary layout)
  const buildSummarySheet = () => {
    const summaryRows = [
      ['品檢報表年度彙總表'],
      ['']
    ];

    const getQCStats = (qcCode, subFilter) => {
      const stat = {};
      MONTHS.forEach(m => { stat[m] = 0; });
      const qcCounts = counts[qcCode];
      if (qcCounts) {
        for (let sub in qcCounts) {
          if (subFilter && !subFilter(sub)) continue;
          const monthly = qcCounts[sub];
          MONTHS.forEach(m => {
            stat[m] += (monthly[m] || 0);
          });
        }
      }
      return stat;
    };

    const s1 = getQCStats('QC10002-R02', sub => sub !== '射出D');
    const s2 = getQCStats('QC10004-R02', sub => sub.indexOf('QIP') >= 0);
    const s3 = getQCStats('QC10004-R02', sub => sub.indexOf('押出') >= 0);
    const s4 = getQCStats('QC10006-R01');
    const s5 = getQCStats('QC10006-R02');
    const s6 = getQCStats('QC10007-R01');
    const s7 = getQCStats('QC10007-R03');
    const s8 = getQCStats('QC10008-R02');

    // Headers row
    summaryRows.push([
      '月份', '原物料品檢(QC10002-R02)', '', 
      'QIP-射出(QC10004-R02)', '', 
      'QIP-押出(QC10004-R02)', '', 
      '裝配對樣巡檢(QC10006-R01)', '', 
      '半成品品檢(QC10006-R02)', '', 
      '完成品品檢(QC10007-R01)', '', 
      '零組件入庫品檢(QC10007-R03)', '', 
      '出貨檢驗(QC10008-R02)'
    ]);

    MONTHS.forEach(m => {
      summaryRows.push([
        `${m}月`,
        s1[m], '',
        s2[m], '',
        s3[m], '',
        s4[m], '',
        s5[m], '',
        s6[m], '',
        s7[m], '',
        s8[m]
      ]);
    });

    const totalRow = [
      '小計',
      totalArray(s1), '',
      totalArray(s2), '',
      totalArray(s3), '',
      totalArray(s4), '',
      totalArray(s5), '',
      totalArray(s6), '',
      totalArray(s7), '',
      totalArray(s8)
    ];
    summaryRows.push(totalRow);

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), '彙總表');
  };
  buildSummarySheet();

  XLSX.writeFile(wb, `${year}品檢報表統計.xlsx`);
};
