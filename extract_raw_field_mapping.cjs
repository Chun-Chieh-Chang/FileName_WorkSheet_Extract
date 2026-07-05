/**
 * 原始資料欄位映射提取工具
 * 
 * 用途：掃描原始資料夾，提取每個檔案的品檢報表統計對應關係
 * 輸出格式：品管標籤編號(工作表名稱), 欄位名稱(項目欄位), 資料路徑
 * 
 * 邏輯來源：完全對齊 src/utils/browserETL.js 的 ETL 邏輯
 * 
 * 使用方式：
 *   node extract_raw_field_mapping.cjs <原始資料路徑> [品檢報表統計檔案路徑]
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ============================================================
// 常數定義（對齊 browserETL.js）
// ============================================================

const MONTHS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
const LETTER_MONTH = { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, I: 9, J: 10, K: 11, L: 12 };

const FOLDER_QC_MAP = {
  '半成品品檢表': 'QC10006-R02',
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

const QC_LABELS = {
  'QC10002-R02': { name: '原物料/配件進料品檢表', sheetName: '原物料品檢(QC10002-R02)' },
  'QC10004-R02': { name: 'QIP尺寸檢驗紀錄表', sheetName: 'QIP(QC10004-R02)' },
  'QC10006-R01': { name: '裝配對樣巡檢記錄表', sheetName: '裝配對樣巡檢(QC10006-R01)' },
  'QC10006-R02': { name: '半成品檢驗記錄表', sheetName: '半成品品檢(QC10006-R02)' },
  'QC10007-R01': { name: '完成品裝配品檢紀錄表', sheetName: '完成品品檢(QC10007-R01 R02)' },
  'QC10007-R03': { name: '零組件入庫品檢表', sheetName: '零組件入庫品檢(QC10007-R03)' },
  'QC10008-R02': { name: '出貨檢驗報告', sheetName: '出貨檢驗(QC10008-R02)' }
};


// ============================================================
// QC 碼判斷（對齊 browserETL.js）
// ============================================================

function detectQCFromFolder(dirname) {
  const m = dirname.match(/QC\d{5}-R\d{2}/i);
  if (m) return m[0].toUpperCase();
  const keys = Object.keys(FOLDER_QC_MAP);
  for (let i = 0; i < keys.length; i++) {
    if (dirname.indexOf(keys[i]) >= 0) return FOLDER_QC_MAP[keys[i]];
  }
  return null;
}

function determineQCFromSheet(json, initialQC) {
  if (initialQC === 'QC10006-R01') return 'QC10006-R01';
  if (initialQC === 'QC10004-R02') return 'QC10004-R02';

  const scanLimit = Math.min(15, json.length);
  for (let ri = 0; ri < scanLimit; ri++) {
    const row = json[ri];
    if (!row) continue;
    
    const colA = String(row[0] || '').trim();
    if (colA) {
      if (colA.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (colA.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (colA.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (colA.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (colA.indexOf('QC10007-R01') >= 0 || colA.indexOf('QC10007-R02') >= 0) return 'QC10007-R01';
      if (colA.indexOf('QC10008') >= 0) return 'QC10008-R02';
      
      if (ri === 0) {
        const titleKeys = Object.keys(FORM_TITLE_MAP);
        for (let k = 0; k < titleKeys.length; k++) {
          if (colA.indexOf(titleKeys[k]) >= 0) return FORM_TITLE_MAP[titleKeys[k]];
        }
      }
    }
    
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

  // Scan footer (last 30 rows)
  const footerStart = Math.max(0, json.length - 30);
  for (let ri = footerStart; ri < json.length; ri++) {
    const row = json[ri];
    if (!row) continue;
    const colA = String(row[0] || '').trim();
    if (colA) {
      if (colA.indexOf('QC10002-R02') >= 0) return 'QC10002-R02';
      if (colA.indexOf('QC10006-R01') >= 0) return 'QC10006-R01';
      if (colA.indexOf('QC10006-R02') >= 0) return 'QC10006-R02';
      if (colA.indexOf('QC10007-R03') >= 0) return 'QC10007-R03';
      if (colA.indexOf('QC10007-R01') >= 0 || colA.indexOf('QC10007-R02') >= 0) return 'QC10007-R01';
      if (colA.indexOf('QC10008') >= 0) return 'QC10008-R02';
    }
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


// ============================================================
// 子分類判斷（對齊 browserETL.js 的 getRawSubCategory）
// ============================================================

function getRawSubCategory(qc, relPath, fileName, sheetName, qcFolder) {
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
    return null; // QC10004-R02 is processed separately
  }

  return null;
}

// ============================================================
// 子分類映射到品檢報表統計欄位名稱
// ============================================================

function mapSubCategoryToFieldName(qc, subCat) {
  // QC10002-R02: 原物料品檢
  if (qc === 'QC10002-R02') {
    if (subCat === '原料') return '原料';
    if (subCat === '物料-B膠') return 'B膠';
    if (subCat === '物料-收縮膜') return '收縮膜';
    if (subCat === '物料-色粉') return '色粉';
    if (subCat === '物料-空白包裝袋') return '空白包裝袋';
    if (subCat === '物料-空白感壓紙') return '空白感壓紙';
    if (subCat === '物料-塑膠袋') return '塑膠袋';
    if (subCat === '物料-塑膠袋40*50' || subCat === '物料-塑膠袋40X50') return '塑膠袋40*50';
    if (subCat === '物料-紙箱') return '紙箱';
    if (subCat === '物料-過濾網連蓋') return '過濾網連蓋';
    if (subCat === '物料-標籤') return '標籤';
    if (subCat === '射出D') return '射出D';
    return subCat;
  }
  
  // QC10004-R02: QIP（注意：browserETL.js 中的欄位順序是 押出(Setup), 射出(Setup), 押出(巡檢), 射出(廠內)）
  if (qc === 'QC10004-R02') {
    if (subCat === 'QIP-Setup') return '押出(Setup)';
    if (subCat === 'QIP-Patrol') return '射出(Setup)';
    if (subCat === '押出-Setup') return '押出(巡檢)';
    if (subCat === '押出-Patrol') return '射出(廠內)';
    return subCat;
  }
  
  // QC10006-R01: 裝配巡檢
  if (qc === 'QC10006-R01') return '裝配';
  
  // QC10006-R02: 半成品品檢
  if (qc === 'QC10006-R02') {
    if (subCat.indexOf('裝配C') >= 0) return '裝配C';
    return subCat;
  }
  
  // QC10007-R01: 完成品品檢
  if (qc === 'QC10007-R01') return subCat;
  
  // QC10007-R03: 零組件入庫
  if (qc === 'QC10007-R03') {
    if (subCat === '射出') return '射出(廠內)';
    return subCat;
  }
  
  // QC10008-R02: 出貨檢驗
  if (qc === 'QC10008-R02') return subCat;
  
  return subCat;
}


// ============================================================
// 檔案掃描與映射提取
// ============================================================

function scanDirectory(dirPath) {
  const excelFiles = [];
  try {
    const items = fs.readdirSync(dirPath, { withFileTypes: true });
    items.forEach(item => {
      const fullPath = path.join(dirPath, item.name);
      if (item.isDirectory()) {
        excelFiles.push(...scanDirectory(fullPath));
      } else if (item.isFile() && (item.name.endsWith('.xlsx') || item.name.endsWith('.xls'))) {
        if (item.name.startsWith('~$')) return;
        if (item.name.indexOf('空白') >= 0) return;
        excelFiles.push(fullPath);
      }
    });
  } catch (error) {
    console.error(`  ⚠️  掃描目錄失敗: ${error.message}`);
  }
  return excelFiles;
}

function extractYearFromPath(rootDir) {
  const yearMatch = rootDir.match(/(20\d{2})/);
  if (yearMatch) return parseInt(yearMatch[1], 10);
  return new Date().getFullYear();
}

function extractAllMappings(rootDir) {
  if (!fs.existsSync(rootDir)) {
    console.error(`錯誤：目錄不存在 "${rootDir}"`);
    process.exit(1);
  }
  
  console.log(`📂 掃描目錄: ${rootDir}`);
  const year = extractYearFromPath(rootDir);
  console.log(`📊 年份: ${year}`);
  
  const excelFiles = scanDirectory(rootDir);
  console.log(`  找到 ${excelFiles.length} 個 Excel 檔案`);
  
  const allMappings = [];
  const stats = { totalFiles: excelFiles.length, mappedFiles: 0, unmappedFiles: 0, qcCounts: {} };
  
  // 分類檔案（對齊 browserETL.js 的分類邏輯）
  const qipInjFiles = [];
  const qipExtFiles = [];
  const generalFiles = [];
  
  excelFiles.forEach(filePath => {
    const normalizedPath = filePath.replace(/\\/g, '/');
    const pathLower = normalizedPath.toLowerCase();
    const isExtrusion = pathLower.includes('押出檢驗-' + year);
    const isInjection = pathLower.includes('射出檢驗-' + year);
    if (isInjection) qipInjFiles.push(filePath);
    else if (isExtrusion) qipExtFiles.push(filePath);
    else generalFiles.push(filePath);
  });
  
  console.log(`  分類: 一般檔案 ${generalFiles.length}, 射出檢驗 ${qipInjFiles.length}, 押出檢驗 ${qipExtFiles.length}`);
  
  // 處理一般檔案
  generalFiles.forEach(filePath => {
    const normalizedPath = filePath.replace(/\\/g, '/');
    const pathParts = normalizedPath.split('/');
    let qcFolder = "", initialQC = null, folderIdx = -1;
    
    for (let i = 0; i < pathParts.length; i++) {
      const qc = detectQCFromFolder(pathParts[i]);
      if (qc) { initialQC = qc; qcFolder = pathParts[i]; folderIdx = i; break; }
    }
    
    if (!initialQC) { stats.unmappedFiles++; return; }
    
    const relPath = pathParts.slice(folderIdx + 1, pathParts.length - 1).join('/');
    const fileName = pathParts[pathParts.length - 1];
    
    try {
      const wb = XLSX.readFile(filePath);
      const seenQC7R1BaseNames = new Set();
      
      wb.SheetNames.forEach(sheetName => {
        if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別') return;
        if (sheetName.indexOf('Sheet') >= 0) return;
        if (/^QC[-_]?\d+/i.test(sheetName.trim())) return;
        if (sheetName.trim().indexOf('出貨') === 0) return;
        
        const ws = wb.Sheets[sheetName];
        if (!ws) return;
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        
        let actualQC = determineQCFromSheet(json, initialQC);
        let subCat = null;
        
        // 特殊覆寫邏輯（對齊 browserETL.js）
        const isSemiFinishedTable = /半成品品檢表/i.test(fileName) || /半成品品檢表/i.test(relPath || '');
        if (isSemiFinishedTable) {
          actualQC = 'QC10006-R02';
          subCat = '裝配C';
        } else if (initialQC === 'QC10007-R03') {
          actualQC = 'QC10007-R03';
          if (relPath && relPath.indexOf('射出D') >= 0 && relPath.indexOf('射出D(組件)') < 0) {
            actualQC = 'QC10002-R02';
          }
        }
        
        if (subCat === null) {
          subCat = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
        }
        
        if (!subCat) return;
        
        if (actualQC === 'QC10007-R01') {
          const baseName = sheetName.replace(/\s*\([^)]+\)\s*$/, '').trim();
          if (seenQC7R1BaseNames.has(baseName)) return;
          seenQC7R1BaseNames.add(baseName);
        }
        
        const sheetName2 = QC_LABELS[actualQC]?.sheetName;
        if (!sheetName2) return;
        
        const fieldName = mapSubCategoryToFieldName(actualQC, subCat);
        allMappings.push({ sheetName: sheetName2, fieldName: fieldName, dataPath: filePath });
        if (!stats.qcCounts[actualQC]) stats.qcCounts[actualQC] = 0;
        stats.qcCounts[actualQC]++;
      });
      stats.mappedFiles++;
    } catch (e) {
      stats.unmappedFiles++;
    }
  });
  
  // 處理射出檢驗檔案 (QIP Injection)
  qipInjFiles.forEach(filePath => {
    const normalizedPath = filePath.replace(/\\/g, '/');
    const parts = normalizedPath.split('/');
    if (parts.length < 2) return;
    const parentDir = parts[parts.length - 2];
    const mMatch = parentDir.match(/-(\d{2})$/);
    if (!mMatch) return;
    
    allMappings.push({ sheetName: 'QIP(QC10004-R02)', fieldName: '射出(Setup)', dataPath: filePath });
    const pathLower = normalizedPath.toLowerCase();
    const isPatrol = pathLower.includes('qip-' + year + '(1~10)') || pathLower.includes('qip-' + year + '(1-10)');
    if (isPatrol) {
      allMappings.push({ sheetName: 'QIP(QC10004-R02)', fieldName: '射出(廠內)', dataPath: filePath });
    }
    if (!stats.qcCounts['QC10004-R02']) stats.qcCounts['QC10004-R02'] = 0;
    stats.qcCounts['QC10004-R02']++;
    stats.mappedFiles++;
  });
  
  // 處理押出檢驗檔案 (QIP Extrusion)
  qipExtFiles.forEach(filePath => {
    const normalizedPath = filePath.replace(/\\/g, '/');
    const parts = normalizedPath.split('/');
    if (parts.length < 2) return;
    const fileName = parts[parts.length - 1];
    const isDateCodeFile = /\d{6}[a-zA-Z]?/i.test(fileName);
    if (!isDateCodeFile) return;
    
    allMappings.push({ sheetName: 'QIP(QC10004-R02)', fieldName: '押出(Setup)', dataPath: filePath });
    allMappings.push({ sheetName: 'QIP(QC10004-R02)', fieldName: '押出(巡檢)', dataPath: filePath });
    if (!stats.qcCounts['QC10004-R02']) stats.qcCounts['QC10004-R02'] = 0;
    stats.qcCounts['QC10004-R02']++;
    stats.mappedFiles++;
  });
  
  return { mappings: allMappings, stats };
}

function toCSV(mappings) {
  const header = '品管標籤編號,欄位名稱,資料路徑';
  const rows = mappings.map(m => `"${m.sheetName}","${m.fieldName}","${m.dataPath}"`);
  return [header, ...rows].join('\n');
}


// ============================================================
// 主程式
// ============================================================

function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.log('用法: node extract_raw_field_mapping.cjs <原始資料路徑>');
    console.log('範例: node extract_raw_field_mapping.cjs "F:\\2025 報表"');
    process.exit(0);
  }
  
  const rootDir = args[0];
  console.log('='.repeat(80));
  console.log('📊 原始資料欄位映射提取工具');
  console.log('='.repeat(80));
  console.log('');
  
  const startTime = Date.now();
  const { mappings, stats } = extractAllMappings(rootDir);
  const extractTime = Date.now() - startTime;
  
  console.log('');
  console.log('='.repeat(80));
  console.log('📈 提取統計');
  console.log('='.repeat(80));
  console.log(`總檔案數: ${stats.totalFiles}`);
  console.log(`成功映射: ${stats.mappedFiles}`);
  console.log(`未映射: ${stats.unmappedFiles}`);
  console.log(`總映射筆數: ${mappings.length}`);
  console.log(`提取耗時: ${extractTime}ms`);
  
  console.log('');
  console.log('各 QC 標籤檔案數:');
  for (const qc in stats.qcCounts) {
    const label = QC_LABELS[qc];
    console.log(`  ${qc} (${label?.name || 'Unknown'}): ${stats.qcCounts[qc]} 個檔案`);
  }
  
  const output = toCSV(mappings);
  const lines = output.split('\n');
  console.log('');
  console.log('前 20 筆預覽:');
  lines.slice(0, 21).forEach(line => console.log('  ' + line));
  if (lines.length > 21) console.log(`  ... 還有 ${lines.length - 21} 筆`);
  
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const baseName = path.basename(rootDir);
  const outputFile = `${baseName}_field_mapping_${timestamp}.csv`;
  fs.writeFileSync(outputFile, '\uFEFF' + output, 'utf8');
  console.log('');
  console.log(`✅ 已保存到: ${outputFile}`);
}

main();
