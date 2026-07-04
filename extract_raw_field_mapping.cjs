/**
 * 原始資料欄位映射提取工具
 * 
 * 用途：掃描原始資料夾，提取每個檔案的 QC 標籤和欄位名稱
 * 輸出格式：品管標籤編號,欄位名稱,資料路徑
 * 
 * 使用方式：
 *   node extract_raw_field_mapping.cjs <原始資料路徑>
 *   node extract_raw_field_mapping.cjs "F:\2026 報表\進料檢驗-2026"
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ============================================================
// 工具函數
// ============================================================

/**
 * 從檔案路徑或內容推斷 QC 標籤
 */
function inferQCLabelFromPath(filePath, folderName) {
  const folderPath = path.dirname(filePath);
  const folderBasename = path.basename(folderPath);
  
  // 優先使用資料夾名稱判斷
  if (folderBasename === '原料' || folderBasename === '物料' || folderPath.indexOf('進料檢驗') >= 0) {
    return 'QC10002-R02';
  }
  if (folderBasename === '射出' || folderBasename === '押出' || folderBasename === 'QIP' || folderPath.indexOf('射出檢驗') >= 0 || folderPath.indexOf('押出檢驗') >= 0) {
    return 'QC10004-R02';
  }
  if (folderBasename === '裝配巡檢' || folderPath.indexOf('裝配巡檢') >= 0) {
    return 'QC10006-R01';
  }
  if (folderBasename === '半成品' || folderBasename === '裝配檢驗' || folderPath.indexOf('半成品檢驗') >= 0) {
    return 'QC10006-R02';
  }
  if (folderBasename === '完成品' || folderPath.indexOf('完成品') >= 0) {
    return 'QC10007-R01';
  }
  if (folderBasename === '零組件入庫' || folderBasename === 'Tubing' || folderBasename === '裝配A' || folderBasename === '裝配B' || folderBasename === '裝配C' || folderPath.indexOf('零組件入庫') >= 0) {
    return 'QC10007-R03';
  }
  if (folderBasename === '出貨檢驗' || folderBasename === 'ICU' || folderPath.indexOf('出貨檢驗') >= 0) {
    return 'QC10008-R02';
  }
  
  return 'Unknown';
}

/**
 * 從 Excel 檔案提取欄位名稱（第 2 行的欄位標題）
 */
function extractFieldNames(excelFilePath) {
  const fieldNames = [];
  
  try {
    const wb = XLSX.readFile(excelFilePath);
    
    wb.SheetNames.forEach(sheetName => {
      // 過濾特定名稱的工作表
      if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別') return;
      if (sheetName.indexOf('Sheet') >= 0) return;
      if (/^QC[-_]?\d+/i.test(sheetName.trim())) return;
      
      const sheet = wb.Sheets[sheetName];
      if (!sheet || !sheet['!ref']) return;
      
      // 從第 2 行（row index 1）提取欄位名稱
      const range = XLSX.utils.decode_range(sheet['!ref']);
      
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddr = XLSX.utils.encode_cell({ r: 1, c: col });
        const cell = sheet[cellAddr];
        
        if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
          fieldNames.push(String(cell.v).trim());
        }
      }
    });
  } catch (error) {
    console.error(`  ⚠️  讀取失敗: ${error.message}`);
  }
  
  return fieldNames;
}

/**
 * 遞迴掃描目錄下的所有 Excel 檔案
 */
function scanDirectory(dirPath) {
  const excelFiles = [];
  
  try {
    const items = fs.readdirSync(dirPath, { withFileTypes: true });
    
    items.forEach(item => {
      const fullPath = path.join(dirPath, item.name);
      
      if (item.isDirectory()) {
        // 遞迴掃描子目錄
        excelFiles.push(...scanDirectory(fullPath));
      } else if (item.isFile() && (item.name.endsWith('.xlsx') || item.name.endsWith('.xls'))) {
        excelFiles.push(fullPath);
      }
    });
  } catch (error) {
    console.error(`  ⚠️  掃描目錄失敗: ${error.message}`);
  }
  
  return excelFiles;
}

/**
 * 從 Excel 檔案提取所有欄位映射
 */
function extractAllMappings(rootDir) {
  if (!fs.existsSync(rootDir)) {
    console.error(`錯誤：目錄不存在 "${rootDir}"`);
    process.exit(1);
  }
  
  console.log(`📂 掃描目錄: ${rootDir}`);
  
  // 掃描所有 Excel 檔案
  const excelFiles = scanDirectory(rootDir);
  console.log(`  找到 ${excelFiles.length} 個 Excel 檔案`);
  
  const allMappings = [];
  
  excelFiles.forEach(filePath => {
    const folderName = path.basename(path.dirname(filePath));
    const qcLabel = inferQCLabelFromPath(filePath, folderName);
    
    // 提取欄位名稱
    const fieldNames = extractFieldNames(filePath);
    
    if (fieldNames.length > 0) {
      // 為每個欄位創建映射
      fieldNames.forEach(fieldName => {
        allMappings.push({
          qcLabel: qcLabel,
          fieldName: fieldName,
          dataPath: filePath
        });
      });
      
      console.log(`  ✓ ${path.basename(filePath)}: ${qcLabel} - ${fieldNames.length} 個欄位`);
    }
  });
  
  return allMappings;
}

/**
 * 將映射數據轉換為 CSV 格式
 */
function toCSV(mappings) {
  const header = '品管標籤編號,欄位名稱,資料路徑';
  const rows = mappings.map(m => {
    const qcLabel = `"${m.qcLabel}"`;
    const fieldName = `"${m.fieldName}"`;
    const dataPath = `"${m.dataPath}"`;
    return `${qcLabel},${fieldName},${dataPath}`;
  });
  
  return [header, ...rows].join('\n');
}

// ============================================================
// 主程式
// ============================================================

function main() {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('用法: node extract_raw_field_mapping.cjs <原始資料路徑>');
    console.log('');
    console.log('範例:');
    console.log('  node extract_raw_field_mapping.cjs "F:\\2026 報表\\進料檢驗-2026"');
    process.exit(0);
  }
  
  const rootDir = args[0];
  
  console.log('='.repeat(80));
  console.log('📊 原始資料欄位映射提取工具');
  console.log('='.repeat(80));
  console.log('');
  
  // 提取映射數據
  const startTime = Date.now();
  const mappings = extractAllMappings(rootDir);
  const extractTime = Date.now() - startTime;
  
  // 輸出統計資訊
  console.log('');
  console.log('='.repeat(80));
  console.log('📈 提取統計');
  console.log('='.repeat(80));
  console.log(`總映射數量: ${mappings.length}`);
  console.log(`提取耗時: ${extractTime}ms`);
  
  // 按 QC 標籤分組統計
  const qcGroups = {};
  mappings.forEach(m => {
    if (!qcGroups[m.qcLabel]) {
      qcGroups[m.qcLabel] = 0;
    }
    qcGroups[m.qcLabel]++;
  });
  
  console.log('');
  console.log('各 QC 標籤欄位數:');
  Object.entries(qcGroups).forEach(([qc, count]) => {
    console.log(`  ${qc}: ${count} 個欄位`);
  });
  
  // 生成輸出
  console.log('');
  console.log('='.repeat(80));
  console.log('📄 輸出結果');
  console.log('='.repeat(80));
  
  const output = toCSV(mappings);
  
  // 輸出到控制台
  console.log('');
  console.log(output);
  
  // 保存到檔案
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const baseName = path.basename(rootDir);
  const outputFile = `${baseName}_field_mapping_${timestamp}.csv`;
  
  fs.writeFileSync(outputFile, output, 'utf8');
  console.log('');
  console.log(`✅ 已保存到: ${outputFile}`);
  console.log('');
}

main();
