const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const reportPath = 'C:\\Users\\USER\\.gemini\\antigravity-ide\\brain\\d3a71cc0-7f63-4e2b-8f04-2690a18b1f76\\all_qip_sheets.md';

function getBaseSheetName(sheetName) {
  return sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
}

let md = `# 全年度射出品檢工作表完整明細 (2025 & 2026)\n\n`;
md += `本清單包含 2025 年及 2026 年射出巡檢資料夾下所有 Excel 檔案及其中提取的工作表與去重後的基準名稱。\n\n`;
md += `| 年份 | 月份 | 來源檔案名稱 | 原始工作表名稱 (Original Sheet) | 去重後基準名稱 (Base Sheet Name) | 是否計入巡檢統計 |\n`;
md += `| --- | --- | --- | --- | --- | --- |\n`;

let idx = 1;

[2025, 2026].forEach(year => {
  const targetDir = `RawData/${year}/射出檢驗-${year}/QIP-${year}(1~10)`;
  if (!fs.existsSync(targetDir)) {
    console.log(`Directory not found for ${year}`);
    return;
  }

  const subDirs = fs.readdirSync(targetDir).sort();
  subDirs.forEach(sub => {
    const fullSubPath = path.join(targetDir, sub);
    const stat = fs.statSync(fullSubPath);
    if (!stat.isDirectory()) return;

    const mMatch = sub.match(/-(\d{2})$/);
    if (!mMatch) return;

    const month = parseInt(mMatch[1], 10);
    if (month < 1 || month > 12) return;

    const files = fs.readdirSync(fullSubPath).filter(f => f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$')).sort();

    files.forEach(file => {
      let wb;
      try {
        wb = XLSX.readFile(path.join(fullSubPath, file));
      } catch (e) {
        return;
      }

      const uniqueBaseInFile = new Set();
      const fileRows = [];

      const isExtrusion = targetDir.indexOf('押出') >= 0;

      wb.SheetNames.forEach(sheetName => {
        const baseName = getBaseSheetName(sheetName);
        const snLower = sheetName.toLowerCase();
        const isSetup = (snLower.indexOf('setup') >= 0 || snLower.indexOf('set up') >= 0 || snLower.indexOf('set-up') >= 0 || snLower === 'setup');
        
        let status = '';
        if (isExtrusion) {
          if (isSetup) {
            status = '❌ Setup工作表 (Skip)';
          } else if (/^(工作表|Sheet)\d+/i.test(sheetName.trim())) {
            status = '❌ 預設未命名工作表 (Skip)';
          } else if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別' || sheetName.startsWith('Sheet1')) {
            status = '❌ 系統工作表 (Skip)';
          } else {
            if (!uniqueBaseInFile.has(baseName)) {
              uniqueBaseInFile.add(baseName);
              status = '✅ 計入 (首筆)';
            } else {
              status = '⚠️ 重複 (已歸併不計入)';
            }
          }
        } else {
          if (/^\d{6}[a-zA-Z]?$/.test(baseName)) {
            if (!uniqueBaseInFile.has(baseName)) {
              uniqueBaseInFile.add(baseName);
              status = '✅ 計入 (首筆)';
            } else {
              status = '⚠️ 重複 (已歸併不計入)';
            }
          } else {
            if (isSetup) {
              status = '❌ Setup工作表 (Skip)';
            } else if (/^(工作表|Sheet)\d+/i.test(sheetName.trim())) {
              status = '❌ 預設未命名工作表 (Skip)';
            } else if (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別' || sheetName.startsWith('Sheet1')) {
              status = '❌ 系統工作表 (Skip)';
            } else {
              status = '❌ 非Date Code格式 (Skip)';
            }
          }
        }

        fileRows.push({
          sheet: sheetName,
          base: baseName,
          status: status
        });
      });

      fileRows.forEach(row => {
        md += `| ${year} | ${month}月 | \`${file}\` | \`${row.sheet}\` | \`${row.base}\` | ${row.status} |\n`;
      });
    });
  });
});

fs.writeFileSync(reportPath, md, 'utf8');
console.log(`\nReport generated successfully!`);
console.log(`Saved report to: ${reportPath}`);
