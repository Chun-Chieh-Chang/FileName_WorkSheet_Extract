const fs = require('fs');

const Q = '"';
const NL = '\n';

function cn(str) {
  return str.split('').map(c => {
    const code = c.charCodeAt(0);
    if (code > 127) return '\\u' + code.toString(16).padStart(4, '0');
    return c;
  }).join('');
}

const lines = [];

// genSummary function
lines.push('');
lines.push('// ============ SUMMARY REPORT ============');
lines.push('console.log("\\n=== ' + cn('生成') + ' 2026' + cn('品檢報表統計') + '.xlsx ===");');
lines.push('');
lines.push('function countCodes(fp) {');
lines.push('  const wb = xlsx.readFile(fp);');
lines.push('  const codes = new Set();');
lines.push('  const pat = /QC' + '\\d' + '{5}-R' + '\\d' + '{2}/i;');
lines.push('  wb.SheetNames.forEach((sn) => {');
lines.push('    const ws = wb.Sheets[sn];');
lines.push('    if (!ws["!ref"]) return;');
lines.push('    const meta = decode_range(ws["!ref"]);');
lines.push('    for (let R = meta.s.r; R <= meta.e.r; R++) {');
lines.push('      for (let C = meta.s.c; C <= meta.e.c; C++) {');
lines.push('        const cell = ws[encode_cell({r: R, c: C})];');
lines.push('        if (cell && cell.v) {');
lines.push('          const m = String(cell.v).match(pat);');
lines.push('          if (m) codes.add(m[0].toUpperCase());');
lines.push('        }');
lines.push('      }');
lines.push('    }');
lines.push('  });');
lines.push('  return {sheets: wb.SheetNames.length, codes: codes.size, codeSet: codes};');
lines.push('}');
lines.push('');
lines.push('function processCategory(catPath, catName, groupNames) {');
lines.push('  const totalRows = [["#", "' + cn('子分類') + '", "' + cn('檔案數') + '", "' + cn('總工作表數') + '", "' + cn('唯一QC編碼數') + '"]];');
lines.push('  let totalFiles = 0, totalSheets = 0;');
lines.push('  const allCodes = new Set();');
lines.push('  groupNames.forEach((g, i) => {');
lines.push('    const gpath = path.join(catPath, g);');
lines.push('    if (!fs.existsSync(gpath)) return;');
lines.push('    const files = findXlsx(gpath);');
lines.push('    let sheets = 0;');
lines.push('    const codes = new Set();');
lines.push('    files.forEach((f) => {');
lines.push('      try {');
lines.push('        const info = countCodes(path.join(gpath, f));');
lines.push('        sheets += info.sheets;');
lines.push('        info.codeSet.forEach((c) => codes.add(c));');
lines.push('      } catch(e) {}');
lines.push('    });');
lines.push('    totalFiles += files.length;');
lines.push('    totalSheets += sheets;');
lines.push('    codes.forEach((c) => allCodes.add(c));');
lines.push('    totalRows.push([i + 1, g, files.length, sheets, codes.size]);');
lines.push('  });');
lines.push('  totalRows.push(["", "' + cn('合計') + '", totalFiles, totalSheets, allCodes.size]);');
lines.push('  return {rows: totalRows, files: totalFiles, sheets: totalSheets, codes: allCodes.size, codeSet: allCodes};');
lines.push('}');
lines.push('');
lines.push('function processFlatCat(baseDir, prefix) {');
lines.push('  const totalRows = [["#", "' + cn('子分類') + '", "' + cn('檔案數') + '", "' + cn('總工作表數') + '", "' + cn('唯一QC編碼數') + '"]];');
lines.push('  let totalFiles = 0, totalSheets = 0;');
lines.push('  const allCodes = new Set();');
lines.push('  const dirs = fs.readdirSync(baseDir, {withFileTypes: true})');
lines.push('    .filter((x) => x.isDirectory()).map((x) => x.name).sort();');
lines.push('  dirs.forEach((dn, i) => {');
lines.push('    const dp = path.join(baseDir, dn);');
lines.push('    const files = findXlsx(dp);');
lines.push('    if (files.length === 0) return;');
lines.push('    let sheets = 0;');
lines.push('    const codes = new Set();');
lines.push('    files.forEach((f) => {');
lines.push('      try {');
lines.push('        const info = countCodes(path.join(dp, f));');
lines.push('        sheets += info.sheets;');
lines.push('        info.codeSet.forEach((c) => codes.add(c));');
lines.push('      } catch(e) {}');
lines.push('    });');
lines.push('    totalFiles += files.length;');
lines.push('    totalSheets += sheets;');
lines.push('    codes.forEach((c) => allCodes.add(c));');
lines.push('    totalRows.push([i + 1, dn, files.length, sheets, codes.size]);');
lines.push('  });');
lines.push('  totalRows.push(["", "' + cn('合計') + '", totalFiles, totalSheets, allCodes.size]);');
lines.push('  return {rows: totalRows, files: totalFiles, sheets: totalSheets, codes: allCodes.size, codeSet: allCodes, dirNames: dirs};');
lines.push('}');
lines.push('');
lines.push('function writeSummarySheet(wb, catName, result) {');
lines.push('  const ws = aoa_to_sheet(result.rows);');
lines.push('  ws["!cols"] = [{wch: 5}, {wch: 25}, {wch: 10}, {wch: 12}, {wch: 12}];');
lines.push('  addSheet(wb, ws, catName.substring(0, 31));');
lines.push('  return result;');
lines.push('}');
lines.push('');
lines.push('const summaryWb = {SheetNames: [], Sheets: {}};');
lines.push('addSheet(summaryWb, aoa_to_sheet([["#", "標題"]]), "placeholder");');
lines.push('');

// 零組件入庫
lines.push('// --- 零組件入庫 ---');
lines.push('const zcResult = processCategory(RD + "/" + "' + cn('零組件入庫-2026') + '", "' + cn('零組件入庫') + '", ["' + cn('射出-2026') + '", "' + cn('射出C-2026') + '", "' + cn('射出D-2026') + '", "' + cn('射出D(組件)-2026') + '", "' + cn('裝配A-2026') + '", "' + cn('裝配B-2026') + '", "' + cn('裝配C-2026') + '", "Tubing-2026"]);');
lines.push('writeSummarySheet(summaryWb, "' + cn('零組件入庫') + '", zcResult);');
lines.push('');

// Monthly aggregation for 零組件入庫
lines.push('// Monthly aggregation for 射出, 射出D, Tubing');
lines.push('const monthlyZC = [["' + cn('月份') + '", "' + cn('檔案數') + '", "' + cn('工作表數') + '"]];');
lines.push('const zcMonthData = {};');
lines.push('["' + cn('射出-2026') + '", "' + cn('射出D-2026') + '", "Tubing-2026"].forEach((g) => {');
lines.push('  const gpath = path.join(RD + "/" + "' + cn('零組件入庫-2026') + '", g);');
lines.push('  if (!fs.existsSync(gpath)) return;');
lines.push('  const files = findXlsx(gpath);');
lines.push('  files.forEach((f) => {');
lines.push('    const rel = path.relative(gpath, path.dirname(path.join(gpath, f)));');
lines.push('    const top = rel === "" ? "(root)" : rel.split(path.sep)[0];');
lines.push('    const mm = top.match(/(\\' + '\\d' + '{2})$/);');
lines.push('    if (mm) {');
lines.push('      const m = mm[1];');
lines.push('      if (!zcMonthData[m]) zcMonthData[m] = {files: 0, sheets: 0};');
lines.push('      try {');
lines.push('        const info = countCodes(path.join(gpath, f));');
lines.push('        zcMonthData[m].files += 1;');
lines.push('        zcMonthData[m].sheets += info.sheets;');
lines.push('      } catch(e) {}');
lines.push('    }');
lines.push('  });');
lines.push('});');
lines.push('Object.keys(zcMonthData).sort().forEach((m) => {');
lines.push('  monthlyZC.push([m, zcMonthData[m].files, zcMonthData[m].sheets]);');
lines.push('});');
lines.push('if (monthlyZC.length > 1) {');
lines.push('  const mzWs = aoa_to_sheet(monthlyZC);');
lines.push('  mzWs["!cols"] = [{wch: 10}, {wch: 10}, {wch: 12}];');
lines.push('  addSheet(summaryWb, mzWs, "' + cn('零組件入庫月計') + '");');
lines.push('}');
lines.push('');

// QIP尺寸檢驗
lines.push('// --- QIP尺寸檢驗 ---');
lines.push('const qipResult = processCategory(RD + "/" + "' + cn('QIP尺寸檢驗-2026') + '", "' + cn('QIP尺寸檢驗') + '", ["' + cn('QIP-2026(1~10)') + '", "' + cn('QIP-2026(射出ACD)') + '", "' + cn('QIP-SET UP-2026') + '"]);');
lines.push('writeSummarySheet(summaryWb, "' + cn('QIP尺寸檢驗') + '", qipResult);');
lines.push('');

// Monthly QIP
lines.push('const monthlyQIP = [["' + cn('月份') + '", "' + cn('檔案數') + '", "' + cn('工作表數') + '"]];');
lines.push('const qipMonthData = {};');
lines.push('["' + cn('QIP-2026(1~10)') + '", "' + cn('QIP-2026(射出ACD)') + '", "' + cn('QIP-SET UP-2026') + '"].forEach((g) => {');
lines.push('  const gpath = path.join(RD + "/" + "' + cn('QIP尺寸檢驗-2026') + '", g);');
lines.push('  if (!fs.existsSync(gpath)) return;');
lines.push('  const files = findXlsx(gpath);');
lines.push('  files.forEach((f) => {');
lines.push('    const rel = path.relative(gpath, path.dirname(path.join(gpath, f)));');
lines.push('    const top = rel === "" ? "(root)" : rel.split(path.sep)[0];');
lines.push('    const mm = top.match(/(\\' + '\\d' + '{2})$/);');
lines.push('    if (mm) {');
lines.push('      const m = mm[1];');
lines.push('      if (!qipMonthData[m]) qipMonthData[m] = {files: 0, sheets: 0};');
lines.push('      try {');
lines.push('        const info = countCodes(path.join(gpath, f));');
lines.push('        qipMonthData[m].files += 1;');
lines.push('        qipMonthData[m].sheets += info.sheets;');
lines.push('      } catch(e) {}');
lines.push('    }');
lines.push('  });');
lines.push('});');
lines.push('Object.keys(qipMonthData).sort().forEach((m) => {');
lines.push('  monthlyQIP.push([m, qipMonthData[m].files, qipMonthData[m].sheets]);');
lines.push('});');
lines.push('if (monthlyQIP.length > 1) {');
lines.push('  const mqWs = aoa_to_sheet(monthlyQIP);');
lines.push('  mqWs["!cols"] = [{wch: 10}, {wch: 10}, {wch: 12}];');
lines.push('  addSheet(summaryWb, mqWs, "' + cn('QIP尺寸檢驗月計') + '");');
lines.push('}');
lines.push('');

// 射出檢驗
lines.push('// --- 射出檢驗 ---');
lines.push('const injResult = processFlatCat(RD + "/" + "' + cn('射出檢驗-2026') + '/QIP-2026(1~10)", "' + cn('射出檢驗') + '");');
lines.push('writeSummarySheet(summaryWb, "' + cn('射出檢驗') + '", injResult);');
lines.push('');

// 押出檢驗
lines.push('// --- 押出檢驗 ---');
lines.push('const outResult = processFlatCat(RD + "/" + "' + cn('押出檢驗-2026') + '", "' + cn('押出檢驗') + '");');
lines.push('writeSummarySheet(summaryWb, "' + cn('押出檢驗') + '", outResult);');
lines.push('');

// 裝配巡檢
lines.push('// --- 裝配巡檢 ---');
lines.push('const asResult = processFlatCat(RD + "/" + "' + cn('裝配巡檢-2026') + '", "' + cn('裝配巡檢') + '");');
lines.push('writeSummarySheet(summaryWb, "' + cn('裝配巡檢') + '", asResult);');
lines.push('');

// 裝配檢驗
lines.push('// --- 裝配檢驗 ---');
lines.push('const aqResult = processFlatCat(RD + "/" + "' + cn('裝配檢驗-2026') + '", "' + cn('裝配檢驗') + '");');
lines.push('writeSummarySheet(summaryWb, "' + cn('裝配檢驗') + '", aqResult);');
lines.push('');

// 進料檢驗
lines.push('// --- 進料檢驗 ---');
lines.push('const inResult = processFlatCat(RD + "/" + "' + cn('進料檢驗-2026') + '", "' + cn('進料檢驗') + '");');
lines.push('writeSummarySheet(summaryWb, "' + cn('進料檢驗') + '", inResult);');
lines.push('');

// 出貨檢驗
lines.push('// --- 出貨檢驗 ---');
lines.push('const shResult = processFlatCat(RD + "/" + "' + cn('出貨檢驗-2026') + '", "' + cn('出貨檢驗') + '");');
lines.push('writeSummarySheet(summaryWb, "' + cn('出貨檢驗') + '", shResult);');
lines.push('');

// Annual overview
lines.push('// --- 年度總覽 ---');
lines.push('const grandTotals = [');
lines.push('  ["#", "' + cn('類別') + '", "' + cn('檔案數') + '", "' + cn('工作表數') + '", "' + cn('QC編碼數') + '"],');
lines.push('  [1, "' + cn('零組件入庫') + '", zcResult.files, zcResult.sheets, zcResult.codes],');
lines.push('  [2, "' + cn('QIP尺寸檢驗') + '", qipResult.files, qipResult.sheets, qipResult.codes],');
lines.push('  [3, "' + cn('射出檢驗') + '", injResult.files, injResult.sheets, injResult.codes],');
lines.push('  [4, "' + cn('押出檢驗') + '", outResult.files, outResult.sheets, outResult.codes],');
lines.push('  [5, "' + cn('裝配巡檢') + '", asResult.files, asResult.sheets, asResult.codes],');
lines.push('  [6, "' + cn('裝配檢驗') + '", aqResult.files, aqResult.sheets, aqResult.codes],');
lines.push('  [7, "' + cn('進料檢驗') + '", inResult.files, inResult.sheets, inResult.codes],');
lines.push('  [8, "' + cn('出貨檢驗') + '", shResult.files, shResult.sheets, shResult.codes],');
lines.push('];');
lines.push('const grandFiles = zcResult.files + qipResult.files + injResult.files + outResult.files + asResult.files + aqResult.files + inResult.files + shResult.files;');
lines.push('const grandSheets = zcResult.sheets + qipResult.sheets + injResult.sheets + outResult.sheets + asResult.sheets + aqResult.sheets + inResult.sheets + shResult.sheets;');
lines.push('const grandCodes = new Set([...zcResult.codeSet, ...qipResult.codeSet, ...injResult.codeSet, ...outResult.codeSet, ...asResult.codeSet, ...aqResult.codeSet, ...inResult.codeSet, ...shResult.codeSet]);');
lines.push('grandTotals.push(["", "' + cn('年度合計') + '", grandFiles, grandSheets, grandCodes.size]);');
lines.push('const gtWs = aoa_to_sheet(grandTotals);');
lines.push('gtWs["!cols"] = [{wch: 5}, {wch: 20}, {wch: 10}, {wch: 12}, {wch: 12}];');
lines.push('addSheet(summaryWb, gtWs, "' + cn('年度總覽') + '");');
lines.push('');
lines.push('console.log("  ' + cn('總檔案數') + ':", grandFiles);');
lines.push('console.log("  ' + cn('總工作表數') + ':", grandSheets);');
lines.push('console.log("  ' + cn('總QC編碼數') + ':", grandCodes.size);');
lines.push('');
lines.push('writeWB(summaryWb, ed + "/2026' + cn('品檢報表統計') + '.xlsx");');
lines.push('console.log("  OK: 2026' + cn('品檢報表統計') + '.xlsx");');
lines.push('');
lines.push('console.log("\\n========================================");');
lines.push('console.log("' + cn('全部完成') + '!");');
lines.push('console.log("' + cn('輸出目錄') + ':", ed);');
lines.push('console.log("========================================");');

fs.appendFileSync("p2026.cjs", lines.join(NL), 'utf8');
console.log("Summary report appended. Total lines:", lines.length);
