const fs = require('fs');

const Q = '"';
const NL = '\n';

// Helper to create unicode escapes
function uc(str) {
  return str.split('').map(c => {
    const code = c.charCodeAt(0);
    if (code > 127) return '\\u' + code.toString(16).padStart(4, '0');
    return c;
  }).join('');
}

const lines = [];
lines.push('const fs = require("fs");');
lines.push('const xlsx = require("xlsx");');
lines.push('const path = require("path");');
lines.push('const {encode_cell, encode_range, decode_range, aoa_to_sheet} = xlsx.utils;');
lines.push('');
lines.push('const RD = "F:/Self-developed_Apps/FileName_WorkSheet_Extract/RawData";');
lines.push('const DE = "F:/Self-developed_Apps/FileName_WorkSheet_Extract/DataExtract";');
lines.push('const Y = "2026";');
lines.push('');
lines.push('const DM = {};');
lines.push('DM["QC10001-R01"] = "樣品卡";');
lines.push('DM["QC10002-R01"] = "入庫品檢驗收單";');
lines.push('DM["QC10002-R02"] = "原物料品檢表";');
lines.push('DM["QC10002-R03"] = "原料進料審查規格";');
lines.push('DM["QC10004-R02"] = "QUALITY INSPECTION PLAN RECORD";');
lines.push('DM["QC10005-R01"] = "押出機每日巡檢表";');
lines.push('DM["QC10006-R01"] = "裝配對樣巡檢記錄表";');
lines.push('DM["QC10006-R02"] = "半成品品檢表";');
lines.push('DM["QC10007-R01"] = "完成品品檢表(首頁)";');
lines.push('DM["QC10007-R02"] = "完成品品檢表(綠頁)";');
lines.push('DM["QC10007-R03"] = "零組件入庫品檢表(射出零件品檢表?)";');
lines.push('DM["QC10008-R02"] = "出貨檢驗報告";');
lines.push('');
lines.push('function scanFile(fp) {');
lines.push('  const wb = xlsx.readFile(fp);');
lines.push('  const res = [];');
lines.push('  const pat = /QC' + '\\d' + '{5}-R' + '\\d' + '{2}/i;');
lines.push('  wb.SheetNames.forEach((sn) => {');
lines.push('    const ws = wb.Sheets[sn];');
lines.push('    if (!ws["!ref"]) return;');
lines.push('    const meta = decode_range(ws["!ref"]);');
lines.push('    let code = "";');
lines.push('    for (let R = meta.s.r; R <= meta.e.r && !code; R++) {');
lines.push('      for (let C = meta.s.c; C <= meta.e.c && !code; C++) {');
lines.push('        const cell = ws[encode_cell({r: R, c: C})];');
lines.push('        if (cell && cell.v) {');
lines.push('          const m = String(cell.v).match(pat);');
lines.push('          if (m) code = m[0].toUpperCase();');
lines.push('        }');
lines.push('      }');
lines.push('    }');
lines.push('    res.push({sn, code, name: code ? (DM[code] || "未對照編碼") : ""});');
lines.push('  });');
lines.push('  return res;');
lines.push('}');
lines.push('');
lines.push('function fileRow(f) {');
lines.push('  try {');
lines.push('    return scanFile(f).map((r) => [path.basename(f), r.sn, r.code || "\\u7121", r.code ? r.name : "\\u7121"]);');
lines.push('  } catch (e) {');
lines.push('    return [[path.basename(f), "\\u7121\\u6cd5\\u958b\\u555f", "\\u932f\\u8aa4", e.message]];');
lines.push('  }');
lines.push('}');
lines.push('');
lines.push('function findXlsx(base) {');
lines.push('  return fs.readdirSync(base, {recursive: true}).filter((f) => {');
lines.push('    const fp = path.join(base, f);');
lines.push('    return fs.statSync(fp).isFile() && f.endsWith(".xlsx") && !f.startsWith("~" + String.fromCharCode(36));');
lines.push('  });');
lines.push('}');
lines.push('');
lines.push('function makeSheetWithData(rows) {');
lines.push('  return aoa_to_sheet(rows);');
lines.push('}');
lines.push('');
lines.push('function addSheet(wb, ws, name) {');
lines.push('  wb.SheetNames.push(name);');
lines.push('  wb.Sheets[name] = ws;');
lines.push('}');
lines.push('');
lines.push('function writeWB(wb, outPath) {');
lines.push('  const buf = xlsx.write(wb, {bookType: "xlsx", type: "buffer"});');
lines.push('  fs.writeFileSync(outPath, buf);');
lines.push('}');
lines.push('');
lines.push('function aR(base, sub, outPath, mode) {');
lines.push('  const cp = path.join(base, sub);');
lines.push('  if (!fs.existsSync(cp)) { console.log("  SKIP:", sub); return; }');
lines.push('  const files = findXlsx(cp);');
lines.push('  if (files.length === 0) { console.log("  SKIP (no files):", sub); return; }');
lines.push('  const wb = {SheetNames: [], Sheets: {}};');
lines.push('  addSheet(wb, makeSheetWithData([["#", "標題"]]), "placeholder");');
lines.push('  if (mode === "monthly") {');
lines.push('    const groups = {};');
lines.push('    files.forEach((f) => {');
lines.push('      const rel = path.relative(cp, path.dirname(path.join(cp, f)));');
lines.push('      const top = rel === "" ? "(root)" : rel.split(path.sep)[0];');
lines.push('      const mm = top.match(/(\\' + '\\d' + '{2})$/);');
lines.push('      const sname = mm ? sub.split("-")[0] + "-" + mm[1] : top;');
lines.push('      if (!groups[sname]) groups[sname] = [];');
lines.push('      groups[sname].push(f);');
lines.push('    });');
lines.push('    Object.keys(groups).sort().forEach((gn) => {');
lines.push('      const rows = [["' + uc('檔名稱') + '", "' + uc('個工作表名稱') + '", "' + uc('表單編碼') + '", "' + uc('表單名稱') + '"]];');
lines.push('      groups[gn].forEach((f) => rows.push(...fileRow(path.join(cp, f))));');
lines.push('      const ws = makeSheetWithData(rows);');
lines.push('      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];');
lines.push('      addSheet(wb, ws, gn.substring(0, 31));');
lines.push('    });');
lines.push('  } else {');
lines.push('    const rows = [["' + uc('檔名稱') + '", "' + uc('個工作表名稱') + '", "' + uc('表單編碼') + '", "' + uc('表單名稱') + '"]];');
lines.push('    files.forEach((f) => rows.push(...fileRow(path.join(cp, f))));');
lines.push('    const ws = makeSheetWithData(rows);');
lines.push('    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];');
lines.push('    addSheet(wb, ws, sub.substring(0, 31));');
lines.push('  }');
lines.push('  writeWB(wb, outPath);');
lines.push('  console.log("  OK:", path.basename(outPath));');
lines.push('}');
lines.push('');
lines.push('function sSub(base, subs, outPath) {');
lines.push('  const wb = {SheetNames: [], Sheets: {}};');
lines.push('  addSheet(wb, makeSheetWithData([["#", "標題"]]), "placeholder");');
lines.push('  subs.forEach((sf) => {');
lines.push('    const sp = path.join(base, sf);');
lines.push('    if (!fs.existsSync(sp)) { console.log("  SKIP DIR:", sf); return; }');
lines.push('    const files = findXlsx(sp);');
lines.push('    if (files.length === 0) { console.log("  SKIP (no files):", sf); return; }');
lines.push('    const rows = [["' + uc('檔名稱') + '", "' + uc('個工作表名稱') + '", "' + uc('表單編碼') + '", "' + uc('表單名稱') + '"]];');
lines.push('    files.forEach((f) => rows.push(...fileRow(path.join(sp, f))));');
lines.push('    const ws = makeSheetWithData(rows);');
lines.push('    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];');
lines.push('    addSheet(wb, ws, sf.substring(0, 31));');
lines.push('  });');
lines.push('  writeWB(wb, outPath);');
lines.push('  console.log("  OK:", path.basename(outPath));');
lines.push('}');
lines.push('');
lines.push('function sSubFlat(base, prefix, outPath) {');
lines.push('  const wb = {SheetNames: [], Sheets: {}};');
lines.push('  addSheet(wb, makeSheetWithData([["#", "標題"]]), "placeholder");');
lines.push('  const dirs = fs.readdirSync(base, {withFileTypes: true})');
lines.push('    .filter((x) => x.isDirectory()).map((x) => x.name).sort();');
lines.push('  dirs.forEach((dn) => {');
lines.push('    const sp = path.join(base, dn);');
lines.push('    const files = findXlsx(sp);');
lines.push('    if (files.length === 0) return;');
lines.push('    const rows = [["' + uc('檔名稱') + '", "' + uc('個工作表名稱') + '", "' + uc('表單編碼') + '", "' + uc('表單名稱') + '"]];');
lines.push('    files.forEach((f) => rows.push(...fileRow(path.join(sp, f))));');
lines.push('    const ws = makeSheetWithData(rows);');
lines.push('    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];');
lines.push('    addSheet(wb, ws, (prefix + "-" + dn.replace(/' + '\\s' + '/g, "")).substring(0, 31));');
lines.push('  });');
lines.push('  writeWB(wb, outPath);');
lines.push('  console.log("  OK:", path.basename(outPath));');
lines.push('}');
lines.push('');

// Main execution section
lines.push('console.log("========================================");');
lines.push('console.log("2026 ' + uc('年度品管報表統計工具') + '");');
lines.push('console.log("========================================");');
lines.push('const ed = path.join(DE, Y);');
lines.push('if (!fs.existsSync(ed)) fs.mkdirSync(ed, {recursive: true});');
lines.push('');

// Helper to build Chinese strings for console.log and paths
function cn(str) {
  return str.split('').map(c => {
    const code = c.charCodeAt(0);
    if (code > 127) return '\\u' + code.toString(16).padStart(4, '0');
    return c;
  }).join('');
}

// 零組件入庫
lines.push('console.log("\\n=== ' + cn('零組件入庫-2026') + ' ===");');
lines.push('aR(RD + "/" + "' + cn('零組件入庫-2026') + '", "' + cn('射出-2026') + '", ed + "/零組件入庫-2026_射出.xlsx", "monthly");');
lines.push('aR(RD + "/零組件入庫-2026", "' + cn('射出C-2026') + '", ed + "/零組件入庫-2026_射出C.xlsx", "perfile");');
lines.push('aR(RD + "/零組件入庫-2026", "' + cn('射出D-2026') + '", ed + "/零組件入庫-2026_射出D.xlsx", "monthly");');
lines.push('aR(RD + "/零組件入庫-2026", "' + cn('射出D(組件)-2026') + '", ed + "/零組件入庫-2026_射出D(組件).xlsx", "perfile");');
lines.push('aR(RD + "/零組件入庫-2026", "' + cn('裝配A-2026') + '", ed + "/零組件入庫-2026_裝配A.xlsx", "perfile");');
lines.push('aR(RD + "/零組件入庫-2026", "' + cn('裝配B-2026') + '", ed + "/零組件入庫-2026_裝配B.xlsx", "perfile");');
lines.push('aR(RD + "/零組件入庫-2026", "' + cn('裝配C-2026') + '", ed + "/零組件入庫-2026_裝配C.xlsx", "perfile");');
lines.push('aR(RD + "/零組件入庫-2026", "Tubing-2026", ed + "/零組件入庫-2026_Tubing.xlsx", "monthly");');
lines.push('');

// QIP尺寸檢驗
lines.push('console.log("\\n=== QIP' + cn('尺寸檢驗-2026') + ' ===");');
lines.push('sSub(RD + "/QIP' + cn('尺寸檢驗-2026') + '", ["QIP-2026(1~10)", "QIP-2026(' + cn('射出ACD') + ')", "QIP-SET UP-2026"], ed + "/QIP' + cn('尺寸檢驗-2026') + '.xlsx");');
lines.push('');

// 射出檢驗
lines.push('console.log("\\n=== ' + cn('射出檢驗-2026') + ' ===");');
lines.push('{');
lines.push('  const injBase = RD + "/' + cn('射出檢驗-2026') + '/QIP-2026(1~10)";');
lines.push('  if (fs.existsSync(injBase)) {');
lines.push('    const wb = {SheetNames: [], Sheets: {}};');
lines.push('    addSheet(wb, makeSheetWithData([["#", "\\u6a19\\u984c"]]), "placeholder");');
lines.push('    const groups = fs.readdirSync(injBase, {withFileTypes: true})');
lines.push('      .filter((x) => x.isDirectory()).map((x) => x.name).sort();');
lines.push('    groups.forEach((gn) => {');
lines.push('      const gp = path.join(injBase, gn);');
lines.push('      const files = findXlsx(gp);');
lines.push('      if (files.length === 0) return;');
lines.push('      const rows = [["\\u6a19\\u984c"]];');
lines.push('      files.forEach((f) => rows.push(...fileRow(path.join(gp, f))));');
lines.push('      const ws = makeSheetWithData(rows);');
lines.push('      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];');
lines.push('      addSheet(wb, ws, gn.substring(0, 31));');
lines.push('    });');
lines.push('    writeWB(wb, ed + "/' + cn('射出檢驗-2026') + '_QIP-2026(1~10).xlsx");');
lines.push('    console.log("  OK: " + path.basename(ed + "/射出檢驗-2026_QIP-2026(1~10).xlsx"));');
lines.push('  }');
lines.push('}');
lines.push('');

// 押出檢驗
lines.push('console.log("\\n=== ' + cn('押出檢驗-2026') + ' ===");');
lines.push('sSubFlat(RD + "/" + "' + cn('押出檢驗-2026') + '", "' + cn('押出-QIP') + '", ed + "/" + "' + cn('押出檢驗-2026') + '.xlsx");');
lines.push('');

// 裝配巡檢
lines.push('console.log("\\n=== ' + cn('裝配巡檢-2026') + ' ===");');
lines.push('sSubFlat(RD + "/" + "' + cn('裝配巡檢-2026') + '", "' + cn('裝配巡檢') + '", ed + "/" + "' + cn('裝配巡檢-2026') + '.xlsx");');
lines.push('');

// 裝配檢驗
lines.push('console.log("\\n=== ' + cn('裝配檢驗-2026') + ' ===");');
lines.push('sSubFlat(RD + "/" + "' + cn('裝配檢驗-2026') + '", "' + cn('裝配檢驗') + '", ed + "/" + "' + cn('裝配檢驗-2026') + '.xlsx");');
lines.push('');

// 進料檢驗
lines.push('console.log("\\n=== ' + cn('進料檢驗-2026') + ' ===");');
lines.push('sSubFlat(RD + "/" + "' + cn('進料檢驗-2026') + '", "' + cn('進料') + '", ed + "/" + "' + cn('進料檢驗-2026') + '.xlsx");');
lines.push('');

// 出貨檢驗
lines.push('console.log("\\n=== ' + cn('出貨檢驗-2026') + ' ===");');
lines.push('sSubFlat(RD + "/" + "' + cn('出貨檢驗-2026') + '", "' + cn('出貨') + '", ed + "/" + "' + cn('出貨檢驗-2026') + '.xlsx");');

fs.writeFileSync("p2026.cjs", lines.join(NL), 'utf8');
console.log("Part 1 generated:", lines.length, "lines,",(lines.join(NL)).length, "bytes");
