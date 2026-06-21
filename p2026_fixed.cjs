const fs = require("fs");
const xlsx = require("xlsx");
const path = require("path");
const {encode_cell, encode_range, decode_range, aoa_to_sheet} = xlsx.utils;

const RD = "F:/Self-developed_Apps/FileName_WorkSheet_Extract/RawData";
const DE = "F:/Self-developed_Apps/FileName_WorkSheet_Extract/DataExtract";
const Y = "2026";

const DM = {};
DM["QC10001-R01"] = "樣品卡";
DM["QC10002-R01"] = "入庫品檢驗收單";
DM["QC10002-R02"] = "原物料品檢表";
DM["QC10002-R03"] = "原料進料審查規格";
DM["QC10004-R02"] = "QUALITY INSPECTION PLAN RECORD";
DM["QC10005-R01"] = "押出機每日巡檢表";
DM["QC10006-R01"] = "裝配對樣巡檢記錄表";
DM["QC10006-R02"] = "半成品品檢表";
DM["QC10007-R01"] = "完成品品檢表(首頁)";
DM["QC10007-R02"] = "完成品品檢表(綠頁)";
DM["QC10007-R03"] = "零組件入庫品檢表(射出零件品檢表?)";
DM["QC10008-R02"] = "出貨檢驗報告";

function scanFile(fp) {
  const wb = xlsx.readFile(fp);
  const res = [];
  const pat = /QC\d{5}-R\d{2}/i;
  wb.SheetNames.forEach((sn) => {
    const ws = wb.Sheets[sn];
    if (!ws["!ref"]) return;
    const meta = decode_range(ws["!ref"]);
    let code = "";
    for (let R = meta.s.r; R <= meta.e.r && !code; R++) {
      for (let C = meta.s.c; C <= meta.e.c && !code; C++) {
        const cell = ws[encode_cell({r: R, c: C})];
        if (cell && cell.v) {
          const m = String(cell.v).match(pat);
          if (m) code = m[0].toUpperCase();
        }
      }
    }
    res.push({sn, code, name: code ? (DM[code] || "未對照編碼") : ""});
  });
  return res;
}

function fileRow(f) {
  try {
    return scanFile(f).map((r) => [path.basename(f), r.sn, r.code || "\u7121", r.code ? r.name : "\u7121"]);
  } catch (e) {
    return [[path.basename(f), "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", e.message]];
  }
}

function findXlsx(base) {
  const tildeDollar = "~" + String.fromCharCode(36);
  return fs.readdirSync(base, {recursive: true}).filter((f) => {
    const fp = path.join(base, f);
    return fs.statSync(fp).isFile() && f.endsWith(".xlsx") && !f.includes(tildeDollar);
  });
}

function makeSheetWithData(rows) {
  return aoa_to_sheet(rows);
}

function addSheet(wb, ws, name) {
  wb.SheetNames.push(name);
  wb.Sheets[name] = ws;
}

function writeWB(wb, outPath) {
  const buf = xlsx.write(wb, {bookType: "xlsx", type: "buffer"});
  fs.writeFileSync(outPath, buf);
}

function aR(base, sub, outPath, mode) {
  const cp = path.join(base, sub);
  if (!fs.existsSync(cp)) { console.log("  SKIP:", sub); return; }
  const files = findXlsx(cp);
  if (files.length === 0) { console.log("  SKIP (no files):", sub); return; }
  const wb = {SheetNames: [], Sheets: {}};
  addSheet(wb, makeSheetWithData([["#", "標題"]]), "placeholder");
  if (mode === "monthly") {
    const groups = {};
    files.forEach((f) => {
      const rel = path.relative(cp, path.dirname(path.join(cp, f)));
      const top = rel === "" ? "(root)" : rel.split(path.sep)[0];
      const mm = top.match(/(\\d{2})$/);
      const sname = mm ? sub.split("-")[0] + "-" + mm[1] : top;
      if (!groups[sname]) groups[sname] = [];
      groups[sname].push(f);
    });
    Object.keys(groups).sort().forEach((gn) => {
      const rows = [["\u6a94\u540d\u7a31", "\u500b\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      groups[gn].forEach((f) => rows.push(...fileRow(path.join(cp, f))));
      const ws = makeSheetWithData(rows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, gn.substring(0, 31));
    });
  } else {
    const rows = [["\u6a94\u540d\u7a31", "\u500b\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
    files.forEach((f) => rows.push(...fileRow(path.join(cp, f))));
    const ws = makeSheetWithData(rows);
    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
    addSheet(wb, ws, sub.substring(0, 31));
  }
  writeWB(wb, outPath);
  console.log("  OK:", path.basename(outPath));
}

function sSub(base, subs, outPath) {
  const wb = {SheetNames: [], Sheets: {}};
  addSheet(wb, makeSheetWithData([["#", "標題"]]), "placeholder");
  subs.forEach((sf) => {
    const sp = path.join(base, sf);
    if (!fs.existsSync(sp)) { console.log("  SKIP DIR:", sf); return; }
    const files = findXlsx(sp);
    if (files.length === 0) { console.log("  SKIP (no files):", sf); return; }
    const rows = [["\u6a94\u540d\u7a31", "\u500b\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
    files.forEach((f) => rows.push(...fileRow(path.join(sp, f))));
    const ws = makeSheetWithData(rows);
    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
    addSheet(wb, ws, sf.substring(0, 31));
  });
  writeWB(wb, outPath);
  console.log("  OK:", path.basename(outPath));
}

function sSubFlat(base, prefix, outPath) {
  const wb = {SheetNames: [], Sheets: {}};
  addSheet(wb, makeSheetWithData([["#", "標題"]]), "placeholder");
  const dirs = fs.readdirSync(base, {withFileTypes: true})
    .filter((x) => x.isDirectory()).map((x) => x.name).sort();
  dirs.forEach((dn) => {
    const sp = path.join(base, dn);
    const files = findXlsx(sp);
    if (files.length === 0) return;
    const rows = [["\u6a94\u540d\u7a31", "\u500b\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
    files.forEach((f) => rows.push(...fileRow(path.join(sp, f))));
    const ws = makeSheetWithData(rows);
    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
    addSheet(wb, ws, (prefix + "-" + dn.replace(/\s/g, "")).substring(0, 31));
  });
  writeWB(wb, outPath);
  console.log("  OK:", path.basename(outPath));
}

console.log("========================================");
console.log("2026 \u5e74\u5ea6\u54c1\u7ba1\u5831\u8868\u7d71\u8a08\u5de5\u5177");
console.log("========================================");
const ed = path.join(DE, Y);
if (!fs.existsSync(ed)) fs.mkdirSync(ed, {recursive: true});

console.log("\n=== \u96f6\u7d44\u4ef6\u5165\u5eab-2026 ===");
aR(RD + "/" + "\u96f6\u7d44\u4ef6\u5165\u5eab-2026", "\u5c04\u51fa-2026", ed + "/零組件入庫-2026_射出.xlsx", "monthly");
aR(RD + "/零組件入庫-2026", "\u5c04\u51faC-2026", ed + "/零組件入庫-2026_射出C.xlsx", "perfile");
aR(RD + "/零組件入庫-2026", "\u5c04\u51faD-2026", ed + "/零組件入庫-2026_射出D.xlsx", "monthly");
aR(RD + "/零組件入庫-2026", "\u5c04\u51faD(\u7d44\u4ef6)-2026", ed + "/零組件入庫-2026_射出D(組件).xlsx", "perfile");
aR(RD + "/零組件入庫-2026", "\u88dd\u914dA-2026", ed + "/零組件入庫-2026_裝配A.xlsx", "perfile");
aR(RD + "/零組件入庫-2026", "\u88dd\u914dB-2026", ed + "/零組件入庫-2026_裝配B.xlsx", "perfile");
aR(RD + "/零組件入庫-2026", "\u88dd\u914dC-2026", ed + "/零組件入庫-2026_裝配C.xlsx", "perfile");
aR(RD + "/零組件入庫-2026", "Tubing-2026", ed + "/零組件入庫-2026_Tubing.xlsx", "monthly");

console.log("\n=== QIP\u5c3a\u5bf8\u6aa2\u9a57-2026 ===");
sSub(RD + "/QIP\u5c3a\u5bf8\u6aa2\u9a57-2026", ["QIP-2026(1~10)", "QIP-2026(\u5c04\u51faACD)", "QIP-SET UP-2026"], ed + "/QIP\u5c3a\u5bf8\u6aa2\u9a57-2026.xlsx");

console.log("\n=== \u5c04\u51fa\u6aa2\u9a57-2026 ===");
{
  const injBase = RD + "/\u5c04\u51fa\u6aa2\u9a57-2026/QIP-2026(1~10)";
  if (fs.existsSync(injBase)) {
    const wb = {SheetNames: [], Sheets: {}};
    addSheet(wb, makeSheetWithData([["#", "\u6a19\u984c"]]), "placeholder");
    const groups = fs.readdirSync(injBase, {withFileTypes: true})
      .filter((x) => x.isDirectory()).map((x) => x.name).sort();
    groups.forEach((gn) => {
      const gp = path.join(injBase, gn);
      const files = findXlsx(gp);
      if (files.length === 0) return;
      const rows = [["\u6a19\u984c"]];
      files.forEach((f) => rows.push(...fileRow(path.join(gp, f))));
      const ws = makeSheetWithData(rows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, gn.substring(0, 31));
    });
    writeWB(wb, ed + "/\u5c04\u51fa\u6aa2\u9a57-2026_QIP-2026(1~10).xlsx");
    console.log("  OK: " + path.basename(ed + "/射出檢驗-2026_QIP-2026(1~10).xlsx"));
  }
}

console.log("\n=== \u62bc\u51fa\u6aa2\u9a57-2026 ===");
sSubFlat(RD + "/" + "\u62bc\u51fa\u6aa2\u9a57-2026", "\u62bc\u51fa-QIP", ed + "/" + "\u62bc\u51fa\u6aa2\u9a57-2026.xlsx");

console.log("\n=== \u88dd\u914d\u5de1\u6aa2-2026 ===");
sSubFlat(RD + "/" + "\u88dd\u914d\u5de1\u6aa2-2026", "\u88dd\u914d\u5de1\u6aa2", ed + "/" + "\u88dd\u914d\u5de1\u6aa2-2026.xlsx");

console.log("\n=== \u88dd\u914d\u6aa2\u9a57-2026 ===");
sSubFlat(RD + "/" + "\u88dd\u914d\u6aa2\u9a57-2026", "\u88dd\u914d\u6aa2\u9a57", ed + "/" + "\u88dd\u914d\u6aa2\u9a57-2026.xlsx");

console.log("\n=== \u9032\u6599\u6aa2\u9a57-2026 ===");
sSubFlat(RD + "/" + "\u9032\u6599\u6aa2\u9a57-2026", "\u9032\u6599", ed + "/" + "\u9032\u6599\u6aa2\u9a57-2026.xlsx");

console.log("\n=== \u51fa\u8ca8\u6aa2\u9a57-2026 ===");
sSubFlat(RD + "/" + "\u51fa\u8ca8\u6aa2\u9a57-2026", "\u51fa\u8ca8", ed + "/" + "\u51fa\u8ca8\u6aa2\u9a57-2026.xlsx");
// ============ SUMMARY REPORT ============
console.log("\n=== \u751f\u6210 2026\u54c1\u6aa2\u5831\u8868\u7d71\u8a08.xlsx ===");

function countCodes(fp) {
  const wb = xlsx.readFile(fp);
  const codes = new Set();
  const pat = /QC\d{5}-R\d{2}/i;
  wb.SheetNames.forEach((sn) => {
    const ws = wb.Sheets[sn];
    if (!ws["!ref"]) return;
    const meta = decode_range(ws["!ref"]);
    for (let R = meta.s.r; R <= meta.e.r; R++) {
      for (let C = meta.s.c; C <= meta.e.c; C++) {
        const cell = ws[encode_cell({r: R, c: C})];
        if (cell && cell.v) {
          const m = String(cell.v).match(pat);
          if (m) codes.add(m[0].toUpperCase());
        }
      }
    }
  });
  return {sheets: wb.SheetNames.length, codes: codes.size, codeSet: codes};
}

function processCategory(catPath, catName, groupNames) {
  const totalRows = [["#", "\u5b50\u5206\u985e", "\u6a94\u6848\u6578", "\u7e3d\u5de5\u4f5c\u8868\u6578", "\u552f\u4e00QC\u7de8\u78bc\u6578"]];
  let totalFiles = 0, totalSheets = 0;
  const allCodes = new Set();
  groupNames.forEach((g, i) => {
    const gpath = path.join(catPath, g);
    if (!fs.existsSync(gpath)) return;
    const files = findXlsx(gpath);
    let sheets = 0;
    const codes = new Set();
    files.forEach((f) => {
      try {
        const fp = path.join(gpath, f);
        const info = countCodes(fp);
        sheets += info.sheets;
        info.codeSet.forEach((c) => codes.add(c));
      } catch(e) {}
    });
    totalFiles += files.length;
    totalSheets += sheets;
    codes.forEach((c) => allCodes.add(c));
    totalRows.push([i + 1, g, files.length, sheets, codes.size]);
  });
  totalRows.push(["", "\u5408\u8a08", totalFiles, totalSheets, allCodes.size]);
  return {rows: totalRows, files: totalFiles, sheets: totalSheets, codes: allCodes.size, codeSet: allCodes};
}

function processFlatCat(baseDir, prefix) {
  const totalRows = [["#", "\u5b50\u5206\u985e", "\u6a94\u6848\u6578", "\u7e3d\u5de5\u4f5c\u8868\u6578", "\u552f\u4e00QC\u7de8\u78bc\u6578"]];
  let totalFiles = 0, totalSheets = 0;
  const allCodes = new Set();
  const dirs = fs.readdirSync(baseDir, {withFileTypes: true})
    .filter((x) => x.isDirectory()).map((x) => x.name).sort();
  dirs.forEach((dn, i) => {
    const dp = path.join(baseDir, dn);
    const files = findXlsx(dp);
    if (files.length === 0) return;
    let sheets = 0;
    const codes = new Set();
    files.forEach((f) => {
      try {
        const info = countCodes(path.join(dp, f));
        sheets += info.sheets;
        info.codeSet.forEach((c) => codes.add(c));
      } catch(e) {}
    });
    totalFiles += files.length;
    totalSheets += sheets;
    codes.forEach((c) => allCodes.add(c));
    totalRows.push([i + 1, dn, files.length, sheets, codes.size]);
  });
  totalRows.push(["", "\u5408\u8a08", totalFiles, totalSheets, allCodes.size]);
  return {rows: totalRows, files: totalFiles, sheets: totalSheets, codes: allCodes.size, codeSet: allCodes, dirNames: dirs};
}

function writeSummarySheet(wb, catName, result) {
  const ws = aoa_to_sheet(result.rows);
  ws["!cols"] = [{wch: 5}, {wch: 25}, {wch: 10}, {wch: 12}, {wch: 12}];
  addSheet(wb, ws, catName.substring(0, 31));
  return result;
}

const summaryWb = {SheetNames: [], Sheets: {}};
addSheet(summaryWb, aoa_to_sheet([["#", "標題"]]), "placeholder");

// --- 零組件入庫 ---
const zcResult = processCategory(RD + "/" + "\u96f6\u7d44\u4ef6\u5165\u5eab-2026", "\u96f6\u7d44\u4ef6\u5165\u5eab", ["\u5c04\u51fa-2026", "\u5c04\u51faC-2026", "\u5c04\u51faD-2026", "\u5c04\u51faD(\u7d44\u4ef6)-2026", "\u88dd\u914dA-2026", "\u88dd\u914dB-2026", "\u88dd\u914dC-2026", "Tubing-2026"]);
writeSummarySheet(summaryWb, "\u96f6\u7d44\u4ef6\u5165\u5eab", zcResult);

// Monthly aggregation for 射出, 射出D, Tubing
const monthlyZC = [["\u6708\u4efd", "\u6a94\u6848\u6578", "\u5de5\u4f5c\u8868\u6578"]];
const zcMonthData = {};
["\u5c04\u51fa-2026", "\u5c04\u51faD-2026", "Tubing-2026"].forEach((g) => {
  const gpath = path.join(RD + "/" + "\u96f6\u7d44\u4ef6\u5165\u5eab-2026", g);
  if (!fs.existsSync(gpath)) return;
  const files = findXlsx(gpath);
  files.forEach((f) => {
    const rel = path.relative(gpath, path.dirname(path.join(gpath, f)));
    const top = rel === "" ? "(root)" : rel.split(path.sep)[0];
    const mm = top.match(/(\\d{2})$/);
    if (mm) {
      const m = mm[1];
      if (!zcMonthData[m]) zcMonthData[m] = {files: 0, sheets: 0};
      try {
        const info = countCodes(path.join(gpath, f));
        zcMonthData[m].files += 1;
        zcMonthData[m].sheets += info.sheets;
      } catch(e) {}
    }
  });
});
Object.keys(zcMonthData).sort().forEach((m) => {
  monthlyZC.push([m, zcMonthData[m].files, zcMonthData[m].sheets]);
});
if (monthlyZC.length > 1) {
  const mzWs = aoa_to_sheet(monthlyZC);
  mzWs["!cols"] = [{wch: 10}, {wch: 10}, {wch: 12}];
  addSheet(summaryWb, mzWs, "\u96f6\u7d44\u4ef6\u5165\u5eab\u6708\u8a08");
}

// --- QIP尺寸檢驗 ---
const qipResult = processCategory(RD + "/" + "QIP\u5c3a\u5bf8\u6aa2\u9a57-2026", "QIP\u5c3a\u5bf8\u6aa2\u9a57", ["QIP-2026(1~10)", "QIP-2026(\u5c04\u51faACD)", "QIP-SET UP-2026"]);
writeSummarySheet(summaryWb, "QIP\u5c3a\u5bf8\u6aa2\u9a57", qipResult);

const monthlyQIP = [["\u6708\u4efd", "\u6a94\u6848\u6578", "\u5de5\u4f5c\u8868\u6578"]];
const qipMonthData = {};
["QIP-2026(1~10)", "QIP-2026(\u5c04\u51faACD)", "QIP-SET UP-2026"].forEach((g) => {
  const gpath = path.join(RD + "/" + "QIP\u5c3a\u5bf8\u6aa2\u9a57-2026", g);
  if (!fs.existsSync(gpath)) return;
  const files = findXlsx(gpath);
  files.forEach((f) => {
    const rel = path.relative(gpath, path.dirname(path.join(gpath, f)));
    const top = rel === "" ? "(root)" : rel.split(path.sep)[0];
    const mm = top.match(/(\\d{2})$/);
    if (mm) {
      const m = mm[1];
      if (!qipMonthData[m]) qipMonthData[m] = {files: 0, sheets: 0};
      try {
        const info = countCodes(path.join(gpath, f));
        qipMonthData[m].files += 1;
        qipMonthData[m].sheets += info.sheets;
      } catch(e) {}
    }
  });
});
Object.keys(qipMonthData).sort().forEach((m) => {
  monthlyQIP.push([m, qipMonthData[m].files, qipMonthData[m].sheets]);
});
if (monthlyQIP.length > 1) {
  const mqWs = aoa_to_sheet(monthlyQIP);
  mqWs["!cols"] = [{wch: 10}, {wch: 10}, {wch: 12}];
  addSheet(summaryWb, mqWs, "QIP\u5c3a\u5bf8\u6aa2\u9a57\u6708\u8a08");
}

// --- 射出檢驗 ---
const injResult = processFlatCat(RD + "/" + "\u5c04\u51fa\u6aa2\u9a57-2026/QIP-2026(1~10)", "\u5c04\u51fa\u6aa2\u9a57");
writeSummarySheet(summaryWb, "\u5c04\u51fa\u6aa2\u9a57", injResult);

// --- 押出檢驗 ---
const outResult = processFlatCat(RD + "/" + "\u62bc\u51fa\u6aa2\u9a57-2026", "\u62bc\u51fa\u6aa2\u9a57");
writeSummarySheet(summaryWb, "\u62bc\u51fa\u6aa2\u9a57", outResult);

// --- 裝配巡檢 ---
const asResult = processFlatCat(RD + "/" + "\u88dd\u914d\u5de1\u6aa2-2026", "\u88dd\u914d\u5de1\u6aa2");
writeSummarySheet(summaryWb, "\u88dd\u914d\u5de1\u6aa2", asResult);

// --- 裝配檢驗 ---
const aqResult = processFlatCat(RD + "/" + "\u88dd\u914d\u6aa2\u9a57-2026", "\u88dd\u914d\u6aa2\u9a57");
writeSummarySheet(summaryWb, "\u88dd\u914d\u6aa2\u9a57", aqResult);

// --- 進料檢驗 ---
const inResult = processFlatCat(RD + "/" + "\u9032\u6599\u6aa2\u9a57-2026", "\u9032\u6599\u6aa2\u9a57");
writeSummarySheet(summaryWb, "\u9032\u6599\u6aa2\u9a57", inResult);

// --- 出貨檢驗 ---
const shResult = processFlatCat(RD + "/" + "\u51fa\u8ca8\u6aa2\u9a57-2026", "\u51fa\u8ca8\u6aa2\u9a57");
writeSummarySheet(summaryWb, "\u51fa\u8ca8\u6aa2\u9a57", shResult);

// --- 年度總覽 ---
const grandTotals = [
  ["#", "\u985e\u5225", "\u6a94\u6848\u6578", "\u5de5\u4f5c\u8868\u6578", "QC\u7de8\u78bc\u6578"],
  [1, "\u96f6\u7d44\u4ef6\u5165\u5eab", zcResult.files, zcResult.sheets, zcResult.codes],
  [2, "QIP\u5c3a\u5bf8\u6aa2\u9a57", qipResult.files, qipResult.sheets, qipResult.codes],
  [3, "\u5c04\u51fa\u6aa2\u9a57", injResult.files, injResult.sheets, injResult.codes],
  [4, "\u62bc\u51fa\u6aa2\u9a57", outResult.files, outResult.sheets, outResult.codes],
  [5, "\u88dd\u914d\u5de1\u6aa2", asResult.files, asResult.sheets, asResult.codes],
  [6, "\u88dd\u914d\u6aa2\u9a57", aqResult.files, aqResult.sheets, aqResult.codes],
  [7, "\u9032\u6599\u6aa2\u9a57", inResult.files, inResult.sheets, inResult.codes],
  [8, "\u51fa\u8ca8\u6aa2\u9a57", shResult.files, shResult.sheets, shResult.codes],
];
const grandFiles = zcResult.files + qipResult.files + injResult.files + outResult.files + asResult.files + aqResult.files + inResult.files + shResult.files;
const grandSheets = zcResult.sheets + qipResult.sheets + injResult.sheets + outResult.sheets + asResult.sheets + aqResult.sheets + inResult.sheets + shResult.sheets;
const grandCodes = new Set([...zcResult.codeSet, ...qipResult.codeSet, ...injResult.codeSet, ...outResult.codeSet, ...asResult.codeSet, ...aqResult.codeSet, ...inResult.codeSet, ...shResult.codeSet]);
grandTotals.push(["", "\u5e74\u5ea6\u5408\u8a08", grandFiles, grandSheets, grandCodes.size]);
const gtWs = aoa_to_sheet(grandTotals);
gtWs["!cols"] = [{wch: 5}, {wch: 20}, {wch: 10}, {wch: 12}, {wch: 12}];
addSheet(summaryWb, gtWs, "\u5e74\u5ea6\u7e3d\u89bd");

console.log("  \u7e3d\u6a94\u6848\u6578:", grandFiles);
console.log("  \u7e3d\u5de5\u4f5c\u8868\u6578:", grandSheets);
console.log("  \u7e3dQC\u7de8\u78bc\u6578:", grandCodes.size);

writeWB(summaryWb, ed + "/2026\u54c1\u6aa2\u5831\u8868\u7d71\u8a08.xlsx");
console.log("  OK: 2026\u54c1\u6aa2\u5831\u8868\u7d71\u8a08.xlsx");

console.log("\n========================================");
console.log("\u5168\u90e8\u5b8c\u6210!");
console.log("\u8f38\u51fa\u76ee\u9304:", ed);
console.log("========================================");