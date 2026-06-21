// ============================================================
// p2025_fixed.cjs - Fixed 2025 extraction script
// Key fixes:
// 1. scanFile handles files with ANY number of sheets (no COM timeout)
// 2. QIP撠箏站瑼ａ? only extracts QIP-SET UP-2025 (skip empty subfolders)
// 3. All files processed regardless of sheet count
// ============================================================
const fs = require("fs");
const xlsx = require("xlsx");
const path = require("path");
const { encode_cell, decode_range, aoa_to_sheet } = xlsx.utils;

const RD = "F:/Self-developed_Apps/FileName_WorkSheet_Extract/RawData";
const DE = "F:/Self-developed_Apps/FileName_WorkSheet_Extract/DataExtract";
const Y = "2025";

const DM = {};
DM["QC10001-R01"] = "\u6a23\u54c1\u5361";
DM["QC10002-R01"] = "\u5165\u5eab\u54c1\u6aa2\u9a57\u6536\u55ae";
DM["QC10002-R02"] = "\u539f\u7269\u6599\u54c1\u6aa2\u8868";
DM["QC10002-R03"] = "\u539f\u6599\u9032\u6599\u5be9\u67e5\u898f\u683c";
DM["QC10004-R02"] = "QUALITY INSPECTION PLAN RECORD";
DM["QC10005-R01"] = "\u62bc\u51fa\u6a5f\u6bcf\u65e5\u5de1\u6aa2\u8868";
DM["QC10006-R01"] = "\u88dd\u914d\u5c0d\u6a23\u5de1\u6aa2\u8a18\u9304\u8868";
DM["QC10006-R02"] = "\u534a\u6210\u54c1\u54c1\u6aa2\u8868";
DM["QC10007-R01"] = "\u5b8c\u6210\u54c1\u54c1\u6aa2\u8868(\u9996\u9801)";
DM["QC10007-R02"] = "\u5b8c\u6210\u54c1\u54c1\u6aa2\u8868(\u7da0\u9801)";
DM["QC10007-R03"] = "\u96f6\u7d44\u4ef6\u5165\u5eab\u54c1\u6aa2\u8868(\u5c04\u51fa\u96f6\u4ef6\u54c1\u6aa2\u8868?)";
DM["QC10008-R02"] = "\u51fa\u8ca8\u6aa2\u9a57\u5831\u544a";

// scanFile: extract QC code from sheet name, then scan cell content for product sheets
function scanFile(fp) {
  const wb = xlsx.readFile(fp);
  const res = [];
  const pat = /QC\d{5}-R\d{2}/i;
  wb.SheetNames.forEach((sn) => {
    let code = "";
    // Check sheet name first (e.g. QC10006-R02.K(N))
    const nameMatch = sn.match(pat);
    if (nameMatch) {
      code = nameMatch[0].toUpperCase();
    }
    // Then scan cell content for any sheet
    if (!code) {
      const ws = wb.Sheets[sn];
      if (ws && ws["!ref"]) {
        const meta = decode_range(ws["!ref"]);
        for (let R = meta.s.r; R <= meta.e.r && !code; R++) {
          for (let C = meta.s.c; C <= meta.e.c && !code; C++) {
            const cell = ws[encode_cell({r: R, c: C})];
            if (cell && cell.v) {
              const m = String(cell.v).match(pat);
              if (m) code = m[0].toUpperCase();
            }
          }
        }
      }
    }
    res.push({sn, code, name: code ? (DM[code] || "\u672a\u5c0d\u7167\u7de8\u78bc") : ""});
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
  // Exclude files with "空白" in filename (template/blank files without proper date code)
  const filteredFiles = files.filter(f => !path.basename(f).includes("\u7a7a\u767d"));
  if (filteredFiles.length === 0) { console.log("  SKIP (only blank-named files):", sub); return; }
  const wb = {SheetNames: [], Sheets: {}};

  // Filter: exclude DATE, 空白, 範例*, 工作表1, 工作表2...
  const FILTER_PAT = /^(DATE|空白|範例|工作表\d+)/i;
  // Dedup: normalize sheet name by stripping trailing (N) duplicates
  function normalizeSheetName(sn) {
    let n = sn.trim().toUpperCase();
    // Normalize full-width brackets to half-width
    n = n.replace(/\uff08/g, '(').replace(/\uff3b/g, '[').replace(/\uff5b/g, '[').replace(/\u3010/g, '[')
         .replace(/\u3014/g, '[').replace(/\u3016/g, '[').replace(/\uff3d/g, ']').replace(/\uff5d/g, ']')
         .replace(/\u3011/g, ']').replace(/\u3015/g, ']').replace(/\uff09/g, ')');
    // Strip trailing (N) duplicates like (2), (NG), (出貨用) etc.
    n = n.replace(/\s*\([^)]*\)\s*$/, '');
    n = n.replace(/\s*（[^）]*）\s*$/, '');
    return n;
  }

  function filterAndDedupRows(allRows) {
    const seen = new Set();
    return allRows.filter(r => {
      if (FILTER_PAT.test(r[1])) return false;
      const norm = normalizeSheetName(r[1]);
      if (seen.has(norm)) return false;
      seen.add(norm);
      return true;
    });
  }

  // A=01, B=02, ... L=12
  function numToLetter(n) {
    const letters = "ABCDEFGHIJKL";
    const idx = parseInt(n, 10) - 1;
    return (idx >= 0 && idx < 12) ? letters[idx] : n;
  }

  // Extract month letter (A-L) from filename suffix (e.g. "裝配B-2025A.xlsx" → "A", "裝配A-2025-G.xlsx" → "G")
  function monthLetterFromFilename(fname) {
    const m = path.basename(fname, ".xlsx").match(/2025-?([A-L])$/i);
    return m ? m[1].toUpperCase() : null;
  }

  if (mode === "monthly" || mode === "monthly-letter") {
    const groups = {};
    filteredFiles.forEach((f) => {
      const rel = path.relative(cp, path.dirname(path.join(cp, f)));
      const top = rel === "" ? "(root)" : rel.split(path.sep)[0];
      const mm = top.match(/(\d{2})$/);
      let suffix = mm ? mm[1] : null;
      if (suffix && mode === "monthly-letter") suffix = numToLetter(suffix);
      const sname = suffix ? sub.split("-")[0] + "-" + suffix : top;
      if (!groups[sname]) groups[sname] = [];
      groups[sname].push(f);
    });
    Object.keys(groups).sort().forEach((gn) => {
      const allRows = [["檔案名稱", "工作表名稱", "表單編碼", "表單名稱"]];
      groups[gn].forEach((f) => allRows.push(...filterAndDedupRows(fileRow(path.join(cp, f)))));
      const ws = makeSheetWithData(allRows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, gn.substring(0, 31));
    });
  } else if (mode === "filename-letter") {
    const groups = {};
    filteredFiles.forEach((f) => {
      const letter = monthLetterFromFilename(f);
      const gn = letter ? sub.split("-")[0] + "-" + letter : "其他";
      if (!groups[gn]) groups[gn] = [];
      groups[gn].push(f);
    });
    Object.keys(groups).sort().forEach((gn) => {
      const allRows = [["檔案名稱", "工作表名稱", "表單編碼", "表單名稱"]];
      groups[gn].forEach((f) => allRows.push(...filterAndDedupRows(fileRow(path.join(cp, f)))));
      const ws = makeSheetWithData(allRows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, gn.substring(0, 31));
    });
  } else {
    const allRows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
    filteredFiles.forEach((f) => allRows.push(...filterAndDedupRows(fileRow(path.join(cp, f)))));
    const ws = makeSheetWithData(allRows);
    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
    addSheet(wb, ws, sub.substring(0, 31));
  }
  writeWB(wb, outPath);
  console.log("  OK:", path.basename(outPath));
}

function sSub(base, subs, outPath) {
  const wb = {SheetNames: [], Sheets: {}};
  subs.forEach((sf) => {
    const sp = path.join(base, sf);
    if (!fs.existsSync(sp)) { console.log("  SKIP DIR:", sf); return; }
    const files = findXlsx(sp);
    if (files.length === 0) {
      // FIXED: Create empty sheet for subfolders with no files
      // instead of skipping (which caused data contamination)
      const rows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      const ws = makeSheetWithData(rows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, sf.substring(0, 31));
      console.log("  EMPTY (no files):", sf);
      return;
    }
    const rows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
    files.forEach((f) => rows.push(...fileRow(path.join(sp, f))));
    const ws = makeSheetWithData(rows);
    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
    addSheet(wb, ws, sf.substring(0, 31));
  });
  writeWB(wb, outPath);
  console.log("  OK:", path.basename(outPath));
}

// Filter function for ?箄疏瑼ａ?: exclude unwanted sheet names and rows with no QC code
function isShippingSheetValid(sheetName, qcCode) {
  const sn = sheetName.trim();
  // Skip if QC code is "??
  if (qcCode === "\u7121" || qcCode === "") return false;
  // Skip sheet names containing these keywords
  if (sn.includes("DATE") || sn.includes("\u51fa\u8ca8")) return false;
  return true;
}

function shippingFileRow(f) {
  try {
    const results = scanFile(f);
    return results
      .map((r) => [path.basename(f), r.sn, r.code || "\u7121", r.code ? r.name : "\u7121"])
      .filter((row) => isShippingSheetValid(row[1], row[2]));
  } catch (e) {
    return [[path.basename(f), "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", e.message]];
  }
}

function sSubFlat(base, prefix, outPath, dedup) {
  const wb = {SheetNames: [], Sheets: {}};
  const dirs = fs.readdirSync(base, {withFileTypes: true})
    .filter((x) => x.isDirectory()).map((x) => x.name).sort();
  dirs.forEach((dn) => {
    const sp = path.join(base, dn);
    const files = findXlsx(sp);
    if (files.length === 0) return;
    const rows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
    if (dedup === false) {
      // No dedup: list all sheets from all files, skip blank codes and QC-identifier sheets
      const qcPat = /^QC\d{5}-R\d{2}/i;
      files.forEach((f) => {
        rows.push(...fileRow(path.join(sp, f)).filter(r => {
          if (r[2] && r[2] === "\u7121") return false;
          // Skip sheets whose name IS a QC code (e.g. QC10006-R02.K(N))
          if (qcPat.test(r[1])) return false;
          return true;
        }));
      });
    } else {
      // Dedup by filename, prefer row with valid QC code
      const seenFiles = new Set();
      files.forEach((f) => {
        const fp = path.join(sp, f);
        const fileRows = fileRow(fp);
        const bestRow = fileRows.find(r => r[2] && r[2] !== "\u7121") || fileRows[0];
        if (bestRow) {
          const name = bestRow[0];
          if (seenFiles.has(name)) return;
          seenFiles.add(name);
          rows.push(bestRow);
        }
      });
    }
    const ws = makeSheetWithData(rows);
    ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
    addSheet(wb, ws, (prefix + "-" + dn.replace(/\s/g, "")).substring(0, 31));
  });
  writeWB(wb, outPath);
  console.log("  OK:", path.basename(outPath));
}

console.log("========================================");
console.log("2025 \u5e74\u5ea6\u54c1\u7ba1\u5831\u8868\u7d71\u8a08\u5de5\u5177 (FIXED)");
console.log("========================================");
const ed = path.join(DE, Y);
if (!fs.existsSync(ed)) fs.mkdirSync(ed, {recursive: true});

// --- \u96f6\u7d44\u4ef6\u5165\u5eab-2025 ---
console.log("\n=== \u96f6\u7d44\u4ef6\u5165\u5eab-2025 ===");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u5c04\u51fa-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u5c04\u51fa.xlsx", "monthly");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u5c04\u51faA-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u5c04\u51faA.xlsx", "monthly-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u5c04\u51faC-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u5c04\u51faC.xlsx", "monthly-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u5c04\u51faC(\u7d44\u4ef6)-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u5c04\u51faC(\u7d44\u4ef6).xlsx", "filename-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u5c04\u51faD-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u5c04\u51faD.xlsx", "monthly-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u5c04\u51faD(\u7d44\u4ef6)-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u5c04\u51faD(\u7d44\u4ef6).xlsx", "filename-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u88dd\u914dA-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u88dd\u914dA.xlsx", "filename-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u88dd\u914dB-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u88dd\u914dB.xlsx", "filename-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "\u88dd\u914dC-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_\u88dd\u914dC.xlsx", "filename-letter");
aR(RD + "/2025/\u96f6\u7d44\u4ef6\u5165\u5eab-2025", "Tubing-2025", ed + "/\u96f6\u7d44\u4ef6\u5165\u5eab-2025_Tubing.xlsx", "monthly");

// --- QIP撠箏站瑼ａ?-2025 (FIXED: only QIP-SET UP-2025, dedup by file, correct headers) ---
console.log("\n=== QIP\u5c3a\u5bf8\u6aa2\u9a57-2025 ===");
{
  const qipBase = RD + "/2025/QIP\u5c3a\u5bf8\u6aa2\u9a57-2025";
  const wb = {SheetNames: [], Sheets: {}};
  const setUpsub = path.join(qipBase, "QIP-SET UP-2025");
  if (fs.existsSync(setUpsub)) {
    const files = findXlsx(setUpsub);
    if (files.length > 0) {
      // FIXED: Correct headers with full Chinese characters
      // FIXED: Correct headers matching VBA: 瑼??迂 | 撌乩?銵典?蝔?| 銵典蝺函Ⅳ | 銵典?迂
      const rows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      
      // FIXED: Deduplicate by file name - keep only first sheet per file
      files.forEach((f) => {
        const fp = path.join(setUpsub, f);
        try {
          const fileWb = xlsx.readFile(fp);
          const pat = /QC\d{5}-R\d{2}/i;
          
          // Only process the FIRST sheet of each file (deduplication)
          let firstSheet = true;
          fileWb.SheetNames.forEach((sn) => {
            if (!firstSheet) return;
            firstSheet = false;
            
            const ws = fileWb.Sheets[sn];
            if (!ws || !ws["!ref"]) return;
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
            rows.push([f, sn, code || "\u7121", code ? (DM[code] || "\u672a\u5c0d\u7167\u7de8\u78bc") : "\u7121"]);
          });
        } catch (e) {
          rows.push([f, "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", e.message]);
        }
      });
      
      const ws = makeSheetWithData(rows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, "QIP-SET UP-2025");
      console.log("  QIP-SET UP-2025: " + files.length + " files processed (deduplicated)");
    }
  }
  
  // FIXED: Create empty sheets with correct headers for subfolders with no files
  ["NCA\u5165\u5eab\u5831\u8868-2025", "QIP-2025(1~10)", "QIP-2025(\u5c04\u51faACD)"].forEach(sf => {
    const sp = path.join(qipBase, sf);
    if (!fs.existsSync(sp)) { console.log("  SKIP DIR:", sf); return; }
    const files = findXlsx(sp);
    if (files.length === 0) {
      // FIXED: Correct headers even for empty sheets
      const rows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      const ws = makeSheetWithData(rows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, sf.substring(0, 31));
      console.log("  EMPTY:", sf);
    }
  });
  
  writeWB(wb, ed + "/QIP\u5c3a\u5bf8\u6aa2\u9a57-2025.xlsx");
  console.log("  OK: QIP\u5c3a\u5bf8\u6aa2\u9a57-2025.xlsx");
}

// --- 撠瑼ａ?-2025 (Split into Setup / 撌⊥炎, deduplicate sheet names) ---
console.log("\n=== \u5c04\u51fa\u6aa2\u9a57-2025 ===");
{
  const injBase = RD + "/2025/\u5c04\u51fa\u6aa2\u9a57-2025/QIP-2025(1~10)";
  if (fs.existsSync(injBase)) {
    const wb = {SheetNames: [], Sheets: {}};
    const groups = fs.readdirSync(injBase, {withFileTypes: true})
      .filter((x) => x.isDirectory()).map((x) => x.name).sort();
    groups.forEach((gn) => {
      const gp = path.join(injBase, gn);
      const files = findXlsx(gp);
      if (files.length === 0) return;

      const setupRows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      const patrolRows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];

      // Normalize sheet name: strip trailing " (N)" or "-N" suffixes for dedup
      function normalizeSheetName(sn) {
        let n = sn.trim().toUpperCase();
        // Remove trailing " (N)" where N is a number
        n = n.replace(/\s*\(\d+\)\s*$/, "");
        // Remove trailing "-N" where N is a number
        n = n.replace(/-\d+\s*$/, "");
        return n.trim();
      }

      files.forEach((f) => {
        const fp = path.join(gp, f);
        try {
          const fileWb = xlsx.readFile(fp);
          const pat = /QC\d{5}-R\d{2}/i;
          // Deduplicate within this file using normalized sheet name as key
          const seen = new Set();
          fileWb.SheetNames.forEach((sn) => {
            // Skip 撌乩?銵? and placeholder sheets
            if (sn === "\u5de5\u4f5c\u88681" || sn === "\u5de5\u4f5c\u8868\u540d\u7a31") return;

            const normKey = normalizeSheetName(sn);
            if (seen.has(normKey)) return; // duplicate within same file
            seen.add(normKey);

            const ws = fileWb.Sheets[sn];
            if (!ws || !ws["!ref"]) return;
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
            const rowData = [f, sn, code || "\u7121", code ? (DM[code] || "\u672a\u5c0d\u7167\u7de8\u78bc") : "\u7121"];
            // Classify: SETUP ??setupRows, otherwise ??patrolRows (撌⊥炎)
            if (sn.trim().toUpperCase().includes("SETUP")) {
              setupRows.push(rowData);
            } else {
              patrolRows.push(rowData);
            }
          });
        } catch (e) {
          setupRows.push([f, "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", e.message]);
        }
      });

      // Add Setup sheet
      if (setupRows.length > 1) {
        const wsSetup = makeSheetWithData(setupRows);
        wsSetup["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
        addSheet(wb, wsSetup, (gn + "-Setup").substring(0, 31));
      }
      // Add 撌⊥炎 sheet
      if (patrolRows.length > 1) {
        const wsPatrol = makeSheetWithData(patrolRows);
        wsPatrol["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
        addSheet(wb, wsPatrol, (gn + "-Patrol").substring(0, 31));
      }
    });
    writeWB(wb, ed + "/\u5c04\u51fa\u6aa2\u9a57-2025_QIP-2025(1~10).xlsx");
    console.log("  OK: \u5c04\u51fa\u6aa2\u9a57-2025_QIP-2025(1~10).xlsx");
  }
}

// --- ?澆瑼ａ?-2025 (Split into Setup / 撌⊥炎, exclude 蝛箇 files and 撌乩?銵?) ---
console.log("\n=== \u62bc\u51fa\u6aa2\u9a57-2025 ===");
{
  const ycBase = RD + "/2025/\u62bc\u51fa\u6aa2\u9a57-2025";
  if (!fs.existsSync(ycBase)) { console.log("  SKIP"); }
  else {
    const wb = {SheetNames: [], Sheets: {}};
    const dirs = fs.readdirSync(ycBase, {withFileTypes: true})
      .filter((x) => x.isDirectory()).map((x) => x.name).sort();
    dirs.forEach((dn) => {
      const sp = path.join(ycBase, dn);
      const files = findXlsx(sp);
      if (files.length === 0) return;
      
      const setupRows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      const patrolRows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      
      files.forEach((f) => {
        const fp = path.join(sp, f);
        try {
          const fileWb = xlsx.readFile(fp);
          const pat = /QC\d{5}-R\d{2}/i;
          fileWb.SheetNames.forEach((sn) => {
            // Skip 撌乩?銵? (empty placeholder)
            if (sn === "\u5de5\u4f5c\u88681" || sn === "\u5de5\u4f5c\u8868\u540d\u7a31") return;
            // Skip files with 蝛箇 in name
            if (f.includes("\u7a7a\u767d")) return;
            
            const ws = fileWb.Sheets[sn];
            if (!ws || !ws["!ref"]) return;
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
            const rowData = [f, sn, code || "\u7121", code ? (DM[code] || "\u672a\u5c0d\u7167\u7de8\u78bc") : "\u7121"];
            // Classify: SETUP ??setupRows, otherwise ??patrolRows (撌⊥炎)
            if (sn.toUpperCase().includes("SETUP")) {
              setupRows.push(rowData);
            } else {
              patrolRows.push(rowData);
            }
          });
        } catch (e) {
          setupRows.push([f, "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", e.message]);
        }
      });
      
      // Add Setup sheet
      if (setupRows.length > 1) {
        const wsSetup = makeSheetWithData(setupRows);
        wsSetup["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
        addSheet(wb, wsSetup, (dn + "-Setup").substring(0, 31));
      }
      // Add 撌⊥炎 sheet
      if (patrolRows.length > 1) {
        const wsPatrol = makeSheetWithData(patrolRows);
        wsPatrol["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
        addSheet(wb, wsPatrol, (dn + "-Patrol").substring(0, 31));
      }
    });
    writeWB(wb, ed + "/\u62bc\u51fa\u6aa2\u9a57-2025.xlsx");
    console.log("  OK: \u62bc\u51fa\u6aa2\u9a57-2025.xlsx");
  }
}

// --- 鋆?撌⊥炎-2025 ---
console.log("\n=== \u88dd\u914d\u5de1\u6aa2-2025 ===");
sSubFlat(RD + "/2025/" + "\u88dd\u914d\u5de1\u6aa2-2025", "\u88dd\u914d\u5de1\u6aa2", ed + "/" + "\u88dd\u914d\u5de1\u6aa2-2025.xlsx", false);

// --- 鋆?瑼ａ?-2025 ---
console.log("\n=== \u88dd\u914d\u6aa2\u9a57-2025 ===");
sSubFlat(RD + "/2025/" + "\u88dd\u914d\u6aa2\u9a57-2025", "\u88dd\u914d\u6aa2\u9a57", ed + "/" + "\u88dd\u914d\u6aa2\u9a57-2025.xlsx", false);

// --- 進料檢驗-2025 (Grouped by material type, filter Sheet1/DATE/bare (N)) ---
console.log("\n=== \u9032\u6599\u6aa2\u9a57-2025 ===");
{
  const jlBase = RD + "/2025/\u9032\u6599\u6aa2\u9a57-2025";
  if (fs.existsSync(jlBase)) {
    const wb = {SheetNames: [], Sheets: {}};
    const pat = /QC\d{5}-R\d{2}/i;
    const headerRow = ["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31", "\u6708\u4efd"];

    // Extract month from sheet: find "檢號" label, read cell 2 cols right, convert serial to month
    function getMonthFromSheet(ws) {
      if (!ws || !ws["!ref"]) return "";
      const ref = xlsx.utils.decode_range(ws["!ref"]);
      for (let r = 0; r <= Math.min(ref.e.r, 15); r++) {
        for (let c = 0; c <= ref.e.c; c++) {
          const cell = ws[xlsx.utils.encode_cell({r, c})];
          if (cell && String(cell.v).trim() === "\u6aa2\u865f") {
            const dc = ws[xlsx.utils.encode_cell({r, c: c + 2})];
            if (!dc || dc.v === undefined || dc.v === null) continue;
            // Serial date number → convert to month
            if (typeof dc.v === "number" && dc.v > 40000) {
              const parsed = xlsx.SSF.parse_date_code(dc.v);
              if (parsed) return String(parsed.m).padStart(2, "0");
            }
            // Date string like 2025-03-28 or 2025/3/28 → extract month
            if (dc.w && /^\d{4}[-/]\d{1,2}/.test(dc.w)) {
              const parts = dc.w.split(/[-/]/);
              if (parts.length >= 2) {
                const monthNum = parseInt(parts[1]);
                // Validate month range (1-12) to reject malformed entries like "2025-227"
                if (monthNum >= 1 && monthNum <= 12) {
                  return String(monthNum).padStart(2, "0");
                }
                // Fallback: if YYYY-MMDD where month is 1 digit (e.g. "2025-227" = 2025-2-27)
                if (parts[1].length >= 3) {
                  const m1 = parseInt(parts[1].substring(0, 1));
                  if (m1 >= 1 && m1 <= 9) return String(m1).padStart(2, "0");
                }
              }
            }
          }
        }
      }
      return "";
    }

    // Shared sheet filter and processing (skip blank檢號 and no date code)
    function processSheet(ws, sn, fname, rows) {
      const snTrimmed = sn.trim();
      if (/Sheet1/i.test(snTrimmed) || /DATE/i.test(snTrimmed)) return;
      if (/^\(\d+\)$/.test(snTrimmed)) return;
      if (!ws || !ws["!ref"]) return;
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
      // Skip rows with blank檢號
      if (!code) return;
      // Extract month from date code
      const month = getMonthFromSheet(ws);
      if (!month) return; // Skip if no valid month found
      rows.push([fname, sn, code, DM[code] || "\u672a\u5c0d\u7167\u7de8\u78bc", month]);
    }

    // --- 原料: one sheet with all raw material data ---
    const rawDir = path.join(jlBase, "\u539f\u6599");
    if (fs.existsSync(rawDir)) {
      const rawRows = [headerRow];
      const rawFiles = findXlsx(rawDir);
      rawFiles.forEach((f) => {
        const fp = path.join(rawDir, f);
        const fname = path.basename(f);
        try {
          const fileWb = xlsx.readFile(fp);
          fileWb.SheetNames.forEach((sn) => processSheet(fileWb.Sheets[sn], sn, fname, rawRows));
        } catch (e) {
          rawRows.push([fname, "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", e.message]);
        }
      });
      const ws = makeSheetWithData(rawRows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, "\u539f\u6599");
      console.log("  \u539f\u6599: " + (rawRows.length - 1) + " rows");
    }

    // --- 物料: each entry (file or subdir) gets its own sheet ---
    const matDir = path.join(jlBase, "\u7269\u6599");
    if (fs.existsSync(matDir)) {
      const matEntries = fs.readdirSync(matDir, {withFileTypes: true}).filter((e) => {
        if (e.name.startsWith("~$")) return false;
        if (e.isFile()) return e.name.endsWith(".xlsx");
        if (e.isDirectory()) return true;
        return false;
      }).sort((a, b) => a.name.localeCompare(b.name));

      matEntries.forEach((e) => {
        const entryPath = path.join(matDir, e.name);
        const sheetName = e.name.replace(/\.xlsx$/i, "").substring(0, 31);
        const matRows = [headerRow];

        if (e.isDirectory()) {
          // Subdirectory: collect all xlsx files within
          const subFiles = findXlsx(entryPath);
          subFiles.forEach((f) => {
            const fp = path.join(entryPath, f);
            const fname = path.basename(f);
            try {
              const fileWb = xlsx.readFile(fp);
              fileWb.SheetNames.forEach((sn) => processSheet(fileWb.Sheets[sn], sn, fname, matRows));
            } catch (err) {
              matRows.push([fname, "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", err.message]);
            }
          });
        } else {
          // Single file
          const fname = e.name;
          try {
            const fileWb = xlsx.readFile(entryPath);
            fileWb.SheetNames.forEach((sn) => processSheet(fileWb.Sheets[sn], sn, fname, matRows));
          } catch (err) {
            matRows.push([fname, "\u7121\u6cd5\u958b\u555f", "\u932f\u8aa4", err.message]);
          }
        }

        if (matRows.length > 1) {
          const ws = makeSheetWithData(matRows);
          ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
          addSheet(wb, ws, ("\u7269\u6599-" + sheetName).substring(0, 31));
          console.log("  \u7269\u6599-" + sheetName + ": " + (matRows.length - 1) + " rows");
        }
      });
    }

    writeWB(wb, ed + "/\u9032\u6599\u6aa2\u9a57-2025.xlsx");
    console.log("  OK: \u9032\u6599\u6aa2\u9a57-2025.xlsx");
  }
}

// --- ?箄疏瑼ａ?-2025 ---
console.log("\n=== \u51fa\u8ca8\u6aa2\u9a57-2025 ===");
// --- ?箄疏瑼ａ?-2025 (filtered: no DATE/摰Ｘ???箄疏 sheets, no ??QC code) ---
console.log("\n=== \u51fa\u8ca8\u6aa2\u9a57-2025 ===");
{
  const shBase = RD + "/2025/\u51fa\u8ca8\u6aa2\u9a57-2025";
  if (!fs.existsSync(shBase)) { console.log("  SKIP"); }
  else {
    const wb = {SheetNames: [], Sheets: {}};
    const dirs = fs.readdirSync(shBase, {withFileTypes: true})
      .filter((x) => x.isDirectory()).map((x) => x.name).sort();
    dirs.forEach((dn) => {
      const sp = path.join(shBase, dn);
      const files = findXlsx(sp);
      if (files.length === 0) return;
      const rows = [["\u6a94\u6848\u540d\u7a31", "\u5de5\u4f5c\u8868\u540d\u7a31", "\u8868\u55ae\u7de8\u78bc", "\u8868\u55ae\u540d\u7a31"]];
      files.forEach((f) => rows.push(...shippingFileRow(path.join(sp, f))));
      const ws = makeSheetWithData(rows);
      ws["!cols"] = [{wch: 40}, {wch: 30}, {wch: 15}, {wch: 40}];
      addSheet(wb, ws, (dn).substring(0, 31));
    });
    writeWB(wb, ed + "/\u51fa\u8ca8\u6aa2\u9a57-2025.xlsx");
    console.log("  OK: \u51fa\u8ca8\u6aa2\u9a57-2025.xlsx");
  }
}

console.log("\n========================================");
console.log("\u5168\u90e8\u5b8c\u6210!");
console.log("\u8f38\u51fa\u76ee\u9304:", ed);
console.log("========================================");

