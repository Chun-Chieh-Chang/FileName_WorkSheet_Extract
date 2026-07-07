import * as XLSX from 'xlsx';
import { 
  detectQCFromFolder, 
  determineQCFromSheet, 
  getRawSubCategory, 
  extractRawMonth,
  findDateInSheet 
} from './browserETL';

const LETTER_MONTH = { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, I: 9, J: 10, K: 11, L: 12 };

/**
 * Parses a single Excel file, scanning all sheets for QC codes.
 * @param {File} file - The file object to parse.
 * @param {Object} mappings - The QC code mappings map.
 * @param {number} year - The target year for ETL tracking.
 * @returns {Promise<Object>} Resolves with the parsing results.
 */
export const parseExcelFile = (file, mappings, year = new Date().getFullYear()) => {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const results = [];
        
        // Track duplicates in the same file to mimic ETL deduplication
        const seenQC7R1BaseNames = new Set();
        const seenInjPatrolBaseNames = new Set();
        const seenExtPatrolBaseNames = new Set();

        const filePath = file.webkitRelativePath || file.name;
        const normalizedPath = filePath.replace(/\\/g, '/');
        const pathLower = normalizedPath.toLowerCase();
        
        const isInjection = pathLower.includes('射出檢驗-' + year);
        const isExtrusion = pathLower.includes('押出檢驗-' + year);
        
        workbook.SheetNames.forEach((sheetName) => {
          const ws = workbook.Sheets[sheetName];
          let foundCode = "";
          let foundName = "";
          
          if (ws) {
            const regEx = /QC\d{5}-R\d{2}/i;
            const cellKeys = Object.keys(ws);
            for (let key of cellKeys) {
              if (key.startsWith('!')) continue;
              const cell = ws[key];
              if (cell) {
                const val = String(cell.w || cell.v || '');
                const match = val.match(regEx);
                if (match) {
                  foundCode = match[0].toUpperCase();
                  foundName = mappings[foundCode] || "未對照編碼";
                  break;
                }
              }
            }
          }
          
          // Calculate ETL Inclusion status
          let etlStatus = "未納入";
          const etlTimestamp = new Date().toLocaleString('zh-TW', { hour12: false });
          
          if (isInjection) {
            // QIP Injection ETL validation logic
            const parts = normalizedPath.split('/');
            if (parts.length >= 2) {
              const parentDir = parts[parts.length - 2];
              const mMatch = parentDir.match(/-(\d{2})$/);
              if (!mMatch) {
                etlStatus = "狀態異常";
              } else {
                const month = parseInt(mMatch[1], 10);
                if (month < 1 || month > 12) {
                  etlStatus = "狀態異常";
                } else {
                  const isPatrol = pathLower.includes('qip-' + year + '(1~10)') || pathLower.includes('qip-' + year + '(1-10)');
                  if (isPatrol) {
                    const baseName = sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
                    if (/^\d{6}[a-zA-Z]?$/.test(baseName)) {
                      if (!seenInjPatrolBaseNames.has(baseName)) {
                        seenInjPatrolBaseNames.add(baseName);
                        etlStatus = "已納入";
                      } else {
                        etlStatus = "未納入"; // duplicate sheet in same file
                      }
                    } else {
                      etlStatus = "未納入"; // invalid sheet name (Date Code constraint)
                    }
                  } else {
                    // Setup counts count files, so sheets are marked as included under Setup
                    etlStatus = "已納入";
                  }
                }
              }
            } else {
              etlStatus = "狀態異常";
            }
          } else if (isExtrusion) {
            // QIP Extrusion ETL validation logic
            const parts = normalizedPath.split('/');
            if (parts.length >= 2) {
              const parentDir = parts[parts.length - 2];
              const mMatch = parentDir.match(/-(\d{2})$/);
              if (!mMatch) {
                etlStatus = "狀態異常";
              } else {
                const month = parseInt(mMatch[1], 10);
                if (month < 1 || month > 12) {
                  etlStatus = "狀態異常";
                } else {
                  const isDateCodeFile = /\d{6}[a-zA-Z]?/i.test(file.name);
                  if (!isDateCodeFile) {
                    etlStatus = "未納入"; // file skipped (does not match Date Code file constraint)
                  } else {
                    const isSkip = (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別' || sheetName.indexOf('Sheet1') === 0 || sheetName.indexOf('.K(') >= 0 || sheetName.indexOf('範例樣本') >= 0 || /^QC[-_]?\d+/i.test(sheetName.trim()) || /^(工作表|Sheet)\d+/i.test(sheetName.trim()) || /^(工作表|Sheet)/i.test(sheetName.trim()));
                    if (isSkip) {
                      etlStatus = "未納入";
                    } else {
                      const snLower = sheetName.toLowerCase();
                      const isSetup = (snLower.indexOf('setup') >= 0 || snLower.indexOf('set up') >= 0 || snLower.indexOf('set-up') >= 0 || snLower === 'setup');
                      if (isSetup) {
                        etlStatus = "未納入"; // Setup sheet itself is not counted as Patrol
                      } else {
                        const baseName = sheetName.replace(/(?:[-_\s]\d+|\(\d+\)|（\d+）)$/, '').trim();
                        if (!seenExtPatrolBaseNames.has(baseName)) {
                          seenExtPatrolBaseNames.add(baseName);
                          etlStatus = "已納入";
                        } else {
                          etlStatus = "未納入"; // duplicate sheet in same file
                        }
                      }
                    }
                  }
                }
              }
            } else {
              etlStatus = "狀態異常";
            }
          } else {
            // General QC files ETL validation logic
            const pathParts = filePath.split(/[\\/]/);
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

            if (!initialQC) {
              etlStatus = "未納入"; // Not in recognized QC folder
            } else {
              const relPath = pathParts.slice(folderIdx + 1, pathParts.length - 1).join('/');
              const fileName = file.name;
              
              const isSkipSheet = (sheetName === 'DATE' || sheetName === '空白' || sheetName === '範例' || sheetName === '客戶別' || sheetName.indexOf('Sheet') >= 0 || /^QC[-_]?\d+/i.test(sheetName.trim()) || sheetName.trim().indexOf('出貨') === 0);
              const isBlankFile = fileName.indexOf('空白') >= 0;
              
              if (isSkipSheet || isBlankFile) {
                etlStatus = "未納入";
              } else {
                const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                let actualQC = determineQCFromSheet(json, initialQC, relPath);
                let subCat = null;
                let month = null;

                // Overrides and custom logic
                const isSemiFinishedTable = /半成品品檢表/i.test(fileName) || /半成品品檢表/i.test(relPath || '');
                if (isSemiFinishedTable) {
                  actualQC = 'QC10006-R02';
                  subCat = '裝配C';
                  const sheetMatch = sheetName.match(/(\d{2})([A-L])/i);
                  if (sheetMatch) {
                    const yr = parseInt(sheetMatch[1], 10);
                    if (yr >= 15 && yr <= 99) {
                      month = LETTER_MONTH[sheetMatch[2].toUpperCase()] || null;
                    }
                  }
                  if (!month) {
                    const fnMonthMatch = fileName.match(/半成品品檢表\d{4}[-_](\d{1,2})/i);
                    if (fnMonthMatch) month = parseInt(fnMonthMatch[1], 10);
                  }
                  if (!month) {
                    const sheetYymmdd = sheetName.match(/(\d{2})(\d{2})(\d{2})/);
                    if (sheetYymmdd) {
                      const m = parseInt(sheetYymmdd[2], 10);
                      if (m >= 1 && m <= 12) month = m;
                    }
                  }
                } else if (initialQC === 'QC10007-R03') {
                  actualQC = 'QC10007-R03';
                  if (relPath && relPath.indexOf('Tubing') < 0) {
                    const letterMatch = fileName.match(/[-_]?([A-L])\.xlsx$/i);
                    if (letterMatch) {
                      const derivedMonth = LETTER_MONTH[letterMatch[1].toUpperCase()] || null;
                      const fileDate = findDateInSheet(ws, actualQC);
                      if (fileDate && fileDate.month === derivedMonth) {
                        month = derivedMonth;
                      }
                    }
                  }
                  if (relPath && relPath.indexOf('射出D') >= 0 && relPath.indexOf('射出D(組件)') < 0) {
                    actualQC = 'QC10002-R02';
                  }

                  // Blank template check
                  if (actualQC === 'QC10007-R03' && json && json.length > 3) {
                    const _lotRow = json[3];
                    const _lotVal = (_lotRow && _lotRow.length > 6) ? _lotRow[6] : '';
                    const _lotIsBlank = (_lotVal === '' || _lotVal === null || _lotVal === undefined ||
                                      _lotVal === 0 || String(_lotVal).trim() === '' || String(_lotVal).trim() === '0');
                    if (_lotIsBlank) {
                      etlStatus = "未納入"; // blank sheet skipped
                    }
                  }
                }

                if (etlStatus !== "未納入") {
                  subCat = getRawSubCategory(actualQC, relPath, fileName, sheetName, qcFolder);
                  month = extractRawMonth(ws, fileName, sheetName, year, relPath, json, actualQC);

                  if (actualQC === 'QC10007-R01') {
                    const baseName = sheetName.replace(/\s*\([^)]+\)\s*$/, '').trim();
                    if (seenQC7R1BaseNames.has(baseName)) {
                      etlStatus = "未納入"; // duplicate sheet
                    } else {
                      seenQC7R1BaseNames.add(baseName);
                    }
                  }

                  if (etlStatus !== "未納入") {
                    if (!actualQC || !subCat || !month || month < 1 || month > 12) {
                      etlStatus = "狀態異常";
                    } else {
                      etlStatus = "已納入";
                    }
                  }
                }
              }
            }
          }

          results.push({
            fileName: file.name,
            sheetName: sheetName,
            foundCode: foundCode || "無",
            foundName: foundCode ? foundName : "無",
            status: foundCode ? (mappings[foundCode] ? "matched" : "unmatched") : "none",
            etlStatus: etlStatus,
            etlTimestamp: etlTimestamp
          });
        });
        
        resolve({ success: true, fileName: file.name, sheets: results });
      } catch (err) {
        console.error(`Error parsing ${file.name}:`, err);
        resolve({
          success: false,
          fileName: file.name,
          error: "無法開啟 (可能損壞或格式不支援)"
        });
      }
    };
    
    reader.onerror = () => {
      resolve({
        success: false,
        fileName: file.name,
        error: "讀取檔案時發生錯誤"
      });
    };
    
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Exports the extraction results to a downloadable Excel file.
 * @param {Array} data - The scanned rows.
 * @param {string} folderName - The sheet and file name.
 */
export const exportToExcel = (data, folderName = "品管報表統計") => {
  const wsData = [
    ["檔案名稱", "工作表名稱", "表單編碼", "表單名稱"]
  ];
  
  data.forEach((row) => {
    wsData.push([
      row.fileName,
      row.sheetName,
      row.foundCode,
      row.foundName
    ]);
  });
  
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  
  // Safe sheet name (length <= 31, no invalid chars)
  let safeName = folderName
    .replace(/[:\\/?*\[\]]/g, "_")
    .substring(0, 31);
  if (!safeName) safeName = "Sheet1";
  
  // Fit column widths
  const maxLens = [10, 10, 10, 10];
  wsData.forEach(row => {
    row.forEach((val, i) => {
      const len = val ? String(val).length : 0;
      if (len > maxLens[i]) maxLens[i] = len;
    });
  });
  ws['!cols'] = maxLens.map(len => ({ wch: Math.min(len * 2 + 2, 45) }));
  
  XLSX.utils.book_append_sheet(wb, ws, safeName);
  XLSX.writeFile(wb, `${safeName}_工作表提取結果.xlsx`);
};

/**
 * Parses the summary Excel file dynamically in the client browser.
 * @param {File} file - The Excel file containing the summary sheets.
 * @returns {Promise<Object>} Object mapping sheet names to 2D arrays.
 */
export const parseSummaryExcel = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetData = {};
        
        workbook.SheetNames.forEach((sheetName) => {
          const ws = workbook.Sheets[sheetName];
          if (ws) {
            sheetData[sheetName] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          }
        });
        
        resolve(sheetData);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};
