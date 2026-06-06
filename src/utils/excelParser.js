import * as XLSX from 'xlsx';

/**
 * Parses a single Excel file, scanning all sheets for QC codes.
 * @param {File} file - The file object to parse.
 * @param {Object} mappings - The QC code mappings map.
 * @returns {Promise<Object>} Resolves with the parsing results.
 */
export const parseExcelFile = (file, mappings) => {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const results = [];
        
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
          
          results.push({
            fileName: file.name,
            sheetName: sheetName,
            foundCode: foundCode || "無",
            foundName: foundCode ? foundName : "無",
            status: foundCode ? (mappings[foundCode] ? "matched" : "unmatched") : "none"
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
