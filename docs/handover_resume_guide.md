# 交接與重啟指南 (Handover Resume Guide)

## 當前專案狀態
- **專案名稱**：FileName_WorkSheet_Extract
- **最新完成進度**：已完成「全動態欄位架構」、「精準度與效能優化（單次讀取與 100 行限制）」以及「Chrome 本地資料夾儲存與快取優化」。
- **代碼庫狀態**：功能已驗證並整理完畢。

## 核心架構變更摘要
1. **動態欄位解析 (Dynamic Columns)**：
   - 移除了原先寫死 (hardcoded) 的欄位定義。
   - `getRawSubCategory` 自動動態解析並累積所有出現過的欄位，不受限於預定義。
2. **單次讀取效能重構 (Single-Read Optimization)**：
   - 移除雙重 `XLSX.read` 設計，改為單次 `sheetRows: 100` 載入，減少約 50% 的 ZIP 解碼與 XML 解析 CPU 開銷。
3. **UUID 路徑防禦**：
   - 內建 `isUUID` 路徑過濾器，自動在 `browserETL.js` 與 `excelParser.js` 剔除 Chrome 等瀏覽器臨時虛擬化產生的 UUID 資料夾層級，防禦資料污染。
4. **ETL 計算結果緩存 (Caching)**：
   - 在 `App.jsx` 引入 `cachedCounts` 狀態，使重複點擊輸出或個別報表時能以 `0ms` 速度即時下載，免除重複掃描時間。
5. **Chrome 資料夾原生寫入**：
   - 重構 `promptExportDirectory` 使用 `showDirectoryPicker()`，解決了 Chrome 安全阻斷帶斜線下載屬性的問題，實現多個 Excel 檔案直接原生寫入本地資料夾。

## 下一步開發建議 (Next Steps)
1. **GitHub 合併**：在遠端倉庫將 `feature/dynamic-columns` 合併 (PR/Merge) 至 `main` 分支。
2. **UAT 測試**：使用其他年份資料庫進行大規模回歸測試 (Regression Test)，確保沒有其他極端格式的表單被遺漏。
