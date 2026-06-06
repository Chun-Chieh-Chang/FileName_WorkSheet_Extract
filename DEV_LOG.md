# DEV_LOG (開發日誌) - 品管報表統計與提取工具

## 2026-06-06 啟動全新全網頁品管統計工具開發 (Vite + React)

### Motivation:
- 使用者希望將原有的 VBA 巨集（`檔名與工作表提取.txt`）重構並優化為全功能集成於單一操作頁面的網頁應用，並使用**淺色背景 (Light Mode)** 作為主要 UI 風格。
- 解決 Excel 巨集相依性高、表單編碼對照表維護不易（原為硬編碼於代碼內）、無統計圖表看板等問題。

### Diagnosis:
- VBA 機制使用 Excel 內建 UsedRange 搭配 regex 提取 `QCxxxxx-Rxx` 編碼。
- 網頁應用中可使用 `xlsx` (SheetJS) 庫在瀏覽器端完全離線解析 Excel 檔案，讀取其 Sheets 的儲存格內容，並套用對等正則表達式，完全無須後端伺服器，資料不落地保障安全性。

### Implementation Plan:
1. **初始化專案**：使用 `npx create-vite` 初始化 React 專案，清理無用資源 (MECE)。
2. **安裝核心相依性**：安裝 `xlsx` 庫進行 Excel 格式檔案解析。
3. **實作核心邏輯**：
   - `src/utils/db.js`：封裝 `localStorage`，使表單對照表可由使用者直接在網頁上動態管理，排除原先 VBA 硬編碼之弊端。
   - `src/utils/excelParser.js`：實作 `FileReader` 與 SheetJS 儲存格掃描，能快速提取 QC 編碼，並提供 `exportToExcel` 輸出統計結果表格。
4. **UI 設計**：
   - 遵從色彩大師規範 (Light Mode)：底色 `#F9FAFB`，卡片表面色 `#FFFFFF`，標題 `#111827`，CTA 按鈕 `#3B82F6`。
   - 設計為一體化雙欄式面板：
     - 左欄：資料導入（支援資料夾與多檔案選取）、自訂對照表編輯器（支援新增、刪除、匯入/匯出 JSON 備份、一鍵恢復系統預設值）。
     - 右欄：Dashboard 統計圖表看板（總檔案、總工作表、已成功識別表單數、識別覆蓋率）、提取結果互動數據表（支援搜尋過濾、狀態篩選、一鍵導出 Excel）。

---

### Status:
- [x] 專案初始化與結構清理 (Vite-React)
- [x] 資料庫管理模組開發 (`src/utils/db.js`)
- [x] Excel 解析與導出模組開發 (`src/utils/excelParser.js`)
- [x] 淺色背景 Design Tokens 與 CSS 排版開發 (`src/index.css`)
- [x] 主程式與介面整合開發 (`src/App.jsx`)
- [x] 確效與安裝 `xlsx` 函式庫
- [x] 修復對照表編碼欄位文字折行問題：設定 `white-space: nowrap` 與調整彈性佈局寬度及垂直居中。
- [x] 新增頁尾作者資訊：「Developed by Wesley Chang @Mouldex, 2026.」


