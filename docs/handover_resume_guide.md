# 交接與重啟指南 (Handover Resume Guide)

## 當前專案狀態
- **專案名稱**：FileName_WorkSheet_Extract
- **分支**：`feature/dynamic-columns`（已合併至 `main`）
- **部署**：GitHub Pages（透過 GitHub Actions 自動部署）
- **最新完成進度**：民國年日期支援、Chrome UUID 防禦、半成品雙重驗證

## 核心架構摘要

### 1. 雙次選讀效能架構 (Double-Read Selective Parsing)
- 第一次 `XLSX.read` 使用 `bookSheets: true` 僅讀取工作表目錄
- 過濾出目標工作表後，第二次使用 `sheets: targetSheets` + `sheetRows: 100` 精準解析
- 避免解壓縮與解析非目標工作表的 XML，大幅提升效能

### 2. 動態欄位解析 (Dynamic Columns)
- `getRawSubCategory` 自動動態解析所有出現的品檢子類別
- 新品項無須手動新增至代碼，自動展開為 Excel 欄位

### 3. UUID 路徑防禦
- `isUUID` 過濾器支援標準 UUID、32 位十六進位、及長度 ≥ 24 的系統生成隨機字串
- 應用於 `browserETL.js`、`excelParser.js`、`App.jsx` 三處

### 4. 民國年 (ROC Year) 日期支援
- `parseDateFromString` 與 `findDateInSheetFallback` 支援三位數民國年（如 `112/03/15`）
- 支援 `.`（點號）作為日期分隔符

### 5. 半成品品檢表雙重驗證
- 要求工作表內容的 QC 編碼為 `QC10006-R02` **且** 檔名/路徑含 `半成品品檢表`
- 不再需要硬編碼黑名單來排除特定資料夾

### 6. QC 編碼混合式掃描
- 第一階段：Column A 極速掃描（最多 100 行）
- 第二階段回退：前 15 行的 Columns B-H 掃描（舊版 2023 表格相容）

### 7. ETL 計算結果緩存
- `cachedCounts` 狀態避免重複掃描，支援即時重複匯出

## 關鍵檔案索引

| 檔案 | 職責 |
|---|---|
| `src/App.jsx` | 主介面、快取、UI 控制 |
| `src/utils/browserETL.js` | ETL 核心引擎 |
| `src/utils/excelParser.js` | Tab 2 數據表的 Excel 解析器 |
| `src/utils/db.js` | QC 編碼對照表持久化 |
| `scratch/validate_qc_etl.cjs` | 自動確效測試工具 |
| `DEV_LOG.md` | 開發日誌（含 RCA + CAPA） |

## 開發與部署流程

```bash
# 本地開發
npm install && npm run dev

# 生產建構
npm run build

# 確效測試
node scratch/validate_qc_etl.cjs

# 推送部署（GitHub Actions 自動觸發 GitHub Pages）
git push origin main
```
