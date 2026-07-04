# 開發日誌 (DEV_LOG.md)

## 2026-07-04 專案整體優化與 MECE 清理

### 清理作業
- **刪除**：`html_backup_20260704/` 整個備份目錄（功能已整合到主專案）
- **刪除**：5 個臨時 CSV 檔案（`2026品檢報表統計_field_mapping_*.csv`、`進料檢驗-2026_field_mapping_*.csv`）
- **刪除**：`extract_field_mapping.cjs`（已被 `extract_raw_field_mapping.cjs` 取代）
- **刪除**：`data_logic_spec.html`、`品檢數據鏈路圖.html`（與備份重複）
- **刪除**：`scratch/` 目錄（清理腳本已無必要）

### 功能新增
- **欄位映射匯出功能**：在「品管檔案夾掃描」區域新增「📋 匯出欄位映射」按鈕
  - 支援從選取的 Excel 檔案中提取每個工作表的欄位名稱
  - 輸出格式：`品管標籤編號,欄位名稱,資料路徑`
  - 自動推斷每個工作表對應的 QC 標籤編號
  - 使用 `Promise.all` 確保所有檔案讀取完成後再生成 CSV
  - 修復 `FileReader.readAsBinaryString` 與 `XLSX.read` 的相容性问题

### 邏輯優化
- **QC10002-R02 日期檢查範圍擴大**：從 N4/O4 改為檢查 **N4:P5** 範圍內的所有 6 個欄位
- **日期解析強化**：新增月份和日期範圍驗證，確保只有有效日期才被識別
- **工作表過濾規則簡化**：統一為「排除 QC 開頭、Sheet、空白的工作表」

### 按鈕顏色統一 (MECE)
| 按鈕類型 | 用途 | 樣式 |
|---------|------|------|
| btn-primary | 主要操作、匯出功能 | 深藍背景，白色文字 |
| btn-secondary | 次要操作 | 白色背景，深色邊框 |
| btn-danger | 刪除/清除 | 紅色 |
| btn-success | 成功/完成 | 綠色 |
| btn-info | 資訊/查看 | 淺藍色 |
| btn-warning | 警告/注意 | 黃色 |
| btn-dark | 中性操作 | 灰色 |

### 修改檔案
| 檔案 | 變更 |
|------|------|
| `src/App.jsx` | 新增 `exportFieldMapping` 函數；統一按鈕顏色；修復異步檔案讀取 |
| `etl_pipeline.cjs` | 擴大 QC10002-R02 日期檢查範圍；強化日期解析邏輯 |
| `src/utils/browserETL.js` | 同步 etl_pipeline.cjs 的日期檢查邏輯 |
| `extract_raw_field_mapping.cjs` | 新增 - 原始資料欄位映射提取工具 |

---

## 2026-07-03 ETL Bug 修復、UI 優化與 MECE 專案清理

### Bug 修復：browserETL.js 空白檔案過濾邏輯錯誤
- **問題**：`runETLInBrowser` 中遇到檔名含「空白」的檔案時使用 `return` 而非 `continue`，導致整個 ETL 迴圈提前終止，後續檔案未被處理。
- **RCA**：`if (fileName.indexOf('空白') >= 0) return;` 應為 `continue`。
- **CAPA**：修改為 `continue`，確保跳過範本檔案後繼續處理剩餘檔案。

### UI 優化：表格容器與置中佈局
- **問題**：狀態欄位文字縱向排列、左右容器留白不平衡。
- **CAPA**：
  1. 所有表格欄位添加 `white-space: nowrap` 防止文字換行。
  2. 各欄位設置 `min-width`（序號 60px、檔案名稱 200px、工作表名稱 150px、表單編碼 120px、表單對照名稱 180px、狀態 120px）。
  3. `.main-grid` 添加 `max-width: 1600px; margin: 0 auto` 確保整體置中。
  4. 移除右側面板固定 `minHeight`，改為 `flex: 1` 使兩欄等高對齊。
  5. 移除 `.table-wrapper` 的 `min-width` 限制，讓表格隨內容自适应。

### UI 優化：報表清空功能
- **需求**：自動載入的 2025/2026 統計檔無法清空，用戶要求可手動刪除。
- **CAPA**：在 Dashboard 頁面添加「🗑 清空所有報表」按鈕，點擊後清除 `summaryFiles`、`summaryData`、`activeSheet`、`selectedItems`。

### MECE 檔案清理
- **刪除**：`task.md`（空任務清單）、`progress.md`（本次優化用）、`task_plan.md`（本次優化用）。
- **保留**：`scratch/cleanup.cjs`（生產工具）、`scratch/test_semi_finished_month.cjs`（驗證腳本）、`generate_styled_reports.cjs`（HTML 報表生成器）、`品檢數據鏈路圖.html`（視覺化文檔）、`data_logic_spec.html`（技術規格）。

---

## 2026-07-01 跨年度對比、月份篩選、獨立報表、Letter 後綴、空白過濾

### 功能新增
1. **跨年度對比模式**：勾選多個年份進行數據對比，圖表 X 軸顯示年份。
2. **月份篩選器**：下拉選單篩選特定月份，圖表/KPI/表格即時更新。
3. **獨立報表匯出**：一次產出 16 個獨立 QC 表單 Excel 檔案。
4. **年份後綴兼容**：`getRawSubCategory` 自動剝除年份後綴（如 `原料-2020` → `原料`）。
5. **Letter 後綴月份提取 (Strategy 9)**：匹配 `YEAR+LETTER` 格式（如 `裝配C-2021A.xlsx` → A=1月）。
6. **空白檔案/工作表過濾**：檔名含「空白」跳過；工作表含「Sheet」或「工作表」跳過。
7. **多檔案匯入**：報表上傳支援 `multiple` 屬性。
8. **工作表內容日期掃描擴展**：`findDateInSheetFallback` 從 10×5 擴充至 20×全部欄位。

### 修改檔案
| 檔案 | 變更 |
|------|------|
| `src/App.jsx` | 新增 `selectedMonth`、`compareYearSelection` state；重構圖表邏輯；新增獨立報表匯出 |
| `etl_pipeline.cjs` | Strategy 9、年份後綴剝除、`findDateInSheetFallback` 擴展、空白過濾 |
| `src/utils/browserETL.js` | 同步以上所有修改 |
| `data_logic_spec.html` | 新增 Sections 8-10 |
| `品檢數據鏈路圖.html` | 更新版本號 |

---

## 2026-07-01 側邊欄寬度調整、字母月份驗證與 scratch 清理

### 問題修復
1. **側邊欄寬度**：`.main-grid` 從 380px → 480px，確保表單編碼完整顯示。
2. **字母月份驗證**：新增 `findDateInSheet()` 交叉比對，檔名字母 A-L 必須與檔案內容日期一致才接受。
3. **scratch 清理**：刪除 9 個調試腳本，保留 `cleanup.cjs` 與 `test_semi_finished_month.cjs`。

---

## 2026-07-01 專案整理：保留數據鏈路圖、MECE 整合

- `品檢數據鏈路圖.html` 加入 git 追蹤。
- 確認所有現有檔案均有明確用途，無冗餘。

---

## 2026-07-01 一鍵生成器動態年份與自動偵測優化

### 功能優化
1. **動態年份匹配**：移除寫死年份 Regex，依據傳入 `year` 參數動態建置。
2. **空白樣板過濾**：前端 ETL 同步 `QC10007-R03` `_lotIsBlank` 判斷守衛。
3. **UI 優化**：新增 `etlYear` state，支援 2010-2040 年動態選擇。
4. **併發狀態保護**：`isScanning` 時禁用下拉選單與按鈕。
5. **前後端日期提取對齊**：`browserETL.js` 單元格指紋與後端 `etl_pipeline.cjs` 完全同步。

---

## 2026-07-01 2026 年 7-12 月幻象數據問題修復

### RCA
- `零組件入庫-2026/裝配C-2026/` 預建整年 G-L 月份空白樣板（批號=0），被誤計為有效記錄。
- `Vivus-20260703.xlsx`（7/3 真實批次）與 `ICU-260904.xlsx`（9/4 預排出貨）為真實資料，保留。

### CAPA
- 新增 `_lotIsBlank` 空白樣板守衛：`QC10007-R03` 批號為空或 0 時跳過。
- 2026 Grand Total 從 4421 → 4394（剔除 12 筆幻象記錄）。

---

## 2026-06-30 啟動 SkillsBuilder 模式 UI/UX 設計與字型優化

### 視覺重塑
1. **中英文字型整合**：`Outfit`（英文）+ `Noto Sans TC`（繁體中文）。
2. **McKinsey 色彩**：`--mck-navy: #0A2540`、`--mck-accent-gold: #C5A059`。
3. **毛玻璃控制台**：`.app-nav` 添加 `backdrop-filter: blur(12px)`。
4. **微動畫**：按鈕 hover `translateY(-1px)`、點擊 `scale(0.98)`。

---

## 2026-06-30 冗餘檔案清理與 MECE 專案重構

### 刪除項目
- `report-templates/`：早期 HTML 設計原型。
- `compare_b_glue.ipynb`、`read_excel.html`：早期測試檔案。
- `DataExtract/2025`、`DataExtract/2026`：早期管線中間產物。

---

## 2026-06-30 新增一鍵全選功能

- 於 `App.jsx` 添加「✅ 一鍵全選」按鈕，支援狀態感知（全選時 disabled）。

---

## 2026-06-30 QIP 射出製程品檢資料提取邏輯重構

### 規則變更
1. 射出 Setup：`RawData/{year}/射出檢驗-{year}/QIP-{year}(1~10)` 資料夾內檔案數。
2. 射出巡檢：僅該子資料夾內，經 Date Code 格式驗證（`^\d{6}[a-zA-Z]?$`）與後綴去重後的工作表數。

---

## 2026-06-30 移除多餘工作表 (彙總表、NCA、品檢地圖)

1. `etl_pipeline.cjs` `writeSummaryExcel` 移除三種工作表生成邏輯。
2. `scratch/delete_sheets.cjs` 清理現有檔案中的多餘工作表。
3. `determineQCFromSheet` 加入 `QC10007-R02` 與 `QC10007` 識別，正確歸入完成品品檢。

---

## 2026-06-30 零組件入庫 (QC10007-R03) 提取與月份邏輯調整

1. 強制 `QC10007-R03` 歸類，`射出D` 重定向至 `QC10002-R02`。
2. 支援 A-L 字母月份後綴（需緊鄰年份，防止誤判 Tubing 的 D 班別）。
3. Tubing 檔案以父資料夾名稱直接對應月份。

---

## 2026-06-30 重新生成品檢報表統計

- 2025 年：12,269 筆記錄。
- 2026 年：4,421 筆記錄。

---

## 2026-06-29 B膠月份與數量提取邏輯對齊

### 問題根因
1. B膠工作表名稱（批號）開頭為歷史日期，導致月份誤配。
2. 日期位於 Column N，舊版僅掃描前 5 欄。
3. Rev C (`N4`) 與 Rev D (`O4`) 版面偏移未適應。

### CAPA
- 水平展開所有 QC 表單日期單元格映射。
- `QC10002-R02` 優先讀 `N4`，失敗讀 `O4`。
- 解析失敗/無日期工作表改為跳過（非預設 1 月）。
- 過濾 `QC[-_]?\d+` 開頭的空白樣板工作表。

---

## 2026-06-28 Excel 報表結構與欄位重構

1. `彙總表` 改為 3x3 並排網格排版。
2. 所有工作表補回首行大標題列。
3. 資料讀取範圍從 Row 1-12 修正為 Row 2-13。
4. 移除半成品中的「裝配A/B/其他」（歸入零組件入庫）。
5. 移除裝配巡檢中的「小計」。

---

## 2026-06-28 麥肯錫風格互動式分析儀表板實作

1. React SPA 動態生成自訂統計圖表。
2. 用戶可勾選細項欄位，以月份為 X 軸生成堆疊柱狀圖。
3. 修復 Vite Base Path 自動載入失敗問題（使用 `import.meta.env.BASE_URL`）。
4. Chart.js 實例生命週期管理（`destroy()` 防止記憶體洩漏）。

---

## 2026-06-28 原物料品檢欄位調整

- 改為單一 14 欄表格：`["月份","原料","B膠","收縮膜","色粉","空白包裝袋","空白感壓紙","塑膠袋","塑膠袋40*50","紙箱","過濾網連蓋","標籤","射出D","小計"]`。
- `"塑膠袋40*50"` 同時接受 `40X50` 或 `40*50` 來源。

---

## 2026-06-28 MECE 代碼清理

- 移除 `etl_pipeline.cjs` Section 4 死代碼（327 行），由 1,467 行減至 ~1,140 行。
- 刪除 `COLORS`、`MONTH_NAMES` 常數（僅在已移除的 Section 4 中使用）。
- `report-templates/` 保留但標記為封存原型。
