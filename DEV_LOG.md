# 開發日誌 (DEV_LOG.md)

## 2026-06-28 Excel 報表結構與欄位重構

### 需求說明
1. 全自動 SkillsBuilder 開發模式啟動。
2. 解析 `etl_pipeline.cjs` 的資料提取與輸出邏輯，以 `2025品檢報表統計(原始標準檔).xlsx` 為樣板進行比對，尋找邏輯與結構缺失。
3. 修正 `裝配對樣巡檢(QC10006-R01)` 移除 `"小計"`。
4. 修正 `半成品品檢(QC10006-R02)` 移除 `"裝配A"`, `"裝配B"`, `"其他"`，因為 `"裝配A"`, `"裝配B"` 應歸類在 `零組件入庫品檢(QC10007-R03)`（已在 pipeline 中按表單編號正確歸類，僅需修正產出 Excel 及報表解析的欄位）。
5. 忽略 NCA 工作表的內容與邏輯。
6. 還原 `彙總表` 的 3x3 並排網格排版格式。
7. 補回各工作表首行 (Row 0) 的大標題列，避免首行偏移，並將資料讀取範圍從 Row 1-12 修正為 Row 2-13。

### 遇到的問題與根因分析 (RCA)
- **問題一：生成的 Excel 工作表佈局與標準樣板不符**
  - *原因*：之前的 `etl_pipeline.cjs` 是將所有子表在 `彙總表` 工作表中以垂直堆疊的方式寫入，且忽略了每個子表的首行大標題列。
- **問題二：半成品及裝配巡檢工作表欄位與樣板有出入**
  - *原因*：之前的代碼中硬編碼了佔位欄位（如裝配A/B），但其實這些應完全歸入零組件入庫工作表。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 重構 `writeSummaryExcel` 中 `彙總表` 的寫入邏輯，使用 2D 矩陣按坐標寫入各區域子表。
  2. 在所有工作表寫入前先添加大標題列，順延寫入欄位與資料。
  3. 修改資料解析端（年報與 Stripe 儀表板 HTML 產生器），將解析資料行數的索引調整為從 index 2 開始，並修正半成品的欄位列。
- **預防措施**：
  - 未來新增或修改品檢類別時，須優先對照 `原始標準檔` 檢查欄位結構。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立。
- [x] 程式碼修改與測試。

---

## 2026-06-28 麥肯錫風格 (McKinsey Style) 互動式分析儀表板實作

### 需求說明
1. 基於已產出的品檢報表數據，動態生成自訂統計圖表。
2. 圖表以月份為 X 軸（1月-12月），數量為 Y 軸。
3. 用戶可以任意勾選想要放進圖表中的細項欄位，各工作表獨立區分。
4. UI 介面採用麥肯錫 (McKinsey) 顧問報告風格（深海軍藍與金色主色調、襯線標題、淺色乾淨背景與極簡網格）。
5. 後續需求：柱狀圖改為堆疊顯示；新增一鍵清空核取按鈕。

### 遇到的問題與根因分析 (RCA)
- **問題一：Vite 本地 Base Path 導致自動載入報表失敗**
  - *原因*：`fetch('/DataExtract/2025品檢報表統計.xlsx')` 中的路徑被解析為網域根路徑，忽略了專案在 Vite 中配置的 `/FileName_WorkSheet_Extract/` 基底路徑。
- **問題二：圖表動態數據更新與記憶體洩漏**
  - *原因*：如果不在組件銷毀或工作表切換時調用 `chartInstance.destroy()`，Chart.js 會在同一個 Canvas 上重複繪製多個實例，造成重疊、閃爍與記憶體洩漏。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 使用 `import.meta.env.BASE_URL` 動態拼接自動載入的路徑，確保不論在本機開發或正式部署時都能正確獲取報表檔案。
  2. 在 React 的 `useEffect` 中添加健全的生命週期管理，每次重繪圖表前強制執行 `chartInstance.current.destroy()` 清除舊實例。
  3. 全面使用 light-theme 淺色背景（如 `#F4F6F9` 和 `#FFFFFF`），搭配襯線字體 Georgia 建立麥肯錫風格視覺。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立。
- [x] 程式碼修改與測試。
- [x] 麥肯錫互動儀表板實作與測試。

---

## 2026-06-28 原物料品檢(QC10002-R02) 欄位調整對應

### 需求說明
用戶微調 `2025品檢報表統計(原始標準檔).xlsx` 中「原物料品檢」工作表的欄位結構，要求 ETL 產出的統計欄位完全對應新的標準格式。

### 工作表結構（最新版）
- 舊版：左右雙表並列（「主料分類」+ 「子項物料分類」）
- 新版：單一主表 14 欄
  - `["月份","原料","B膠","收縮膜","色粉","空白包裝袋","空白感壓紙","塑膠袋","塑膠袋40*50","紙箱","過濾網連蓋","標籤","射出D","小計"]`

### 矯正措施
1. 重構 `writeSummaryExcel` 中原物料工作表的寫入邏輯，改為單一 14 欄表格。
2. 更新 `generateStyledReport` 的解析邏輯，配合新欄位索引讀取各項目數據。
3. 新的小計欄位（col 13）由代碼自動加總，不依賴從 Excel 讀回。
4. `"塑膠袋40*50"` 欄位同時接受 `40X50` 或 `40*50` 來源命名格式。

### 進度追蹤
- [x] `etl_pipeline.cjs` — 重構原物料工作表寫入邏輯。
- [x] `generate_styled_reports.cjs` — 更新解析索引以對應新欄位順序。
- [x] 重新執行 ETL 並驗證 2025/2026 輸出正確。
- [x] 將更新後的 summary Excel 部署至 `public/DataExtract/`。

---

## 2026-06-28 MECE 代碼清理 & 文檔更新

### 需求說明
全自動開發模式：識別過時、冗餘或無效的代碼與檔案，執行 MECE 整合整理，更新開發文檔，建立還原基準點並推送至 GitHub。

### 識別到的問題（Dead Code & Redundancy）

| 項目 | 類型 | 說明 |
|---|---|---|
| `etl_pipeline.cjs` L1083-1409 | 死代碼 | Section 4「HTML Report Generator」共 327 行，含 `generateYearlyReport`、`generateComparisonReport`、`renderSlidesToHTML`、`renderComparisonPage` 等函數 — 已被 `generate_styled_reports.cjs` 完整替代 |
| `etl_pipeline.cjs` `COLORS` 常數 | 死代碼 | 僅在已移除的 Section 4 中使用 |
| `etl_pipeline.cjs` `MONTH_NAMES` 常數 | 死代碼 | 僅在已移除的 Section 4 中使用 |
| `etl_pipeline.cjs` `generateReport()` 呼叫 | 過時呼叫 | 在 `main()` 中透過 `isComparison=true` 呼叫已移除的函數，改為內聯簡化版 HTML |
| `report-templates/` 目錄 | 封存原型 | 設計原型階段，目前功能由 `generate_styled_reports.cjs` 實現，已無任何程式碼參照 |
| `generate_styled_reports.cjs` 文件注釋 | 過時描述 | docstring 中仍保留 "matching report-templates/ style" 描述 |
| `README.md` | 過時 | 未反映當前架構（儀表板、ETL 流程圖、單一統合原物料表格） |

### 矯正措施 (CAPA)
1. **外科手術式移除**：刪除 `etl_pipeline.cjs` 第 1083-1409 行（Section 4 全部）及 `COLORS`、`MONTH_NAMES` 兩個常數。
2. **呼叫替換**：將 `main()` 中的 `generateReport(isComparison=true)` 呼叫替換為內聯 HTML 字串寫入，保留相同功能。
3. **文檔清理**：更新 `generate_styled_reports.cjs` docstring、`README.md`、`DEV_LOG.md`。
4. `report-templates/` 目錄保留但已在 README 中標記為「封存原型，不再使用」。
5. **確效驗證**：執行 `node etl_pipeline.cjs all` 確認輸出完全一致，無任何錯誤。

### 清理成果
- `etl_pipeline.cjs`：由 1,467 行減少至 ~1,140 行（減少 ~21%，移除 327 行死代碼）
- 無任何功能性迴歸
- README、DEV_LOG 完整更新

### 進度追蹤
- [x] 識別死代碼與過時文件。
- [x] 移除 `etl_pipeline.cjs` Section 4 死代碼。
- [x] 修復 `main()` 中的過時函數呼叫。
- [x] 更新 `generate_styled_reports.cjs` docstring。
- [x] 重構 `README.md`。
- [x] 更新 `DEV_LOG.md`。
- [x] 驗證 ETL 管線無迴歸。
- [x] git commit & push。
