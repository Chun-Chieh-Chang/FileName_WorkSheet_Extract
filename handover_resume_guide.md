# 交接與重啟指南 (Handover Resume Guide)

## 當前專案狀態
- **專案名稱**：FileName_WorkSheet_Extract
- **最新完成進度**：已完成「全動態欄位架構」與「極限效能優化 (A51 深水區掃描)」。
- **代碼庫狀態**：功能已驗證並整理完畢，準備合併至主線。所有冗餘的暫存腳本已清除。

## 核心架構變更摘要
1. **動態欄位解析 (Dynamic Columns)**：
   - 移除了原先寫死 (hardcoded) 的陣列定義。
   - `getRawSubCategory` 自動動態解析並累積所有出現過的欄位（如 Biometrix, Vivus 等），不受限於預定義。
2. **深度掃描效能優化 (BrowserETL & ExcelParser)**：
   - 移除 15 行掃描限制，擴展至 `json.length` (最高 100 行)，成功捕捉位於表單底部 (如 A51) 的 QC 編碼。
   - 保留 A 欄掃描限制，極大化 ETL 的處理效能與記憶體利用率。
   - 導入 `sheetRows: 100` 以提早阻斷底層 XLSX 解析器的過度讀取。

## 下一步開發建議 (Next Steps)
1. **GitHub 合併**：在遠端倉庫將 `feature/dynamic-columns` 合併 (Merge/PR) 至 `main` 分支。
2. **UAT 測試**：使用其他年份 (例如 2023, 2024, 2025) 的資料庫進行大規模回歸測試 (Regression Test)，確保沒有其他極端格式的表單被遺漏。
3. **擴充功能**：若未來有新的 QC 編碼加入，只需在 `src/utils/browserETL.js` 的 `FORM_TITLE_MAP` 補上標題映射，前端即可無縫支援。
