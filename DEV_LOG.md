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

---

## 2026-06-19 Codex ↔ Agnes AI 本地翻譯代理整合 (codex-proxy.js)

### Motivation:
- 使用者希望以 Codex (Claude Code CLI) 連接 Agnes AI Responses API，讓 IDE 工具能直接以 `agnes-2.0-flash` 模型運作。
- Agnes API 的結構嚴格遵守 OpenAI Responses API 規範，但 Codex 的封包格式（message role、content type 等）並不完全相容。

---

### 問題現象 (Symptom):
1. **Validation 錯誤爆炸**：每輪對話，Agnes 後端回傳大量 `Field required` / `Input should be valid string` 的 Pydantic 驗證錯誤。
2. **無限循環**：Codex 在多輪工具調用中，因歷程被截斷或 assistant 格式不符，導致模型重複觸發相同工具形成死循環。

---

### 根因分析 (RCA):
| # | 根因 | 說明 |
|---|------|------|
| 1 | **`role: developer` 不合規** | Codex 內部用 `developer` role 傳送 system prompt，Agnes API 僅接受 `system` / `user` / `assistant`。 |
| 2 | **assistant content block 格式錯誤** | Codex 以 `{type: "text"}` 回傳助理內容，Agnes 要求 assistant message 必須是完整的 `{type: "message", role: "assistant", status: "completed", content: [{type: "output_text", annotations: []}]}` 結構。 |
| 3 | **`reasoning` block 殘留** | 模型推理過程的 `reasoning_text` block 直接出現在歷程中，Agnes 不接受此類型。 |
| 4 | **tool_calls 格式差異** | Codex 用 OpenAI Chat Completions 格式（`msg.tool_calls[]`），Agnes 要求 `{type: "function_tool_call", id, call_id, name, arguments}` 格式。 |
| 5 | **tool_result 包在 user message 裡** | Claude-style 的 tool 回傳是包在 user message content 的 `{type: "tool_result"}` block 中，Agnes 要求獨立的 `{type: "function_call_output", call_id, output}` 頂層物件。 |
| 6 | **tools 定義格式差異** | Codex 可能傳入 `{type:"function", function:{name, description, parameters}}` 或 namespace 巢狀結構，Agnes 要求 flat 格式 `{type:"function", name, description, parameters}`。 |

---

### 矯正措施 (CAPA):
部署本地 **HTTP 翻譯代理** (`~/.cc-switch/codex-proxy.js`)：
- 監聽 `127.0.0.1:8000`，攔截所有 `POST /v1/responses` 請求。
- 執行 `mapInputMessages()` 函式，自動將上述所有不合規格式轉換為 Agnes 合規格式。
- 在 `~/.claude/settings.json` 設定 `ANTHROPIC_BASE_URL=http://127.0.0.1:8000/v1`，將 Codex 流量透明重導向。
- 設定 Windows 開機啟動腳本 (`Startup/codex-proxy.bat`) 確保代理常駐。

### 產出文件:
- **`codex_agnes_sop.html`**：含 Light/Dark 模式切換、一鍵複製 PowerShell 腳本、3 步驟圖文說明的完整 SOP 文件。
- **`~/.cc-switch/codex-proxy.js`**：已部署的本地翻譯代理伺服器原始碼。

### Status:
- [x] 根因分析完成 (6 項格式差異)
- [x] codex-proxy.js 翻譯邏輯撰寫與測試通過
- [x] codex_agnes_sop.html 設計與交付 (Premium UI + 深淺模式)
- [x] Windows Startup 開機啟動設定完成

---

## 2026-06-20 啟動 SkillsBuilder 全域開發模式

### Motivation:
- 將 SkillsBuilder 的核心能力整合到當前專案的 OpenCode 全域規則中
- 不需克隆完整倉庫，只需複製關鍵規則、技能與 IDE 配置文件

### Implementation:
1. **建立 `.opencode/` 目錄結構**：
   - `.opencode/rules/rules.md` — PDCA SOP、Color Master Palette、Superpowers Guardrails
   - `.opencode/plugins/superpowers.js` — OpenCode 配置注入插件
2. **安裝 34 個新技能** 到 `~/.agents/skills/`：
   - `deep-research`, `find-skills`, `office-processor`, `vetter` (core)
   - `agent-shield`, `autonomous-executor`, `autoresearch`, `bug-diagnose`, `code-reviewer`, `complexity-reduction`, `dispatching-parallel-agents`, `executing-plans`, `finishing-a-development-branch`, `github-manager`, `gitnexus`, `graphify`, `grill-requirements`, `hooks-enhancer`, `knowledge-bridge`, `loop-operator`, `requesting-code-review`, `receiving-code-review`, `session-memory`, `skill-architect`, `skills-builder`, `spec-architect`, `subagent-driven-development`, `tdd-enforcer`, `using-git-worktrees`, `verification-before-completion`, `web-coder`, `writing-plans`, `writing-skills`
   - `premium-design` (ui)
3. **部署 13 個 IDE 規則文件**：
   - `.claude-plugin/`, `.cursor-plugin/`, `.codex-plugin/`, `.kiro/`, `.continue/`, `.qoder/`, `.trae/`, `.vscode/`
   - `.antigravity.md`, `.clinerules`, `.cursorrules`, `.rules`, `.windsurfrules`
4. **複製 wiki/ 知識庫** — Karpathy LLM Wiki 模式
5. **複製 hooks/ 鉤子系统** — SessionStart 上下文注入

### Status:
- [x] 建立 .opencode/ 目錄結構
- [x] 寫入 rules.md (PDCA SOP + Color Master Palette + Superpowers Guardrails)
- [x] 複製 superpowers.js 插件
- [x] 安裝 34 個新技能
- [x] 部署 13 個 IDE 規則文件
- [x] 複製 wiki/ 知識庫
- [x] 複製 hooks/ 鉤子系统
- [x] 驗證 opencode.json 配置
- [x] 執行 verify.ps1 確效

---

## 2026-06-21 MECE 全面清理與還原基準線

### 動機：
- 歷經多輪提取腳本迭代（v4 → v13）、偵錯檢查腳本（50+ 個 .cjs 檔案）、QC 對應分析後，專案目錄充斥一次性偵錯腳本與過時版本。
- 需建立乾淨的可還原狀態（Git baseline）以區分生產代碼 vs 探索性腳本。

### 清理範疇（MECE 分類）：

| 類別 | 檔案數 | 內容 | 處置 |
|------|-------|------|------|
| 過往版本腳本 | 10 | `_v4.cjs` ~ `_v13.cjs` | 刪除（已被 `p2025_fixed.cjs` 取代） |
| 一次性驗證腳本 | 3 | `_verify.cjs` ~ `_verify3.cjs` | 刪除 |
| 檢查腳本 | 11 | `check_*.cjs`（QC 對應檢查） | 刪除（一次性分析） |
| 偵錯腳本 | 6 | `debug_*.cjs`（Tubing/QIP 偵錯） | 刪除 |
| 分析/比較腳本 | 13 | `analyze_*.cjs`, `compare_*.cjs`, `qc_analysis*.cjs` 等 | 刪除 |
| 過時 2026 腳本 | 4 | `p2026.cjs`, `p2026.mjs`, `process_2026.js` 等 | 刪除（僅保留 `p2026_fixed.cjs`） |
| 暫存/參考圖檔 | 3 | `test123.txt`, `ref_screenshot.png`, `reference_result.png` | 刪除 |
| AI IDE 配置目錄 | 13 個 | `.antigravity.md`, `.cursorrules`, `.kiro/`, `.opencode/` 等 | 移入 `.gitignore` |
| 原始資料目錄 | - | `RawData/`, `RawData_*.xlsx` | 移入 `.gitignore` |
| 生成輸出 | - | `DataExtract/`（舊位置）+ `DataExtract/2025/`（新位置） | 移入 `.gitignore` |

### 提取腳本最終狀態：

| 腳本 | 說明 | 用途 |
|------|------|------|
| `p2025_fixed.cjs` | 2025 年度資料提取主腳本 | ✅ 生產就緒，包含 date code 月份提取、裝配巡檢 per-sheet 計數、異常日期容錯 |
| `p2026_fixed.cjs` | 2026 年度資料提取腳本 | ✅ 同結構，待驗證對應 |
| `gen_summary.cjs` | 統計摘要產生器 | 保留（可能用於 2026） |
| `gen_script.cjs` | 腳本產生器 | 保留 |

### 對應文件狀態：

| 文件 | 狀態 | 說明 |
|------|------|------|
| `task_plan.md` | ✅ | 完整提取→統計表對應矩陣，含分組規則與特殊案例（雙重屬性等） |
| `DEV_LOG.md` | ✅ | 開發歷程記錄 |
| `implement_plan.md` | 保留 | 原始 VBA→Web 重構計畫 |
| `RawData 提取規則.md` | 保留 | 原始資料提取規則文件 |
| `使用說明.md` | 保留 | 使用者說明 |
| `QC表單編碼.txt` | 保留 | QC 編碼參考 |

### 提取邏輯確認（2025）：

| 元件 | 實作狀態 | 關鍵決策 |
|------|---------|---------|
| 進料檢驗月份 | ✅ 已修復 | 從「檢號」右側 date code 提取；`2025-227` 異常格式已容錯降級為 month=2 |
| 裝配巡檢計數 | ✅ 已修復 | `sSubFlat` 加 `false` 參數，每個 sheet = 1 筆（17 rows/月） |
| 裝配C「其他」拆分 | ✅ 已完成 | QC10006-R02 → 半成品品檢; QC10007-R03 → 零組件入庫品檢 |
| QIP Setup/Patrol 分組 | ✅ 已完成 | Sheet 名稱 `-MM-Setup` / `-MM-Patrol` 解析 |
| 月分組規則 | ✅ 完整 | 數值碼/字母碼/Date code/檔案名稱/Sheet 名稱 5 種規則均已實作 |



