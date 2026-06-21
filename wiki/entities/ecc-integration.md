# ECC Integration Overview

ECC（Everything Claude Code）框架的核心能力已整合至 SkillsBuilder 專案，使 SkillsBuilder 成為 ECC 優勢與自身既有能力的功能超集（Superset）。

## ECC Framework Core Concepts

### Skills

Skills 是 SkillsBuilder 中的最小技能單元，以目錄形式存放於 `skills/{category}/{skill-name}/`。每個 Skill 包含 `SKILL.md` 文件，定義技能的名稱、描述、觸發關鍵詞與前置條件。

本次整合新增 15 個技能至 `skills/dev/` 目錄，涵蓋語言專屬 Reviewer/Resolver、安全掃描、Hooks 增強、Harness Optimizer、Loop Operator 等範疇。

### Subagents

ECC 提供 63 個專業化 Subagents，每個 Subagent 專精於特定開發任務（如代碼審查、錯誤診斷、建構修復等）。

SkillsBuilder 現有 42 個技能已整合多種 Subagent 工作流程，本次整合進一步擴充語言專屬審查與建構修復能力。

### Hooks

ECC 的 Hooks 自動化系統可於特定 IDE 事件（如 PreToolUse、PostToolUse、檔案儲存）觸發自動化腳本或 Agent。

本次整合新增 `hooks-enhancer` 技能，提供 4 種 Hook 配置範本：
- `auto-formatter`：儲存後自動執行格式化工具
- `tsc-type-check`：TypeScript 檔案修改後自動執行型別檢查
- `console-log-detector`：偵測代碼中殘留的除錯語句
- `import-validator`：儲存時識別無法解析的 import 路徑

### AgentShield

AgentShield 是 ECC 的安全掃描子系統，於 git push 事件觸發前自動執行代碼安全審查，攔截高風險安全問題。

AgentShield 執行四項掃描：
1. 硬編碼秘鑰（API Key、密碼、Token）偵測
2. `eval()` 與 `exec()` 動態執行指令偵測
3. 相依套件的已知高風險版本掃描
4. SQL 與 Shell 注入風險模式偵測

## 15 新增 ECC 技能清單

### Language Reviewers（6 skills）

 language Reviewers 針對不同程式語言執行語法、風格、最佳實踐審查，並整合對應靜態分析工具。

- **typescript-reviewer**：TypeScript 專屬代碼審查，整合 tsc 與 eslint
- **python-reviewer**：Python 專屬代碼審查，整合 pylint 與 mypy
- **go-reviewer**：Go 專屬代碼審查，整合 gofmt 與 go vet
- **rust-reviewer**：Rust 專屬代碼審查，整合 rustc 與 clippy
- **django-reviewer**：Django 專屬代碼審查，整合 flake8 與 django-lint
- **kotlin-reviewer**：Kotlin 專屬代碼審查，整合 kotlinc 與 detekt

### Language Resolvers（4 skills）

Language Resolvers 針對特定語言的建構錯誤進行自動診斷與修復。

- **typescript-build-resolver**：TypeScript 建構錯誤自動診斷與修復
- **python-build-resolver**：Python 建構錯誤自動診斷與修復
- **go-build-resolver**：Go 建構錯誤自動診斷與修復
- **rust-build-resolver**：Rust 建構錯誤自動診斷與修復

### Tools（5 skills）

這 5 個工具強化 ECC 的核心能力。

- **agent-shield**：安全掃描子系統，於 git push 前執行多層安全審查
- **hooks-enhancer**：Hooks 配置生成器，提供 IDE 專屬的自動化品質檢查範本
- **harness-optimizer**：Context Window 管理與 token 成本優化
- **ecc-migrator**：ECC Skills 遷移工具，將 ECC Workflow Skills 轉換為 SkillsBuilder 格式
- **loop-operator**：Agent 執行迴路監控與異常偵測介入

## Usage

### 安裝與同步

執行 `INSTALL.ps1` 自動同步所有 ECC 技能：

```powershell
powershell -ExecutionPolicy Bypass -File INSTALL.ps1
```

### 觸發技能

在支援的 IDE 中，透過以下方式觸發技能：

- **Kiro**：在提示中輸入技能的觸發關鍵詞
- **Claude Code**：在提示中提及技能名稱或使用 `/skills` 命令
- **Cursor**：在提示中提及技能名稱或使用 `@skills/` 指令

### 配置 Hook

`hooks-enhancer` 技能提供 IDE 專屬的 Hook 配置範本，請參閱 `skills/dev/hooks-enhancer/` 目錄下的範例文件。

### 啟用 AgentShield

將 AgentShield 掛載至 git push 事件前，請參閱 `skills/dev/agent-shield/hook-examples/` 目錄下的 Hook 配置範例。

## 相關文件

- `docs/ecc-integration-guide.md`：ECC 整合完整指南（包含能力對應表、AgentShield 啟用指南、hooks-enhancer 配置教學）
- `wiki/index.md`：SkillsBuilder 主索引頁
