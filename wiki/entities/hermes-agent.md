# Hermes Agent (Nous Research)

**來源**：https://github.com/NousResearch/hermes-agent  
**整合日期**：2026-06-03  
**授權**：MIT  
**核心定位**：The self-improving AI agent — 具備閉環學習能力的自我進化代理

---

## 核心定位

Hermes Agent 是 Nous Research 開發的自我進化 AI 代理，其最核心的差異化在於：
**它是唯一具備內建閉環學習迴路的代理** — 從任務中建立技能、在使用中改善技能、主動沉澱持久記憶、跨 session 搜尋過往對話、並持續建立使用者畫像。

---

## 核心能力矩陣

### 1. Closed Learning Loop (閉環學習)
最核心的差異化能力。三個組成部分：

**a. Agent-Managed Skills (skill_manage)**
- 完成 5+ 工具調用的複雜任務後，自動建立 SKILL.md
- 支援 create / patch / edit / write_file / delete 動作
- `patch` 是首選（最省 token）
- 對應本專案：`skill-creator` skill（已強化）

**b. Periodic Memory Nudges**
- 發現使用者偏好、環境事實、有效工作流時主動儲存
- 不需使用者要求，代理主動觸發
- 對應本專案：`session-memory` skill（新增）

**c. Skills Self-Improve During Use**
- 每次使用 skill 時，若發現更好的做法，自動 patch 改善
- 對應本專案：`soul-evolution` skill（已強化）

### 2. Persistent Memory (持久記憶)
雙軌設計：

| 文件 | 用途 | 容量 |
|------|------|------|
| `MEMORY.md` | 代理工作記憶（環境、慣例、學習到的技巧） | ~2,200 chars / ~800 tokens |
| `USER.md` | 使用者畫像（偏好、溝通風格、技術水準） | ~1,375 chars / ~500 tokens |

- 注入方式：session 開始時作為 frozen snapshot 注入 system prompt
- 優化前提：Character limit 強制記憶保持精簡聚焦
- 安全機制：injection 模式掃描、拒絕含憑證或惡意指令的條目

對應本專案：`session-memory` skill（新增）+ `wiki/MEMORY.md` + `wiki/USER.md`

### 3. Subagent Delegation & Parallelism (子代理並行)
- `delegate_task` 工具派生隔離子代理
- 預設 3 個並行子代理（可配置）
- 各子代理有獨立 context、restricted toolsets、獨立 terminal session
- `execute_code` 工具：寫 Python 腳本調用工具，零 context 成本

對應本專案：`dispatching-parallel-agents` skill（現有）

### 4. Scheduled Automations - Cron (排程自動化)
- 自然語言定義排程任務
- 支援時間型（cron expression）和事件型（平台 delivery）
- 任務結果可送達 Telegram / Discord / Slack 等平台
- 支援 pause / resume / edit 操作

對應本專案：`cron-automations` skill（新增）+ Kiro hooks

### 5. Skills System (技能系統)
- Progressive Disclosure 模式：Level 0（索引列表）→ Level 1（完整內容）→ Level 2（參考文件）
- 相容 agentskills.io 開放標準
- 支援 Skill Bundles（多 skill 一次載入）
- Skills Hub 整合：skills.sh / OpenAI / Anthropic / Hugging Face / GitHub 等多個來源

對應本專案：整個 `skills/` 目錄結構（現有，高度相容）

### 6. SOUL.md / Personality (人格核心)
- SOUL.md 是 system prompt 的第一個區塊
- 支援 /personality 指令切換預設人格
- 分為 IMMUTABLE（核心身份）和 EVOLVABLE（協作策略、學習偏好）兩區域

對應本專案：`soul-evolution` skill（已強化）+ `GEMINI.md`

---

## 技術架構重點

### 記憶系統
- MEMORY.md + USER.md → session 開始時 frozen snapshot 注入
- Session Search：SQLite + FTS5 全文搜尋所有歷史對話
- 外部記憶提供者：Honcho、Mem0、RetainDB 等 8 個插件

### Skills 技術規格
```yaml
---
name: skill-name
description: Brief description
version: 1.0.0
platforms: [macos, linux, windows]  # 可選
metadata:
  hermes:
    tags: [tag1, tag2]
    category: devops
    fallback_for_toolsets: [web]    # 條件性啟用
    requires_toolsets: [terminal]   # 條件性啟用
---
```

### 平台支援
- CLI / TUI、Telegram、Discord、Slack、WhatsApp、Signal
- 六種 terminal backend：local、Docker、SSH、Singularity、Modal、Daytona
- ACP IDE 整合（VS Code、Zed、JetBrains）

---

## 對本專案的整合貢獻

| Hermes 能力 | 整合至 SkillsBuilder | 新增/強化 |
|-------------|---------------------|---------|
| Closed Learning Loop | `soul-evolution` + `skill-creator` | 強化 |
| Persistent Memory | `session-memory` + `knowledge-bridge` | 新增 + 強化 |
| Agent-Managed Skills | `skill-creator` | 強化 |
| Cron Automations | `cron-automations` | 新增 |
| SOUL.md Pattern | `soul-evolution` | 強化 |
| Knowledge Taxonomy | `knowledge-bridge` | 強化 |

---

## 參考連結
- README: https://raw.githubusercontent.com/NousResearch/hermes-agent/main/README.md
- 官方文件: https://hermes-agent.nousresearch.com/docs/
- Skills System: https://hermes-agent.nousresearch.com/docs/user-guide/features/skills
- Memory System: https://hermes-agent.nousresearch.com/docs/user-guide/features/memory
- Features Overview: https://hermes-agent.nousresearch.com/docs/user-guide/features/overview
- Skills Hub: https://agentskills.io
