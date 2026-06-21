# Entity: AutoResearch

## Description
An autonomous machine learning research and codebase optimization framework designed to automate the Propose-Train-Evaluate loop. It operates as a "Ratchet", committing only successfully validated improvements to git and rolling back failures.

## Capabilities
- **Baseline Establishment**: Automatically establishes a baseline run tag and creates target experimental git branches (`autoresearch/<tag>`).
- **Closed-Loop Optimization**: Surgically modifies code, executes training within a fixed time budget, measures loss/efficiency metrics, and decides whether to keep or discard modifications.
- **Incremental Logging**: Appends experiment records to `results.tsv` and registers standard commit logs.

## Magic Phrases
- 「幫我啟動 **AutoResearch** 實驗」
- 「開啟自主優化與 Ratchet 閉環」
- 「Start autonomous ML tuning on this codebase」

## Usage Patterns
- **Experiment Phase**: Used when optimizing hyperparameters, transformer block architectures, attention mechanisms, or optimizers.
- **Refactoring Phase**: Used to safely prune unused code while validating that metrics do not degrade (Simplification Win).

## Guardrails
- **Fixed Time Budget**: Ensure training seconds are strictly fixed to make all validation scores comparable.
- **Strict Rollback**: Immediately revert changes via `git checkout` if validation scores degrade or the script throws errors.
- **Read-Only Constants**: Do NOT touch evaluation constants, tokenizers, or prepare scripts.
