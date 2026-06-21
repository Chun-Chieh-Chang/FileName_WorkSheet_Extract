# Entity: Find Skills

## Description
A package manager interface that enables AI agents to search, discover, install, and update modular skills from the open agent skills ecosystem (ClawHub / `skills.sh`).

## Capabilities
- **Ecosystem Search**: Search the global `skills.sh` registry using natural keywords.
- **Skill Installation**: Autonomously install community skills via command line from GitHub or public registries.
- **Dependency & Update Check**: Check and update all installed skills to maintain current capabilities.

## Magic Phrases
- 「幫我 **Search** 一個可以...的 Skill」
- 「這裏有可以做...的技能嗎？」
- 「Is there a skill that can do X?」

## Usage Patterns
- **Discovery Phase**: Used during session startup or brainstorming when a task requires highly specialized tools not present in the local catalog.
- **Expansion Phase**: Used to expand the agent's capability dynamically instead of writing complex custom workflows from scratch.

## Guardrails
- **Leaderboard First**: Always check the `skills.sh` leaderboard for verified options before performing a blind registry search.
- **Security vetting**: Vetted skills must be preferred. If the user installs custom repositories, notify them and follow safety checklists.
- **Creation Fallback**: If no relevant skill exists in the ecosystem, offer to help directly or guide the user to scaffold a new one via `npx skills init`.
