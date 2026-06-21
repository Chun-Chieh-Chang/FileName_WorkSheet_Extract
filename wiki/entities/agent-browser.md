# Entity: Agent Browser

## Description
A browser automation capability enabling the agent to execute visual workflows, submit web forms, navigate target sites, and extract unstructured information.

## Core Capabilities
*   **Action Automation**: Script-driven page navigation, click sequences, text input, and scroll behaviors.
*   **Visual Recording**: Capturing screenshots of the browser viewport (`scratch/browser_last_state.png`) for manual validation and QA checks.
*   **JSON Scraping**: Converting raw HTML DOM pages into structured JSON payloads.

## Integration
*   **Location**: `skills/core/agent-browser/`
*   **Dependencies**: Playwright or Puppeteer installed in Node.js runtime.

## Verification
*   All visual scripts must output standard JSON data.
*   Visual audits require confirming viewport screenshots are saved in `scratch/`.
