# Concept: Automated Code Gates (自動化代碼防護欄)

This concept outlines the security rules, validation procedures, and testing standards that must be met before code changes are committed, built, or deployed.

## 🛡️ Security Boundaries

1.  **NO DYNAMIC RUNTIMES**: Usage of `eval()`, `exec()`, or direct compilation of raw strings into live instructions is strictly forbidden. Any runtime customization must be achieved via declarative JSON/YAML schemas.
2.  **API & UI Permissions Alignment**: Frontend components that trigger secure backend API actions (such as backups or deletion) must be conditionally hidden or disabled if the active session does not possess corresponding permissions. Never show a button that returns a HTTP 403 on click.
3.  **Import Validation**: All imports must be statically checked to ensure they refer to actual modules in the workspace or explicit dependencies in `package.json` / `pyproject.toml`.

## 🔄 PDCA Surgical Edits

We enforce the PDCA (Plan-Do-Check-Act) protocol for all code edits:
*   **Plan (Diagnosis)**: Scan the dependency tree and analyze the "blast radius" of the change. Do not make trial-and-error changes.
*   **Do (Atomic Edits)**: Keep diffs minimal. Only edit what is required to fulfill the task. Record RCA/CAPA in `DEV_LOG.md`.
*   **Check (Verification)**: Execute automated test suites or verification scripts. Ensure zero compiler warnings and console errors.
*   **Act (Review)**: Push only after testing succeeds locally and the user approves.
