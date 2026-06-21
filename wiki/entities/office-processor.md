# Entity: Office Document Processor

## Description
A data processing skill designed to read, parse, process, and write Microsoft Office files (`.docx`, `.xlsx`, `.pptx`) and Adobe PDF (`.pdf`) formats.

## Core Capabilities
*   **Data Scraper**: Batch extracting tables, charts, metadata, and body text from spreadsheet or text documents.
*   **Document Generation**: Populating template fields, generating formatted PDF reports, and exporting structured CSV sheets.
*   **Formatting Guardrails**: Enforcing strict styling rules (font sizes, layout margins) on generated outputs.

## Integration
*   **Location**: `skills/core/office-processor/`
*   **Dependencies**: Python (`pandas`, `openpyxl`, `python-docx`) or Node.js (`xlsx`, `docx`, `pdf-parse`).

## Verification
*   Read-back validation scripts ensure output files are readable and contain expected fields before completion.
