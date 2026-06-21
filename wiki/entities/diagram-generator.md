# Entity: Excalidraw Diagram Generator

## Description
A code-based graphics synthesis tool that allows the agent to build flowcharts, sequence diagrams, and architecture blueprints in standard copy-pasteable Excalidraw JSON structure.

## Core Capabilities
*   **Coordinate Calculation**: Layout engine that positions boxes, connectors, diamonds, and texts without overlaps.
*   **JSON Schema Integrity**: Output complies with the official Excalidraw JSON format (`type`, `version`, `elements`).
*   **Copy-Paste Portability**: Direct imports of files generated at `scratch/diagram.excalidraw` on [Excalidraw](https://excalidraw.com/).

## Integration
*   **Location**: `skills/dev/diagram-generator/`
*   **Dependencies**: None.

## Verification
*   Validating JSON structure syntax before exporting.
