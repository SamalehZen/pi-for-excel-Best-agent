# HyperFix (Pi for Excel) - PRD

## Project Overview
HyperFix is an open-source, multi-model AI sidebar add-in for Excel. It reads workbooks, makes changes, searches the web, and supports multiple providers (Anthropic, OpenAI, Google, GitHub Copilot).

## Architecture
- **Stack**: TypeScript, Vite, Lit Web Components, Office.js
- **Build**: Vite with custom plugins (stubbing Bedrock, OAuth, heavy deps)
- **UI**: LitElement-based `<pi-sidebar>` component with pi-web-ui library
- **Tools**: Modular tool system with registry pattern
- **Dependencies**: @mariozechner/pi-agent-core, pi-ai, pi-web-ui, just-bash, marked

## Bug Fix - April 1, 2026

### Problem
After merging PRs #28 to #34 (adding new Excel tools), the add-in displayed a blank white page.

### Root Cause
PRs #28, #29 introduced module-level constants that referenced `Excel.*` enum values (e.g., `Excel.ChartType`, `Excel.DataValidationOperator`). These constants are evaluated at import time, before the `Excel` namespace is available from Office.js, causing a fatal `Excel is not defined` error that crashed the entire app initialization.

### Files Fixed
1. **src/tools/create-chart.ts** - Converted `CHART_TYPE_MAP`, `SERIES_BY_MAP`, `LEGEND_POSITION_MAP` from module-level constants to lazy getter functions
2. **src/tools/create-pivot-table.ts** - Converted `AGGREGATION_FUNCTION_MAP` to lazy getter function
3. **src/tools/data-validation.ts** - Converted `OPERATOR_MAP` to lazy getter function
4. **src/tools/range-operations.ts** - Converted `COPY_TYPE_MAP` and `DELETE_SHIFT_MAP` to lazy getter functions

### Pattern Applied
```typescript
// BEFORE (crashes at import time):
const MAP = { key: Excel.SomeEnum.value };

// AFTER (lazy - only evaluated when called):
let _map = null;
function getMap() {
  if (!_map) { _map = { key: Excel.SomeEnum.value }; }
  return _map;
}
```

## PRs Merged (#28-#34)
- #28: create_chart, create_table, create_pivot_table core tools
- #29: data_validation and range_operations core tools
- #30: screenshot_range read-only core tool
- #31: update/delete actions for chart and pivot table tools
- #32: sandboxed bash tool with VFS
- #33: clickable #cite: citations in chat UI
- #34: template picker cards redesign

## Backlog
- P0: None (app is functional)
- P1: Test all new tools inside Excel
- P2: Code-splitting for smaller bundle size (4.5MB taskpane bundle)
