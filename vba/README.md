# VBA Source

This directory contains exported Visual Basic for Applications source code from `ProcureFlow.xlsm`.

Repository structure:

- `modules/` — production standard `.bas` modules
- `examples/` — educational VBA preserved separately from production code
- `classes/` — reserved for class modules if a future design requirement justifies them

The workbook remains the executable artifact, while exported VBA files provide readable and diffable source code for Git/GitHub.

## Production Modules

Phase 9 — Automation implements:

- `modAutomationState.bas` — persistent workflow-state management
- `modRefresh.bas` — controlled Power Query refresh orchestration
- `modQualityControl.bas` — staged Quality Control evaluation
- `modPivotRefresh.bas` — PivotCache and PivotTable refresh orchestration
- `modAutomation.bas` — end-to-end production refresh workflow

Production entry point:

`RefreshProcureFlow`

The production architecture follows the ProcureFlow rule:

VBA orchestrates; VBA does not duplicate the Excel business-calculation engine.

## Educational Source

Phase 8 learning source is preserved separately under:

`examples/phase08/`

Educational modules are not part of the supported production workbook functionality.