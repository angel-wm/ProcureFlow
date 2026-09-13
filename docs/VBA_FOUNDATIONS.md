# VBA Foundations

## Purpose

This document records the VBA foundations practiced during Phase 8 of ProcureFlow.

The examples created in this phase are learning artifacts. They are not production automation and are not part of the final operational behavior of the workbook.

The corresponding VBA source is maintained under:

`vba/examples/phase08/modVBAFoundations.bas`

## Architecture Boundary

ProcureFlow follows the principle that VBA orchestrates Excel processes but does not become a hidden business-logic or mathematical engine.

Business calculations remain primarily in the appropriate Excel formulas, structured tables, and other approved model layers.

Phase 8 focuses on understanding VBA.

Production automation belongs to Phase 9.

## Concepts Practiced

### Macro Recorder and VBE

The Phase began by recording a simple workbook-navigation macro and inspecting the generated VBA in the Visual Basic Editor.

This demonstrated the relationship between:

- an Excel macro;
- VBA source code;
- the VBE;
- executable code stored inside the `.xlsm` workbook.

### Procedures

The exercises introduced `Sub` procedures and procedure calls.

Examples include:

`GoToControl`

`PrintNavigationResult`

### Variables and Data Types

The exercises used `Dim` and explicit types including:

- `Worksheet`
- `ListObject`
- `ListRow`
- `ListColumn`
- `String`
- `Long`
- `Boolean`

`Option Explicit` was used to require explicit variable declarations.

### Object References

Object variables were assigned using `Set`.

The exercises distinguished object references from normal value assignments.

### Excel Object Model

The exercises navigated the Excel object hierarchy using objects such as:

`ThisWorkbook`

`Worksheets`

`Worksheet`

`Range`

`Cells`

`ListObject`

`ListRows`

`ListColumns`

`ListRow`

`ListColumn`

The exercises demonstrated why named Excel Tables and structured objects are preferable to arbitrary cell coordinates for maintainable VBA.

### Conditions

The exercises used:

`If`

`ElseIf`

`Else`

`And`

`Or`

Conditions were applied to actual ProcureFlow outputs such as `InventoryStatus` and `RecommendedOrderQty`.

These exercises read existing business results; they did not recreate replenishment calculations in VBA.

### Loops

The exercises introduced:

`For`

`For Each`

`Do While`

They were used to iterate counters, worksheets, table rows, and table columns.

### Functions

The exercises introduced VBA functions through `WorksheetExists`, including:

- parameters;
- `ByVal`;
- Boolean return values;
- calling a function from a conditional expression.

### Error Handling

Basic runtime error handling was practiced using:

`On Error GoTo`

`Exit Sub`

`Err.Number`

`Err.Description`

The exercises also established that `On Error Resume Next` must not be used as a general mechanism for hiding failures.

## ProcureFlow Examples

The educational code inspected real ProcureFlow structures including:

`20_CALC_Replenishment`

`tblReplenishment`

and fields such as:

`ProductID`

`SiteID`

`InventoryStatus`

`RecommendedOrderQty`

The exercises were intentionally read-only and did not modify operational data or business formulas.

## Learning Artifact vs Production VBA

The Phase 8 module `modVBAFoundations` is an educational artifact.

It may remain temporarily inside `ProcureFlow.xlsm` while Phase 8 is in progress so the exercises can continue.

Before Phase 8 is formally closed, the educational module will be removed from the workbook while its exported `.bas` source and this documentation remain in the repository.

Production VBA implemented in later phases will use separately designed modules with real operational responsibilities.

## Current Status

Phase 8 VBA foundations exercises completed so far include:

- macro recording and inspection;
- VBE fundamentals;
- variables and types;
- Workbook and Worksheet objects;
- Range and Cells;
- ListObject and structured tables;
- ListRow and ListColumn;
- conditional logic;
- `For`, `For Each`, and `Do While`;
- procedures and functions;
- foundational error handling;
- read-only inspection of ProcureFlow business outputs.

Further Phase 8 learning and validation remain before the phase can be closed.
