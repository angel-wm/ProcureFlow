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

## Procedure Design and Visibility

Phase 8 also introduced procedure visibility and responsibility boundaries.

`Public` procedures were used as executable educational entry points.

`Private` procedures and functions were used for internal helper logic that should not be exposed as standalone workbook actions.

Examples include:

- `Public Sub GoToControl`
- `Private Sub PrintNavigationResult`
- `Private Function WorksheetExists`
- `Private Function TableExists`

This reinforces the distinction between a procedure that coordinates an action and helper code that exists only to support that action.

## ByVal and ByRef

The learning exercises explicitly compared `ByVal` and `ByRef`.

`ByVal` passes a value without allowing the called procedure to modify the caller's original variable.

`ByRef` allows the called procedure to modify the original variable.

ProcureFlow will prefer `ByVal` unless modifying the caller's variable is an intentional and clearly documented part of the procedure contract.

## Reusable Validation Functions

The exercises introduced reusable validation functions such as:

`WorksheetExists`

and:

`TableExists`

These functions demonstrate explicit workbook-object validation without using hidden runtime errors as the normal control-flow mechanism.

The validation hierarchy demonstrated was:

`ValidateReplenishmentObjects`
→ `TableExists`
→ `WorksheetExists`

This keeps coordination logic separate from reusable validation logic.

## Structured Error Handling

`GoToControl` was refactored to demonstrate a structured error-handling pattern based on:

`On Error GoTo ErrorHandler`

`CleanExit`

`ErrorHandler`

`Err.Number`

`Err.Description`

and:

`Resume CleanExit`

The example distinguishes normal procedure exit from error handling while avoiding silent failure.

The learning standard established during Phase 8 is that critical failures must not be hidden with unrestricted use of `On Error Resume Next`.

## Code Documentation Standard

The Phase 8 educational VBA module was reviewed and documented with clear English comments.

Comments focus on:

- procedure purpose;
- architectural boundaries;
- non-obvious decisions;
- validation intent;
- error-handling intent;
- educational versus production responsibilities.

Comments are not added mechanically to lines whose behavior is already self-explanatory.

## PivotTables and PivotCaches

Phase 8 introduced the Excel VBA object model for PivotTables.

The exercises inspected:

`Worksheet.PivotTables`

`PivotTable`

`PivotTable.TableRange2`

`PivotTable.CacheIndex`

`PivotTable.PivotCache`

`Workbook.PivotCaches`

`PivotCache`

The exercises demonstrated that a PivotTable and its PivotCache are separate Excel objects and that multiple PivotTables may share the same cache.

The learning exercises inspected the real PivotTables created during ProcureFlow Phase 7 without changing their fields, filters, or layouts.

## Refresh Scope

Phase 8 introduced three different refresh levels:

`PivotTable.RefreshTable`

`PivotCache.Refresh`

`Workbook.RefreshAll`

These methods have different scopes and must not be treated as interchangeable.

`PivotTable.RefreshTable` targets a specific PivotTable report.

`PivotCache.Refresh` targets a specific PivotCache.

`Workbook.RefreshAll` requests refresh operations across refreshable workbook objects and may include connections, queries, and dependent Excel objects.

## Controlled Refresh Exercise

Phase 8 executed a controlled learning example using:

`PivotTable.RefreshTable`

on one existing PivotTable.

The exercise was intentionally limited to a single PivotTable and was not implemented as the ProcureFlow production refresh workflow.

No general `Workbook.RefreshAll` automation was introduced during Phase 8.

## Phase 8 / Phase 9 Boundary

Phase 8 establishes understanding of PivotTable and refresh objects.

Phase 9 owns the design and implementation of the production refresh workflow.

That later workflow must consider the actual ProcureFlow dependency chain, including:

Power Query
→ structured tables
→ calculated model outputs
→ PivotTables
→ reporting outputs

Production automation must also prevent stale outputs and must not report a refresh as successful before its required dependencies have completed.

The Phase 8 examples therefore demonstrate refresh concepts without preempting the automation architecture that must be designed in Phase 9.
