# ProcureFlow — User Guide

## Purpose

This guide explains how to set up, refresh and use the ProcureFlow workbook in its supported operating environment.

ProcureFlow is an Excel-based Procurement & Inventory Management System designed for inventory monitoring, replenishment decisions, supplier-performance analysis, Quality Control, operational reporting and management reporting.

The executable workbook is:

`workbook/ProcureFlow.xlsm`

The supported platform is:

Microsoft Excel 365 Desktop for Windows.

### How to use this guide

Use the shortest path that matches your task:

- before first use, review [Requirements](#2-requirements), [Repository and Folder Structure](#3-repository-and-folder-structure), and [Source Data Setup](#4-source-data-setup);
- for the normal operating sequence, go to [Standard User Workflow](#8-standard-user-workflow);
- for refresh behavior and accepted states, use [Production Refresh Workflow](#9-production-refresh-workflow) through [Failed Refresh](#12-failed-refresh);
- when something goes wrong, start with [Troubleshooting](#20-troubleshooting);
- for the supported production boundary, see [Supported Operating Boundary](#23-supported-operating-boundary).

---

## 1. Operating Model

The normal ProcureFlow workflow is:

source CSV files
→ Power Query
→ structured Excel Tables
→ operational calculations
→ Quality Control
→ PivotTables and analytical outputs
→ operational report
→ management dashboard

Production refresh orchestration is handled by Visual Basic for Applications (VBA).

The supported production entry point is:

`RefreshProcureFlow`

The normal user-facing way to run it is the:

`Refresh ProcureFlow`

button on:

`00_HOME`

Users should not need to copy and paste source data manually into the workbook.

---

## 2. Requirements

To operate ProcureFlow as designed, the following are required:

- Microsoft Excel 365 Desktop for Windows;
- the `ProcureFlow.xlsm` workbook;
- the four official source CSV files;
- the approved repository folder relationship;
- permission to execute the workbook's VBA macros according to the security policy of the user's environment.

The production automation depends on VBA.

If VBA execution is disabled by the user's environment or organizational security policy, the supported production refresh workflow cannot run as designed.

Do not weaken organizational security controls solely to run the workbook.

---

## 3. Repository and Folder Structure

The relevant runtime structure is:

    ProcureFlow/
    ├── workbook/
    │   └── ProcureFlow.xlsm
    └── data/
        └── raw/
            ├── parts_master.csv
            ├── supply_chain_history.csv
            ├── purchase_orders.csv
            └── quality_incidents.csv

The relationship between:

`workbook/`

and:

`data/raw/`

must be preserved.

ProcureFlow derives the raw-data folder from the workbook location through the workbook-scoped configuration name:

`cfg_RawDataFolder`

Production Power Query queries do not contain a machine-specific absolute source path.

---

## 4. Source Data Setup

ProcureFlow uses the Aerospace Supply Chain Performance & Forecasting dataset.

The required source files are:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

Place all four files in:

`data/raw/`

The raw CSV files are intentionally excluded from Git and are not distributed as part of the normal repository history.

The original source files must remain unchanged.

Do not:

- rename the required source files;
- manually clean the CSV files;
- change their schema as part of the normal workflow;
- copy their contents manually into workbook worksheets.

Power Query is responsible for ingestion, typing, preparation and validation.

---

## 5. Opening ProcureFlow

1. Preserve the approved repository folder structure.
2. Confirm that all four source CSV files exist in `data/raw/`.
3. Open:

   `workbook/ProcureFlow.xlsm`

4. Use Microsoft Excel 365 Desktop for Windows.
5. Allow the workbook's VBA functionality only when permitted by the security policy of the environment.
6. Begin from:

   `00_HOME`

The workbook should open without a repair prompt or workbook-content error.

---

## 6. Main User-Facing Worksheets

### `00_HOME`

Primary operational entry point.

It exposes:

- Reporting Date;
- Last Successful Refresh;
- Overall Quality Status;
- navigation to the main functional areas;
- the `Refresh ProcureFlow` production button.

### `01_CONFIG`

Central business-configuration sheet.

Editable configuration inputs are located in:

`B6:B12`

The sheet remains protected against unintended structural edits while the approved configuration inputs remain editable.

### `02_CONTROL`

System-control and Quality Control layer.

Use it to inspect:

- control status;
- exception counts;
- PASS / WARNING / FAIL results;
- automation state;
- refresh failure information;
- reconciliation evidence.

### `40_RPT_Replenishment`

Operational replenishment report.

Use it to determine which Product × Site positions require action and what quantity is recommended.

### `41_DASH_Management`

Management dashboard.

Use it for management-level inventory, procurement and supplier-performance monitoring.

Analytical worksheets remain available for deeper investigation:

- `30_PVT_Inventory`
- `31_PVT_Procurement`
- `32_PVT_Suppliers`

---

## 7. Business Configuration

The approved configurable parameters are:

- Reporting Date;
- Demand History Weeks;
- Review Period Weeks;
- Service Level A;
- Service Level B;
- Service Level C;
- Excess Buffer Weeks.

Initial approved values are:

| Parameter | Initial Value |
|---|---:|
| Demand History Weeks | 26 |
| Review Period Weeks | 1 |
| Service Level A | 99% |
| Service Level B | 97% |
| Service Level C | 95% |
| Excess Buffer Weeks | 4 |
| Reporting Date | Valid selected date |

For the current official dataset, the initial maximum Reporting Date is:

`23-Dec-2024`

Reporting Date must remain within the valid source-data range.

Valid Service Levels must satisfy:

`0 < Service Level < 1`

Therefore:

- `0%` is invalid;
- `100%` is invalid.

This restriction is required because the approved Safety Stock methodology uses:

`NORM.S.INV()`

Worksheet Data Validation and Quality Control enforce the approved Service Level domain.

---

## 8. Standard User Workflow

The recommended operating sequence is:

1. Open `ProcureFlow.xlsm`.
2. Go to `01_CONFIG`.
3. Review the Reporting Date and other configurable business parameters.
4. Return to `00_HOME`.
5. Run `Refresh ProcureFlow`.
6. Allow the complete production workflow to finish.
7. Verify the resulting system state.
8. Confirm Overall Quality Status.
9. Review the operational replenishment report.
10. Review the management dashboard and analytical views as required.

Do not treat refreshed outputs as accepted merely because worksheet values changed.

The workflow state and Quality Control result are part of the acceptance condition.

---

## 9. Production Refresh Workflow

The sole supported production refresh entry point is:

`RefreshProcureFlow`

The user-facing button on `00_HOME` invokes this procedure.

The workflow performs the controlled sequence required to update the system, including:

1. marking the automation attempt as running;
2. refreshing the loaded Power Query outputs synchronously;
3. validating technical refresh completion;
4. recalculating Excel formulas;
5. refreshing the analytical PivotTable layer;
6. evaluating staged Quality Control;
7. persisting the final automation state;
8. advancing `LastSuccessfulRefresh` only when the required workflow has completed successfully.

The production implementation refreshes:

- 8 physically loaded Power Query outputs;
- 5 unique PivotCaches;
- 10 PivotTables.

The complete full-dataset production refresh was measured during Phase 11 at approximately 12 minutes.

This is a measured performance characteristic, not a guaranteed execution time.

Actual duration may vary by machine and environment.

---

## 10. Successful Refresh

A valid successful execution should result in an accepted system state such as:

- `WorkflowStatus = SUCCESS`;
- `PivotRefreshStatus = PASS`;
- Overall Quality Status = `PASS`;
- failure fields blank;
- `LastSuccessfulRefresh` advanced to the newly accepted refresh timestamp.

After a successful refresh, users may proceed to operational and management analysis.

---

## 11. Warning State

A valid workflow may expose a warning condition where applicable.

A WARNING must not be silently interpreted as PASS.

When a WARNING is present:

1. review `02_CONTROL`;
2. identify the warning-producing control;
3. understand the associated exception;
4. determine whether the condition is expected and acceptable before using the affected output for a business decision.

Important status meaning is always represented by text and does not depend exclusively on color.

---

## 12. Failed Refresh

If the workflow fails:

- `WorkflowStatus` is recorded as `FAILED`;
- Quality Control exposes the invalid current state;
- failure information is retained for diagnosis;
- the previously accepted `LastSuccessfulRefresh` is preserved.

A failed attempt must not advance the successful-refresh timestamp.

This is an important safety property.

The existence of a previous `LastSuccessfulRefresh` does not mean the latest attempt succeeded.

When a refresh fails:

1. do not treat the failed attempt as accepted current output;
2. inspect `02_CONTROL`;
3. review the recorded failure step and error information;
4. verify that all four required source files are present;
5. verify that source filenames and repository folder structure are unchanged;
6. verify configuration values;
7. correct the identified issue;
8. run `Refresh ProcureFlow` again;
9. confirm a successful workflow and acceptable Quality Control state before relying on updated outputs.

---

## 13. Quality Control

`02_CONTROL` is the centralized system-control layer.

Quality states include:

- PASS;
- WARNING;
- FAIL;
- NOT EVALUATED where applicable.

Quality Control covers technical, configuration, integrity, calculation and reconciliation conditions.

A valid release-data workflow must not silently hide critical data-quality failures.

Before relying on refreshed business outputs, confirm that the Quality Control state is appropriate.

---

## 14. Replenishment Workflow

For day-to-day inventory action:

1. complete or confirm a successful refresh;
2. open `40_RPT_Replenishment`;
3. review the current Reporting Date and refresh context;
4. use the report filters as required.

Supported operational filters include:

- Site;
- Part Family;
- Criticality;
- Supplier Risk;
- Inventory Status.

The default:

`ACTIONABLE`

view excludes only:

`HEALTHY`

positions.

It does not create a second Inventory Status classification.

The report preserves the authoritative replenishment statuses generated upstream by the calculation engine.

Use the report to identify:

- positions requiring action;
- inventory status;
- replenishment priority;
- Recommended Order Quantity;
- related operational context.

---

## 15. Supplier and Procurement Analysis

For procurement and supplier review, use:

- `31_PVT_Procurement`
- `32_PVT_Suppliers`
- `41_DASH_Management`

Supplier-performance analysis includes validated delivery, Lead-Time and quality metrics.

The global On-Time Delivery Rate shown on the management dashboard is weighted using:

`SUM(OnTimePOCount) / SUM(ReceivedPOCount)`

It must not be interpreted as a simple average of Supplier-level delivery-rate percentages.

---

## 16. Management Dashboard

`41_DASH_Management` provides management-level decision support.

The implemented dashboard includes KPI coverage for:

- STOCKOUT positions;
- CRITICAL positions;
- REORDER positions;
- Recommended Order Quantity;
- Backorder Quantity;
- Open PO Quantity;
- On-Time Delivery Rate;
- Quality Incident Count.

Implemented visual analysis includes:

- Inventory Status by Site;
- Supplier Delivery Performance by Risk Class.

The dashboard is a presentation layer over validated upstream calculations.

It does not implement a separate business-rule engine.

---

## 17. Protection and Editable Areas

Workbook protection is intended to reduce accidental modification of critical structures.

It is not a security boundary.

Validated protection includes:

- `01_CONFIG`;
- critical formulas in `20_CALC_Replenishment`;
- critical formulas in `21_CALC_SupplierPerformance`.

Approved configuration inputs in:

`01_CONFIG!B6:B12`

remain editable.

Critical formulas remain inspectable.

Required filters and navigation remain usable.

No worksheet-protection password is part of the validated Phase 11 baseline.

---

## 18. Raw Data Rules

The normal operating process must preserve source traceability.

Users must not:

- manually modify the official source CSV files;
- replace Power Query preparation with manual worksheet cleanup;
- paste source data directly into final structured tables;
- overwrite calculation formulas with values;
- bypass Quality Control when evaluating whether a refresh is acceptable.

If source data must change, replace it only through the approved source-file workflow and then execute the production refresh again.

---

## 19. What Not to Modify

Normal business users should avoid structural changes to:

- Power Query queries;
- workbook-scoped Defined Names;
- structured data tables;
- calculation formulas;
- Quality Control formulas;
- PivotTables and PivotCaches;
- production VBA modules;
- protected calculation structures.

Those components are part of the maintained technical implementation.

Configuration inputs and supported report filters are the intended normal user controls.

---

## 20. Troubleshooting

Use the current workflow state and Quality Control evidence to diagnose problems. Do not infer success from changed worksheet values or from an older `LastSuccessfulRefresh` timestamp.

### Refresh does not start

**Symptom:** selecting **Refresh ProcureFlow** does not begin the supported production workflow.

**Likely causes:**

- the workbook is not open in Microsoft Excel 365 Desktop for Windows;
- VBA execution is blocked by the environment;
- the supported button on `00_HOME` is not being used.

**Resolution:**

1. open `ProcureFlow.xlsm` in the supported Excel desktop environment;
2. confirm that VBA execution is permitted by the applicable security policy;
3. return to `00_HOME`;
4. run **Refresh ProcureFlow** again.

**Verify:** inspect `02_CONTROL` and confirm that a new workflow attempt is recorded and the refresh proceeds beyond the initial state.

### Source-file error

**Symptom:** refresh reports that a required source cannot be found or loaded.

**Likely causes:**

- one or more required CSV files are missing;
- a required filename has changed;
- the repository relationship between `workbook/` and `data/raw/` has changed.

**Resolution:**

1. confirm that all four approved CSV files exist under `data/raw/`;
2. confirm the exact approved filenames;
3. restore the documented repository folder relationship if it was changed;
4. do not manually rewrite or clean the source files as a workaround;
5. run **Refresh ProcureFlow** again.

**Verify:** confirm that the source step completes without the same file error and then verify the final workflow and Quality Control state.

### Quality Control returns FAIL

**Symptom:** Overall Quality Status or an applicable control reports `FAIL`.

**Likely cause:** a source, configuration, calculation, reconciliation, or workflow condition violates an implemented control.

**Resolution:**

1. open `02_CONTROL`;
2. identify the failing control and its exception count;
3. review the associated result and failure information;
4. correct the underlying source, configuration, or workflow issue;
5. do not manually overwrite the control result;
6. rerun **Refresh ProcureFlow** when the underlying issue has been corrected.

**Verify:** confirm that the affected control no longer reports `FAIL` and that the resulting Overall Quality Status is acceptable before using refreshed outputs.

### Service Level is rejected or produces an invalid condition

**Symptom:** a Service Level entry is rejected by Data Validation or is identified as invalid by Quality Control.

**Likely cause:** the value is outside the supported domain required by the Safety Stock calculation.

Valid Service Levels must satisfy:

`0 < Service Level < 1`

**Resolution:** enter a valid value such as the approved defaults:

- 99%;
- 97%;
- 95%.

Do not use exactly 0% or 100%.

**Verify:** confirm that the value is accepted by worksheet validation and that the Service Level control does not report an invalid configuration after refresh.

### Latest refresh fails after an earlier successful run

**Symptom:** the latest workflow is `FAILED`, while `LastSuccessfulRefresh` still shows an earlier timestamp.

**Likely cause:** this is intentional rollback protection. A failed attempt does not replace the timestamp of the last accepted production refresh.

**Resolution:**

1. inspect the current workflow state in `02_CONTROL`;
2. review the recorded failure step and error information;
3. correct the current failure rather than relying on the historical timestamp;
4. run **Refresh ProcureFlow** again.

**Verify:** confirm a successful or otherwise accepted workflow state and verify that `LastSuccessfulRefresh` advances only after the new run is accepted.

---

## 21. Validated Acceptance Baseline

Phase 11 completed full-system regression and User Acceptance Testing, including Replenishment, Supplier Performance, Refresh, and Data Quality Failure scenarios. Phase 12 then completed final documentation, portfolio preparation, and release acceptance.

The detailed executed evidence is maintained in [Testing](TESTING.md); this guide summarizes only the operating baseline needed by users.

Current released version:

`v1.0.2`

---

## 22. Technical Documentation

For deeper technical information, use the document that matches the question:

- [Project Specification](PROJECT_SPEC.md) — requirements, business rules, and Definition of Done;
- [Architecture](ARCHITECTURE.md) — technical and workbook architecture;
- [Data Dictionary](DATA_DICTIONARY.md) — source-data and logical-model definitions;
- [Formula Catalog](FORMULAS.md) — versioned Excel business formulas;
- [Decision Log](DECISIONS.md) — material project decisions;
- [Testing](TESTING.md) — validation, regression, and acceptance evidence;
- [Current State](CURRENT_STATE.md) — authoritative current project state;
- [VBA Source Overview](../vba/README.md) — production and educational VBA source layout;
- [Data README](../data/README.md) — source-data handling policy.

---

## 23. Supported Operating Boundary

ProcureFlow `v1.0.2` retains the validated `v1.0.0` executable workbook and operating architecture. The `v1.0.2` maintenance release changes documentation readability, navigation, language consistency, and repository-state metadata only.

The supported operating boundary is:

- Excel 365 Desktop for Windows;
- approved workbook structure;
- approved source files and schemas;
- approved configuration controls;
- production refresh through `RefreshProcureFlow`;
- validated Quality Control and reporting layers.

Changes outside this boundary should be treated as maintenance or future development work and validated before being considered part of the supported production state.
