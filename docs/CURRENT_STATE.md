# ProcureFlow — Current State

## Project Status

IN DEVELOPMENT

## Current Version

v0.2.0

## Current Phase

Phase 2 — Workbook Foundation

Status: IN PROGRESS

Target Version: `v0.3.0`

Phase Branch:

`phase/02-workbook-foundation`

## Last Completed Phase

Phase 1 — Data Design

Version: `v0.2.0`

GitHub Pull Request: `#2 — Phase 1 — Data Design`

Merge Commit: `36da5ac`

Phase Tag: `phase-1-complete`

Version Tag: `v0.2.0`

## Implemented

Project governance and formal data-design baseline.

Implemented project assets include:

- repository and documentation governance;
- official dataset policy;
- validated source schemas;
- validated source grains;
- validated keys and relationships;
- field-level data dictionary;
- source-to-target mappings;
- logical dimension/fact model;
- Date dimension design;
- source-supported data-quality rules;
- Phase 1 validation evidence.

No Excel workbook functionality has been implemented yet.

No Power Query queries, formulas, PivotTables, VBA procedures, automation, operational reports or dashboards are considered implemented.

## Phase 1 Completion

Phase 1 formally validated the four official source files and established the logical data model required for workbook construction.

Confirmed source populations:

- Products: 300
- Sites: 6
- Suppliers: 40
- Inventory History: 280,800 rows
- Purchase Orders: 29,666 rows
- Quality Incidents: 368 rows

Validated Inventory History grain:

1 Product × 1 Site × 1 Week

Validated structure:

- 1,800 Product × Site combinations;
- 156 weeks per Product × Site;
- 156 distinct weekly dates;
- all weekly dates are Mondays;
- historical range: 2022-01-03 through 2024-12-23.

Logical dimensions:

- `DimProduct`
- `DimSite`
- `DimSupplier`
- `DimDate`

Logical facts:

- `FactInventoryWeekly`
- `FactPurchaseOrders`
- `FactQualityIncidents`

Validated continuous Date dimension range:

`2022-01-03` through `2025-04-14`

Current maximum Inventory Reporting Date:

`2024-12-23`

## Phase 2 Progress

Phase 2 — Workbook Foundation is now in progress.

Current work:

- create the first physical `ProcureFlow.xlsm` workbook;
- establish the approved worksheet architecture;
- implement the workbook navigation foundation;
- establish configuration structure;
- implement initial Data Validation and defined names;
- establish workbook visual and protection conventions;
- prepare the workbook for Phase 3 Power Query integration.

No Phase 2 workbook functionality is considered implemented until it is created and validated with actual Excel evidence.

## Official Dataset

Aerospace Supply Chain Performance & Forecasting

Official local source files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

The files are stored locally in:

`data/raw/`

and are excluded from Git.

## Known Issues

None currently classified as critical.

Known source characteristic:

`shelf_life_days` is nullable in 274 of 300 Product records and is intentionally treated as an optional attribute.

## Relevant Decisions

Confirmed decisions through Phase 0:

`DEC-001` through `DEC-041`

Confirmed during Phase 1:

`DEC-042` through `DEC-045`

No decisions are currently marked SUPERSEDED.

See:

`docs/DECISIONS.md`

## Next Immediate Step

Create the first physical `workbook/ProcureFlow.xlsm` workbook and establish the approved Phase 2 worksheet foundation.

## Next Phase

Phase 3 — Power Query Pipeline

Status: NOT STARTED

Target Version: `v0.4.0`

Phase 3 must not begin until Phase 2 satisfies its exit criteria and GitHub gate.
