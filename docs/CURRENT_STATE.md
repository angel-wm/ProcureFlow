# ProcureFlow — Current State

## Project Status

IN DEVELOPMENT

## Current Version

v0.2.0

## Current Phase

Phase 1 — Data Design

Status: COMPLETED

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

Begin Phase 2 — Workbook Foundation after creating the Phase 1 completion and version tags.

Phase 2 must begin by reviewing the canonical Phase 1 handoff documentation.

## Next Phase

Phase 2 — Workbook Foundation

Status: NOT STARTED

Target Version: `v0.3.0`

Phase 2 must not be considered started until work begins on its dedicated phase branch.
