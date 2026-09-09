# ProcureFlow — Current State

## Project Status

IN DEVELOPMENT

## Current Version

v0.3.0

## Current Phase

Phase 3 — Power Query Pipeline

Status: IN PROGRESS

Target Version: `v0.4.0`

Phase Branch:

`phase/03-power-query-pipeline`

## Last Completed Phase

Phase 2 — Workbook Foundation

Version: `v0.3.0`

GitHub Pull Request:

`#3 — Phase 2 — Workbook Foundation`

Pull Request Status:

MERGED

Merge Commit:

`bf9dcb8`

Phase Tag:

`phase-2-complete`

Version Tag:

`v0.3.0`

## Implemented

ProcureFlow currently includes the completed project-design, data-design and Workbook Foundation baselines.

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
- Phase 1 validation evidence;
- physical `workbook/ProcureFlow.xlsm`;
- approved 17-sheet physical workbook architecture;
- `00_HOME` navigation foundation;
- workbook-wide return navigation;
- `01_CONFIG` business-parameter structure;
- seven Data Validation-controlled configuration inputs;
- eight workbook-scoped `cfg_*` Defined Names: seven business configuration names plus technical `cfg_RawDataFolder`;
- Reporting Date exposure;
- `02_CONTROL` structural foundation;
- remaining later-phase technical-sheet placeholders;
- implemented physical `16_DATA_Date` worksheet with `tblDate`;
- approved workbook visual design system;
- layer-based worksheet tab colors;
- editable-input visual convention;
- Aptos 11 workbook Normal style;
- incremental worksheet protection for `01_CONFIG`;
- Phase 2 Workbook Foundation validation evidence.
- workbook-relative Power Query raw-data path through `cfg_RawDataFolder`;
- four `src_*` source-access queries;
- four `stg_*` staging queries;
- explicit Power Query data typing and objective Boolean conversion;
- four final dimension queries: Product, Site, Supplier and Date;
- three final fact queries: Inventory Weekly, Purchase Orders and Quality Incidents;
- objective Purchase Order fields `PromisedLeadTimeDays`, `ActualLeadTimeDays`, `IsLateReceipt` and `IsPartialReceipt`;
- seven final structured Excel Tables;
- continuous 1,198-row Date dimension from `2022-01-03` through `2025-04-14`;
- Connection Only behavior for source and staging queries;
- Power Query M source mirrored as 15 versionable `.pq` files under `power-query/`;
- full Power Query Refresh validation;
- 17 Phase 3 workbook Quality Control checks with zero exceptions;
- Phase 3 Power Query test evidence.

## Not Yet Implemented

The following remain assigned to later roadmap phases:

- Power Query ingestion and preparation;
- final structured DATA tables;
- physical Date-dimension data;
- operational replenishment model;
- supplier-performance calculations;
- business formulas;
- Quality Control logic;
- PivotTables;
- PivotCharts;
- Slicers;
- Timelines;
- VBA;
- macros;
- refresh automation;
- replenishment-report logic;
- management-dashboard logic.

No future-phase component is considered implemented without actual implementation and validation evidence.

## Phase 1 Completion

Phase 1 established the validated logical data model.

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

## Phase 2 Completion

Phase 2 established the first physical ProcureFlow workbook and completed the Workbook Foundation required for Power Query integration.

Validated Phase 2 outcomes include:

- macro-enabled physical workbook;
- 17-sheet physical architecture;
- layer and sheet-order conventions;
- base navigation;
- configuration architecture;
- controlled editable inputs;
- Data Validation;
- workbook-scoped configuration names;
- visual system;
- workbook typography;
- protection strategy;
- future-phase implementation boundaries.

Phase 2 Pull Request:

`#3 — Phase 2 — Workbook Foundation`

Merge Commit:

`bf9dcb8`

Phase 2 completed without introducing Phase 3 or later business functionality prematurely.

## Phase 3 Progress

Phase 3 — Power Query Pipeline is now in progress.

Implemented and validated so far:

- portable Power Query source-path resolution;
- `src_*`, `stg_*`, `dim_*` and `fact_*` query layers;
- explicit types and objective transformations;
- seven final structured Excel Tables;
- physical Date dimension;
- Connection Only intermediate queries;
- versioned Power Query M source;
- full Refresh validation;
- 17 workbook Phase 3 validation controls;
- formal Phase 3 testing evidence.

Remaining Phase 3 work:

- review Pull Request `#4 — Phase 3 — Power Query Pipeline`;
- correct any review findings;
- merge the approved Phase 3 state into `main`;
- finalize canonical Phase 3 completion status;
- publish the Phase 3 and version tags.

## Official Dataset

Aerospace Supply Chain Performance & Forecasting

Official local source files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

The files are stored locally in:

`data/raw/`

and remain excluded from Git.

## Known Issues

None currently classified as critical.

Known source characteristic:

`shelf_life_days` is nullable in 274 of 300 Product records and is intentionally treated as an optional attribute.

## Relevant Decisions

Confirmed through Phase 0:

`DEC-001` through `DEC-041`

Confirmed during Phase 1:

`DEC-042` through `DEC-045`

Confirmed during Phase 2:

`DEC-046` through `DEC-048`

Confirmed during Phase 3:

`DEC-049` through `DEC-050`

No decisions are currently marked SUPERSEDED.

See:

`docs/DECISIONS.md`

## Next Immediate Step

Complete the review of Pull Request `#4 — Phase 3 — Power Query Pipeline` and correct any review findings.

After approval, merge Phase 3 into `main`, finalize the canonical completion state and publish the required Phase 3 and version tags.

## Next Phase

Phase 4 — Operational Model

Status: NOT STARTED

Target Version: `v0.5.0`

Phase 4 must not begin until Phase 3 satisfies its exit criteria and GitHub gate.
