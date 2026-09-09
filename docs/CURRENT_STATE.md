# ProcureFlow — Current State

## Project Status

IN DEVELOPMENT

## Current Version

v0.4.1

## Current Phase

Phase 4 — Operational Model

Status: IN PROGRESS

Target Version: `v0.5.0`

Phase Branch:

`phase/04-operational-model`

## Last Completed Phase

Phase 3 — Power Query Pipeline

Phase Completion Version: `v0.4.0`

Current Corrective Version: `v0.4.1`

GitHub Pull Request:

`#4 — Phase 3 — Power Query Pipeline`

Pull Request Status:

MERGED

Merge Commit:

`540aa99`

Phase Tag:

`phase-3-complete`

Version Tag:

`v0.4.0`
## Implemented

ProcureFlow currently includes the completed Project Design, Data Design, Workbook Foundation and Power Query Pipeline baselines.

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

- operational replenishment model;
- supplier-performance calculations;
- business formulas;
- formal Quality Control system;
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

## Phase 3 Completion

Phase 3 — Power Query Pipeline is COMPLETED.

Validated Phase 3 outcomes include:

- workbook-relative Power Query source-path resolution;
- four `src_*` source queries;
- four `stg_*` staging queries;
- four dimension queries;
- three fact queries;
- explicit data typing and objective transformations;
- seven final structured Excel Tables;
- physical continuous Date dimension;
- Connection Only intermediate queries;
- versioned Power Query M source;
- successful full Refresh;
- 17 Phase 3 Quality Control checks with zero exceptions;
- Query Dependencies implementation evidence;
- formal Phase 3 testing evidence.

Phase 3 Pull Request:

`#4 — Phase 3 — Power Query Pipeline`

Merge Commit:

`540aa99`

Phase Tag:

`phase-3-complete`

Version Tag:

`v0.4.0`

Phase 3 completed without introducing Phase 4 operational-model logic prematurely.
## Phase 4 Progress

Phase 4 — Operational Model is now in progress.
Confirmed Phase 4 design:

- `DEC-051 — Inventory Snapshot Date`;
- operational grain fixed at 1 Product × 1 Site;
- expected `tblReplenishment` population fixed at 1,800 rows;
- initial structural column set approved;
- Reporting Date separated from weekly Inventory Snapshot Date;
- Phase 4 structural inputs separated explicitly from Phase 5 business formulas.

Current work:

- define the Product × Site operational grain;
- design the `tblReplenishment` structural model;
- design the `tblSupplierPerformance` structural model;
- expose Reporting Date-dependent operational inputs;
- expose current inventory, blocked stock and backorders;
- prepare relevant historical-demand inputs;
- prepare open Purchase Order inputs;
- prepare Supplier and Lead-Time inputs;
- reduce repeated scanning of large historical tables where appropriate;
- validate operational populations and structural integrity;
- document Phase 4 implementation and test evidence.

Phase 4 establishes the operational structure required by Phase 5.

Safety Stock, Reorder Point, Inventory Position, Target Stock, Recommended Order Quantity, Inventory Status and final supplier-performance business formulas remain assigned to Phase 5 and are not considered implemented during Phase 4.
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

## Phase 4 Replenishment Foundation Implemented

[IMPLEMENTED]

The initial physical `tblReplenishment` operational structure has been created and validated in:

`20_CALC_Replenishment`

Implemented grain:

1 Product × 1 Site

Validated population:

- Products: 300
- Sites: 6
- expected Product × Site rows: 1,800
- actual `tblReplenishment` rows: 1,800
- unique `ProductSiteKey` values: 1,800

Implemented structural and source-supported fields:

- `ProductSiteKey`
- `ProductID`
- `SiteID`
- `PartFamily`
- `CriticalityClass`
- `PrimarySupplierID`
- `SupplierRiskClass`
- `UnitCost`
- `MasterLeadTimeDays`
- `ReportingDate`
- `InventorySnapshotDate`
- `OnHandQty`
- `BlockedQty`
- `BackorderQty`

Validated master-data results:

- master-attribute errors: 0
- distinct Suppliers exposed: 40
- each Product appears once per Site

Validated Reporting Date behavior:

- `ReportingDate` is driven by `cfg_ReportingDate`
- `InventorySnapshotDate` resolves through `tblDate`
- DEC-051 was tested with a non-Monday Reporting Date
- Reporting Date `2024-12-18` correctly resolved Inventory Snapshot Date `2024-12-16`
- Reporting Date was restored to `2024-12-23`

Validated inventory snapshot results:

- source rows at current snapshot: 1,800
- matched Product × Site snapshot rows: 1,800
- On-Hand total reconciliation: PASS
- Blocked total reconciliation: PASS
- Backorder total reconciliation: PASS
- rows where Blocked Quantity exceeds On-Hand Quantity: 0

Phase 5 business formulas remain unimplemented.
## Phase 4 Supplier Performance Foundation Implemented

[IMPLEMENTED]

The structural `tblSupplierPerformance` model has been created and validated in:

`21_CALC_SupplierPerformance`

Implemented grain:

1 row = 1 Supplier

Validated population:

- expected Suppliers: 40
- actual `tblSupplierPerformance` rows: 40
- unique `SupplierID` values: 40

Implemented Phase 4 fields:

- `SupplierID`
- `SupplierRiskClass`
- `ReportingDate`

Phase 5 metric columns have been structurally reserved but remain without business formulas.

Validated results:

- Supplier population reconciliation: PASS
- Supplier master-data errors: 0
- distinct Reporting Dates: 1

No aggregate Supplier Score has been implemented.
## Next Immediate Step

Review the complete Phase 4 operational model against its roadmap scope and identify any remaining structural inputs, documentation or validation required before Phase 4 closeout.

Phase 5 business formulas must remain unimplemented until the Phase 4 exit criteria are satisfied.
## Next Phase

Phase 5 — Business Logic & Advanced Formulas

Status: NOT STARTED

Target Version: `v0.6.0`

Phase 5 must not begin until Phase 4 satisfies its exit criteria and GitHub gate.