# ProcureFlow — Current State

## Project Status

IN DEVELOPMENT

## Current Version

v0.6.1

## Current Phase

Phase 5 — Business Logic & Advanced Formulas

Status: COMPLETED

Version: `v0.6.0`

Phase Branch:

`phase/05-business-logic-advanced-formulas`

## Last Completed Phase

Phase 5 — Business Logic & Advanced Formulas

Phase Completion Version: `v0.6.0`

GitHub Pull Request:

`#6 — Phase 5 — Business Logic & Advanced Formulas`

Pull Request Status:

MERGED

Merge Commit:

`43efd6a76bd9f4ce54093121289486bafd32e2bf`

Phase Tag:

`phase-5-complete`

Version Tag:

`v0.6.0`

## Implemented

ProcureFlow currently includes the completed Project Design, Data Design, Workbook Foundation, Power Query Pipeline, Operational Model and Business Logic baselines.

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
## Phase 4 Completion

Phase 4 — Operational Model is COMPLETED.

Validated Phase 4 outcomes include:

- Product × Site operational grain implemented at exactly 1,800 rows;
- `tblReplenishment` implemented and structurally validated;
- Supplier operational grain implemented at exactly 40 rows;
- `tblSupplierPerformance` implemented and structurally validated;
- Product and Supplier master attributes exposed;
- Reporting Date relationships implemented;
- weekly Inventory Snapshot Date resolution implemented under `DEC-051`;
- On-Hand, Blocked and Backorder quantities reconciled;
- completed historical-demand window implemented under `DEC-052`;
- master and historical Lead-Time inputs prepared;
- Purchase Order inputs prepared for later Reporting-Date-dependent business logic;
- production and validation formulas versioned in `docs/FORMULAS.md`;
- full workbook Refresh validation completed successfully;
- Phase 4 technical exit criteria satisfied;
- Pull Request `#5 — Phase 4 — Operational Model` reviewed and merged.

Phase 5 business formulas were intentionally not implemented during Phase 4.
## Phase 5 Completion

Phase 5 — Business Logic & Advanced Formulas is COMPLETED.

Validated Phase 5 outcomes include:

- functioning Product × Site replenishment engine;
- Historical Demand aggregation;
- Average Weekly Demand;
- demand variability;
- Reporting-Date-safe historical Actual Lead Time;
- Lead-Time variability;
- Effective Lead-Time fallback;
- Reporting-Date-dependent Open PO Quantity;
- Available Stock;
- configurable Service Levels;
- Safety Stock;
- Reorder Point;
- Inventory Position;
- Target Stock;
- Recommended Order Quantity;
- Inventory Status;
- NoRecentDemand safeguard;
- functioning Supplier Performance calculations;
- On-Time and Late Delivery metrics;
- partial-receipt metrics;
- supplier Lead-Time metrics;
- supplier Quality Incident counts;
- formula documentation;
- manual and global reconciliation evidence;
- successful full workbook Refresh;
- measured full Excel recalculation of 51.4639623 seconds.

Pull Request `#6 — Phase 5 — Business Logic & Advanced Formulas` was merged into `main`.

Phase 5 release version:

`v0.6.0`

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

Confirmed during Phase 4:

`DEC-051` through `DEC-052`

Confirmed during Phase 5:

`DEC-053` through `DEC-055`

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

At the Phase 4 close, Phase 5 business formulas had not yet been implemented. They were subsequently implemented and validated during Phase 5.
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

At the Phase 4 close, Phase 5 metric columns had been structurally reserved but did not yet contain business formulas. Those formulas were subsequently implemented and validated during Phase 5.

Validated results:

- Supplier population reconciliation: PASS
- Supplier master-data errors: 0
- distinct Reporting Dates: 1

No aggregate Supplier Score has been implemented.
## Phase 4 Historical Demand Context Implemented

[IMPLEMENTED]

The historical-demand temporal context required by Phase 5 has been implemented and validated in `tblReplenishment`.

Implemented fields:

- `DemandHistoryWeeks`
- `DemandHistoryStartDate`
- `DemandHistoryEndDate`

Current validated values:

- Demand History Weeks: 26
- Demand History Start Date: `2024-06-24`
- Demand History End Date: `2024-12-16`

Validated source coverage:

- distinct historical weeks: 26
- Product × Site combinations per week: 1,800
- total Inventory History rows in configured window: 46,800

The demand-history window uses only completed weekly periods preceding the Inventory Snapshot Date in accordance with `DEC-052`.

At the Phase 4 close, no Phase 5 demand aggregation or statistical business calculation had yet been implemented. Those calculations were subsequently implemented and validated during Phase 5.
## Phase 4 Technical Exit Criteria Satisfied

[CONFIRMADO]

A full workbook `Refresh All` was executed after the Phase 4 operational-model implementation.

Refresh result:

PASS

Post-refresh operational validation:

- `tblReplenishment` rows: 1,800
- unique `ProductSiteKey` values: 1,800
- replenishment master/input formula errors: 0
- Reporting Date relationship: PASS
- historical-demand source rows in configured window: 46,800
- `tblSupplierPerformance` rows: 40
- unique Supplier IDs: 40
- supplier structural formula errors: 0

The Phase 4 exit criteria are technically satisfied:

- operational grains are correct;
- required Phase 5 inputs are available;
- no unresolved critical structural issue is known;
- the model is ready for Phase 5 business-rule formulas.

Phase 4 is formally COMPLETED. Technical implementation, validation, Pull Request review and merge are complete.
## Next Immediate Step

Begin the dedicated Phase 6 — Quality Control System workflow from the completed `v0.6.0` baseline.

Phase 6 must reconstruct state from GitHub and implement the formal ProcureFlow Quality Control and reconciliation framework.

## Next Phase

Phase 6 — Quality Control System

Status: NOT STARTED

Target Version: `v0.7.0`

## Phase 5 Business Logic Validation

Phase 5 core business logic has been implemented and validated in the workbook.

Validated outputs include:

- Historical Demand calculations
- demand average and variability
- Reporting-Date-safe historical Lead Time
- Lead-Time variability and fallback
- Available Stock
- Reporting-Date-dependent Open PO Quantity
- configurable Service Level
- Safety Stock
- Reorder Point
- Inventory Position
- Target Stock
- Recommended Order Quantity
- Inventory Status
- NoRecentDemand
- Supplier delivery performance
- Supplier Lead-Time performance
- partial-receipt metrics
- supplier quality-incident counts

Validation evidence:

- Phase 5 formula errors: 0
- Inventory Status rows: 1,800
- positive Recommended Order lines: 392
- total Recommended Order Quantity: 2,109 units
- Open PO Quantity: 20,146 units
- NoRecentDemand rows: 0
- Master Lead Time fallback rows: 0
- Refresh All: successful
- measured full Excel recalculation: 51.4639623 seconds

Phase 5 technical implementation, validation, documentation and GitHub Pull Request merge are complete.
