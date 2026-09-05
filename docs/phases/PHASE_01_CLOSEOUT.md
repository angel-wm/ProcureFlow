# ProcureFlow — Phase 1 Closeout

## Phase

Phase 1 — Data Design

## Status

READY FOR GITHUB GATE

## Target Version

`v0.2.0`

## Objective

Transform the selected Aerospace Supply Chain Performance & Forecasting dataset into a formally validated data specification before workbook construction begins.

Phase 1 establishes the validated grains, keys, relationships, logical data types, dimensions, facts, Date dimension design, source-to-target mappings and source-supported data-quality rules required by subsequent ProcureFlow implementation phases.

## Completion Summary

Phase 1 completed the formal data-design work required before workbook construction.

The four official source files were inspected and validated directly from immutable local CSV files using PowerShell.

The phase established:

- validated source schemas;
- validated row counts;
- validated grains;
- Primary Keys and composite keys;
- referential integrity;
- source nullability;
- categorical domains;
- observed numeric ranges;
- date ranges;
- logical data types;
- Supplier/Product relationships;
- source-to-target mappings;
- logical dimensions and facts;
- Date dimension requirements;
- weekly-calendar semantics;
- source-supported quality rules;
- formal Phase 1 test evidence.

No Excel workbook functionality or Power Query implementation was performed during this phase.

## Official Dataset

Aerospace Supply Chain Performance & Forecasting

Official source files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

Raw source files are stored locally under:

`data/raw/`

and excluded from Git.

## Validated Source Population

### Products

Source:

`parts_master.csv`

Rows:

300

Validated grain:

1 row = 1 Product

Primary Key:

`part_id`

### Inventory History

Source:

`supply_chain_history.csv`

Rows:

280,800

Validated grain:

1 Product × 1 Site × 1 Week

Composite Key:

`date + site_id + part_id`

### Purchase Orders

Source:

`purchase_orders.csv`

Rows:

29,666

Validated grain:

1 row = 1 Purchase Order record for 1 Product

Primary Key:

`po_id`

### Quality Incidents

Source:

`quality_incidents.csv`

Rows:

368

Validated grain:

1 row = 1 Quality Incident

Primary Key:

`incident_id`

### Entity Populations

- Products: 300
- Sites: 6
- Suppliers: 40
- Product × Site combinations: 1,800
- historical weekly periods: 156

## Completed Deliverables

- field-level source-data dictionary;
- validated source grain definitions;
- Primary Key specification;
- composite-key specification;
- Foreign Key and referential-integrity specification;
- required versus optional field classification;
- null-behavior documentation;
- categorical-domain documentation;
- observed numeric-range documentation;
- source date-range validation;
- logical data-type specification;
- Product dimension specification;
- Site dimension specification;
- Supplier dimension specification;
- Date dimension specification;
- Inventory Weekly fact specification;
- Purchase Order fact specification;
- Quality Incident fact specification;
- source-to-target mapping;
- logical relationship specification;
- logical relationship diagram;
- Purchase Order derived-field specification;
- data-quality rules supported by source semantics;
- Phase 1 test evidence;
- material Phase 1 decision records.

## Canonical Documents Updated

- `docs/CURRENT_STATE.md`
- `docs/ROADMAP.md`
- `docs/DATA_DICTIONARY.md`
- `docs/TESTING.md`
- `docs/ARCHITECTURE.md`
- `docs/DECISIONS.md`
- `docs/phases/PHASE_01_CLOSEOUT.md`

## Logical Data Model

### Dimensions

- `DimProduct`
- `DimSite`
- `DimSupplier`
- `DimDate`

### Facts

- `FactInventoryWeekly`
- `FactPurchaseOrders`
- `FactQualityIncidents`

These are logical data-design structures.

They are not yet implemented as Power Query queries, Excel Tables or Power Pivot objects.

## Product Dimension

Grain:

1 row = 1 Product

Primary Key:

`ProductID`

Core logical attributes:

- ProductID
- PartFamily
- CriticalityClass
- UnitCost
- MasterLeadTimeDays
- PrimarySupplierID
- IsRepairable
- ShelfLifeDays

`supplier_risk_class` is physically sourced from `parts_master.csv` but logically belongs to `DimSupplier`.

## Supplier Dimension

Grain:

1 row = 1 Supplier

Primary Key:

`SupplierID`

Validated current population:

40 Suppliers

Logical attributes:

- SupplierID
- SupplierRiskClass

Validated dependency:

1 Supplier → 1 Supplier Risk Class

Validated relationship:

1 Supplier → many Products

Observed Products per Supplier:

4 through 14

## Site Dimension

Grain:

1 row = 1 Site

Primary Key:

`SiteID`

Validated current population:

6 Sites

No unsupported Site names, locations, regions or other attributes were invented.

## Date Dimension

Grain:

1 row = 1 calendar date

Primary Key:

`Date`

Validated current continuous range:

`2022-01-03` through `2025-04-14`

Required current row count:

1,198

Weekly convention:

ISO 8601

Week start:

Monday

Required logical calendar attributes include:

- Date
- Year
- Quarter
- MonthNumber
- MonthName
- YearMonth
- ISOYear
- ISOWeekNumber
- ISOYearWeek
- WeekStartDate
- DayOfWeekNumber
- DayName

The Date dimension range does not determine the maximum valid Inventory Reporting Date.

## Inventory Weekly Fact

Grain:

1 Product × 1 Site × 1 Week

Validated structure:

- 300 Products;
- 6 Sites;
- 1,800 Product × Site combinations;
- 156 weeks per Product × Site;
- 280,800 rows.

Historical range:

`2022-01-03` through `2024-12-23`

All 156 historical dates are Mondays.

All consecutive historical dates are exactly 7 days apart.

No Product × Site historical week is missing from the validated source range.

Maximum current Inventory Reporting Date:

`2024-12-23`

## Purchase Order Fact

Grain:

1 row = 1 Purchase Order record for 1 Product

Primary Key:

`PurchaseOrderID`

Validated rows:

29,666

The source does not support a traditional Purchase Order Header → Purchase Order Lines model.

That structure will not be invented.

Objective source-supported derived fields include:

- PromisedLeadTimeDays
- ActualLeadTimeDays
- IsLateReceipt
- IsPartialReceipt

Reporting-Date-dependent state such as `IsOpenPO` must remain dynamic rather than being stored as permanent transaction state.

## Purchase Order Validation

Observed quantities:

- Ordered Qty: 1 through 263
- Received Qty: 1 through 263
- full receipts: 26,311
- partial receipts: 3,355
- zero receipts: 0

Chronological validation:

- Promised Date before Order Date: 0
- Receipt Date before Order Date: 0
- Receipt Date after Promised Date: 16,568

Late receipts are valid supplier-performance information and are not source errors.

Promised Lead Time:

10 through 123 days

Actual Lead Time:

8 through 131 days

Observed source-wide Actual Lead Time average:

42.97 days

## Quality Incident Fact

Grain:

1 row = 1 Quality Incident

Primary Key:

`QualityIncidentID`

Validated rows:

368

Observed Scrap Qty:

1 through 13

Negative Scrap Qty:

0

## Referential Integrity

Validated invalid Product references:

- Inventory History: 0
- Purchase Orders: 0
- Quality Incidents: 0

Validated invalid Site references:

- Purchase Orders: 0
- Quality Incidents: 0

Validated Suppliers:

40

Purchase Order Suppliers outside validated Supplier population:

0

Quality Suppliers outside validated Supplier population:

0

Product/Supplier mismatches against Product Primary Supplier:

- Purchase Orders: 0
- Quality Incidents: 0

Blank required relational identifiers:

0

## Nullability

Only one validated source field contains blanks:

`parts_master.shelf_life_days`

Observed:

- non-null: 26
- blank: 274

The field is therefore classified as optional.

No missing Shelf Life value will be fabricated or automatically imputed without a future documented business rule.

## Forecast Validation

Forecast types:

- Baseline
- Adjusted

Baseline rows:

227,059

Adjusted rows:

53,741

Validated relationship:

Baseline → Forecast Uplift = 0

Adjusted → Forecast Uplift <> 0

Violations observed:

0

Adjusted uplift observed range:

-0.30 through 0.42

Forecast Qty fractional rows:

0

Forecast Qty is therefore specified as Whole Number in the current logical design.

## Identifier Formats

Validated current formats:

- Product ID → `P#####`
- Supplier ID → `SUP###`
- Site ID → `SITE##`
- Purchase Order ID → `PO######`
- Quality Incident ID → `QI######`

Invalid observed identifiers:

0

## Phase 1 Decisions

Confirmed during Phase 1:

### DEC-042 — Supplier Risk Logical Ownership

Supplier Risk belongs logically to `DimSupplier`.

### DEC-043 — Inventory Weekly Date Semantics

Inventory History `date` represents Monday-based `WeekStartDate`.

### DEC-044 — Date Dimension Coverage and Inventory Reporting Boundary

`DimDate` covers all required source dates continuously, while Inventory Reporting Date remains bounded by available Inventory History.

### DEC-045 — ISO Weekly Calendar Convention

ProcureFlow uses ISO 8601 semantics for weekly calendar attributes.

Phase 0 decisions remain confirmed through:

`DEC-041`

No decisions are currently marked SUPERSEDED.

## Testing Evidence

Phase 1 source-data validation is documented in:

`docs/TESTING.md`

Validated areas include:

- source availability;
- source schemas;
- source row counts;
- Primary Keys;
- composite keys;
- entity populations;
- referential integrity;
- blank identifiers;
- Product/Supplier consistency;
- Supplier Risk dependency;
- Supplier/Product cardinality;
- Inventory weekly completeness;
- nullability;
- categorical domains;
- identifier formats;
- Product numeric validation;
- Inventory numeric validation;
- forecast semantics;
- Purchase Order quantities;
- Purchase Order chronology;
- Purchase Order Lead Times;
- Quality Incident quantities;
- Date-domain coverage;
- candidate logical data types.

No critical source-data ambiguity remains unresolved.

## Known Source Limitations

The following source limitations remain part of the approved baseline:

- Purchase Orders do not use a traditional header/line structure;
- each Product effectively uses one Primary Supplier in the current source;
- no independent rich Supplier master exists;
- no independent rich Site master exists;
- MOQ and order-multiple data are not available;
- Site descriptive attributes are not available;
- the dataset is synthetic;
- `shelf_life_days` is sparsely populated;
- source limitations must not be filled by invented business data.

## Implementation Status

No workbook functionality has been implemented during Phase 1.

Specifically, none of the following is considered implemented:

- `ProcureFlow.xlsm`;
- Power Query queries;
- Excel data tables;
- physical `DimDate`;
- operational formulas;
- Dynamic Array outputs;
- Quality Control worksheet;
- PivotTables;
- PivotCharts;
- VBA modules;
- macros;
- automated refresh;
- replenishment calculations;
- supplier-performance calculations;
- operational reports;
- management dashboard.

These belong to later roadmap phases.

## Exit Criteria Review

Phase 1 technical exit criteria have been validated.

Confirmed:

- every required source field is understood and documented;
- source grains are validated;
- required Primary Keys are validated;
- the Inventory composite key is validated;
- Foreign Key relationships are validated;
- dimensions are specified;
- facts are specified;
- source-to-target mappings are documented;
- Date dimension design is documented;
- logical relationships are documented;
- data-quality rules are documented;
- Phase 1 test evidence is recorded;
- no unresolved critical data-design ambiguity remains.

Technical Phase 1 work is complete.

Formal phase completion remains subject to the GitHub Gate.

## Git Evidence

Phase branch:

`phase/01-data-design`

Target branch:

`main`

Phase Pull Request:

PENDING

Merge commit:

PENDING

Phase completion tag:

`phase-1-complete` — PENDING

Version tag:

`v0.2.0` — PENDING

The phase must not be marked COMPLETED until the required GitHub workflow is finished.

## GitHub Gate

Pending steps:

1. commit this Phase 1 closeout;
2. confirm a clean Phase 1 branch;
3. publish `phase/01-data-design`;
4. create the Phase 1 Pull Request;
5. review the branch diff and phase evidence;
6. merge the Pull Request into `main`;
7. synchronize local `main`;
8. finalize canonical phase status;
9. create the Phase 1 completion tag;
10. create the `v0.2.0` version tag;
11. publish the final tags.

## Next Phase

Phase 2 — Workbook Foundation

Status:

NOT STARTED

Target Version:

`v0.3.0`

Phase 2 must not begin until Phase 1 completes its GitHub Gate.

## Phase 2 Handoff

When Phase 1 is formally completed, the next phase must begin by reviewing:

1. `docs/CURRENT_STATE.md`
2. `docs/phases/PHASE_01_CLOSEOUT.md`
3. `docs/ROADMAP.md`
4. `docs/PROJECT_SPEC.md`
5. `docs/ARCHITECTURE.md`
6. `docs/DECISIONS.md`
7. `docs/DATA_DICTIONARY.md`
8. `docs/TESTING.md`

The GitHub repository remains the authoritative handoff mechanism.

No reconstruction of the Phase 1 conversation should be required.
