# ProcureFlow — Phase 3 Closeout

## Phase

Phase 3 — Power Query Pipeline

## Status

TECHNICAL WORK COMPLETE — GITHUB GATE PENDING

## Target Version

`v0.4.0`

## Objective

Build a reproducible Power Query ingestion and preparation pipeline from the official raw source files to structured ProcureFlow Excel Tables.

## Completion Summary

Phase 3 completed the technical Power Query pipeline required before the ProcureFlow operational model can be implemented.

The implemented flow is:

Source CSV files
→ `src_*`
→ `stg_*`
→ `dim_*` / `fact_*`
→ structured Excel Tables

The pipeline was implemented and validated in Microsoft Excel 365 Desktop for Windows.

Implemented Phase 3 capabilities include:

- workbook-relative raw-data source resolution;
- four source-access queries;
- four staging queries;
- explicit data typing;
- objective Boolean standardization;
- Product, Site, Supplier and Date dimensions;
- Inventory, Purchase Order and Quality fact outputs;
- objective Purchase Order derived fields;
- Connection Only intermediate-query behavior;
- seven final structured Excel Tables;
- physical Date dimension;
- full Refresh validation;
- workbook Quality Control checks;
- versioned Power Query M source;
- Power Query dependency evidence;
- formal Phase 3 testing evidence.

No operational replenishment formulas, supplier-performance calculations, PivotTables, VBA automation, final reporting logic or dashboard logic were implemented during Phase 3.

## Source Files

Official raw sources:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

Source location:

`data/raw/`

Raw source files remain external to the workbook and excluded from Git.

No manual source cleansing was required.

## Source-Path Strategy

ProcureFlow uses a workbook-relative source-path strategy.

`01_CONFIG` derives the raw-data location from the saved workbook path.

The workbook-scoped Defined Name:

`cfg_RawDataFolder`

exposes the resolved raw-data folder.

Each `src_*` query reads this Defined Name through `Excel.CurrentWorkbook()`.

The production M code therefore does not contain a machine-specific hard-coded raw-data path.

The portability contract requires the approved repository relationship between:

`workbook/`

and:

`data/raw/`

to remain intact.

## Power Query Language

The Power Query implementation is written in the Power Query M language.

M is separate from VBA.

During Phase 3, M is responsible for ingestion and reproducible data preparation.

VBA remains assigned to a later automation phase.

## Implemented Query Inventory

### Source Queries

- `src_PartsMaster`
- `src_SupplyChainHistory`
- `src_PurchaseOrders`
- `src_QualityIncidents`

Responsibilities:

- resolve the configured source path;
- read the approved CSV;
- parse CSV content;
- promote source headers.

Load behavior:

Connection Only.

### Staging Queries

- `stg_Products`
- `stg_InventoryHistory`
- `stg_PurchaseOrders`
- `stg_QualityIncidents`

Responsibilities include:

- logical field naming;
- explicit data types;
- objective representation standardization;
- preservation of valid source nullability.

Load behavior:

Connection Only.

### Dimension Queries

- `dim_Product`
- `dim_Site`
- `dim_Supplier`
- `dim_Date`

### Fact Queries

- `fact_InventoryWeekly`
- `fact_PurchaseOrders`
- `fact_QualityIncidents`

## Purchase Order Derived Fields

`fact_PurchaseOrders` implements the objective historical fields:

- `PromisedLeadTimeDays`
- `ActualLeadTimeDays`
- `IsLateReceipt`
- `IsPartialReceipt`

These values are reproducible properties of the Purchase Order record.

`IsOpenPO` is intentionally not persisted in Power Query because its meaning depends on the configurable Reporting Date and belongs to the later operational Excel model.

## Supplier Risk Ownership

`SupplierRiskClass` is sourced physically from `parts_master.csv` but is modeled logically in `dim_Supplier`.

`dim_Product` retains `PrimarySupplierID` but does not duplicate Supplier Risk as a Product-owned attribute.

This implements the previously validated Supplier → Supplier Risk Class dependency.

## Date Dimension

`dim_Date` is implemented at daily grain.

Current validated range:

`2022-01-03` through `2025-04-14`

Rows:

1,198

Implemented calendar attributes:

- `Date`
- `Year`
- `Quarter`
- `MonthNumber`
- `MonthName`
- `YearMonth`
- `ISOYear`
- `ISOWeekNumber`
- `ISOYearWeek`
- `WeekStartDate`
- `DayOfWeekNumber`
- `DayName`

ISO week semantics use Monday as the beginning of the week.

The Date dimension extends beyond the maximum Inventory History date because procurement and quality dates extend farther into the future.

This does not change the valid maximum Inventory Reporting Date of:

`2024-12-23`

## Final Structured Excel Tables

| Query | Worksheet | Excel Table | Rows |
|---|---|---|---:|
| `dim_Product` | `10_DATA_Products` | `tblProducts` | 300 |
| `dim_Site` | `11_DATA_Sites` | `tblSites` | 6 |
| `dim_Supplier` | `12_DATA_Suppliers` | `tblSuppliers` | 40 |
| `fact_InventoryWeekly` | `13_DATA_Inventory` | `tblInventoryHistory` | 280,800 |
| `fact_PurchaseOrders` | `14_DATA_PurchaseOrders` | `tblPurchaseOrders` | 29,666 |
| `fact_QualityIncidents` | `15_DATA_Quality` | `tblQualityIncidents` | 368 |
| `dim_Date` | `16_DATA_Date` | `tblDate` | 1,198 |

No Phase 3 query is loaded to the Excel Data Model.

## DATA Worksheet Convention

Final DATA worksheets use:

- row 1 for the worksheet title;
- row 2 for HOME navigation;
- row 3 for the structured Excel Table header;
- row 4 onward for refreshed data.

Phase 2 implementation placeholders were removed when the corresponding Phase 3 outputs became real.

Technical date displays use:

`yyyy-mm-dd`

Boolean fields remain logical Power Query values and display as:

`TRUE` / `FALSE`

in Excel.

## Quality Control Validation

`02_CONTROL` now contains Phase 3 technical pipeline controls using the existing ProcureFlow Quality Control framework.

Implemented controls:

`PQ-001` through `PQ-017`

Validated areas include:

- final row counts;
- distinct key counts;
- Date minimum;
- Date maximum;
- Date continuity;
- late-receipt count;
- partial-receipt count;
- nullable Shelf Life preservation.

All 17 controls returned:

`PASS`

All Exception Count values returned:

`0`

Conditional Formatting visually distinguishes PASS and FAIL while retaining textual status labels.

This Phase 3 technical validation does not represent the complete Phase 6 business Quality Control implementation.

## Reconciliation Evidence

Validated final populations:

- Products: 300
- Sites: 6
- Suppliers: 40
- Inventory History: 280,800
- Purchase Orders: 29,666
- Quality Incidents: 368
- Date Dimension: 1,198

Validated source behaviors preserved by Power Query:

- blank Product Shelf Life values: 274
- late Purchase Orders: 16,568
- partial Purchase Orders: 3,355

No valid source behavior was silently removed through cleansing.

## Refresh Validation

A full Excel `Refresh All` was executed after implementation.

Observed results:

- source queries refreshed successfully;
- staging queries refreshed successfully;
- final dimension queries refreshed successfully;
- final fact queries refreshed successfully;
- final Excel Tables remained loaded;
- Phase 3 Quality Control results remained PASS;
- no critical Power Query error remained.

Status:

PASS

## Power Query Source Versioning

The executable Power Query implementation remains embedded in:

`workbook/ProcureFlow.xlsm`

Because the workbook is a binary Git artifact, the corresponding M code is also maintained as text under:

`power-query/`

Repository structure:

- `power-query/src/`
- `power-query/stg/`
- `power-query/dim/`
- `power-query/fact/`

Current `.pq` query source count:

15

The `.pq` files enable:

- Git line-by-line diffs;
- direct GitHub inspection;
- code review;
- technical traceability.

Workbook queries and corresponding `.pq` source must remain synchronized when M code changes.

## Query Dependency Evidence

Power Query Query Dependencies provides implementation evidence of the actual query lineage.

Screenshot:

`screenshots/phase-03-query-dependencies.png`

The captured graph shows the implemented:

`src_*`
→ `stg_*`
→ `dim_*` / `fact_*`

dependency structure.

The file-source nodes shown by Power Query display absolute paths resolved at runtime on the development machine.

Those paths are not hard-coded in the production `src_*` M queries.

The raw-data location is derived through `cfg_RawDataFolder`.

## Phase 3 Decisions

Confirmed during Phase 3:

- `DEC-049` — Workbook-Relative Power Query Source Path
- `DEC-050` — Power Query M Source Versioning

Earlier confirmed decisions remain applicable.

No decision is considered superseded unless explicitly documented.

## Testing Evidence

Formal Phase 3 validation is documented in:

`docs/TESTING.md`

Current Phase 3 testing result:

PASS

No critical Power Query pipeline defect is currently known.

## Exit Criteria Review

Phase 3 technical exit criteria have been validated.

### Valid source files refresh successfully

Status:

PASS

### Expected final tables load correctly

Status:

PASS

Seven required final structured Excel Tables are implemented and loaded.

### Row counts reconcile

Status:

PASS

Final populations reconcile with the validated Phase 1 source baseline.

### No critical Power Query errors remain

Status:

PASS

No unresolved critical Power Query error is currently known.

### No manual cleansing of raw files is required

Status:

PASS

Raw CSV files remain immutable and all required preparation occurs reproducibly through Power Query.

## Technical Phase Result

Technical Phase 3 work is complete.

Phase 3 must not yet be marked formally completed because the GitHub publication gate remains pending.

## Git Evidence

Phase branch:

`phase/03-power-query-pipeline`

Target branch:

`main`

Target version:

`v0.4.0`

Phase Pull Request:

PENDING

Merge commit:

PENDING

Phase completion tag:

`phase-3-complete` — PENDING

Version tag:

`v0.4.0` — PENDING

## GitHub Gate

Status:

PENDING

Required remaining steps:

1. commit the Phase 3 closeout;
2. publish the final Phase 3 branch state;
3. create the Phase 3 Pull Request;
4. review the Pull Request diff and evidence;
5. correct any review findings;
6. merge the Pull Request into `main`;
7. synchronize local `main`;
8. finalize canonical Phase 3 completion status;
9. publish `phase-3-complete`;
10. publish `v0.4.0`.

Phase 3 becomes formally COMPLETED only after this gate is satisfied.

## Next Phase

Phase 4 — Operational Model

Status:

NOT STARTED

Target Version:

`v0.5.0`

Phase 4 must not begin until Phase 3 is formally completed and its GitHub gate is closed.

## Phase 4 Handoff

After Phase 3 is formally completed, Phase 4 must begin by reviewing at minimum:

1. `docs/CURRENT_STATE.md`
2. `docs/phases/PHASE_03_CLOSEOUT.md`
3. `docs/ROADMAP.md`
4. `docs/PROJECT_SPEC.md`
5. `docs/ARCHITECTURE.md`
6. `docs/DECISIONS.md`
7. `docs/DATA_DICTIONARY.md`
8. `docs/TESTING.md`
9. `power-query/README.md`

GitHub remains the authoritative project handoff mechanism.