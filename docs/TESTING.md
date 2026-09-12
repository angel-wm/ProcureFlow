# ProcureFlow — Testing

## Document Status

Status: PHASE 6 TECHNICAL TEST EVIDENCE COMPLETE — GITHUB GATE PENDING

Current Development Phase:

Phase 6 — Quality Control System

Current Released Version:

`v0.6.2`

Phase 6 Target Version:

`v0.7.0`

Formal testing will continue throughout later ProcureFlow phases.

This document records validation evidence only for work that has actually been executed.

A PASS in Phase 1 does not imply that Power Query, workbook functionality, formulas, VBA, PivotTables or reporting have been implemented.

---

# 1. Test Status Values

- PASS — executed and validated successfully.
- WARNING — unusual but potentially valid behavior requiring awareness.
- FAIL — executed validation found an invalid condition.
- NOT RUN — not yet tested.
- NOT APPLICABLE — not applicable in the current phase.

---

# 2. Phase 1 Test Environment

Validation method:

PowerShell inspection of immutable local CSV source files.

Source location:

`data/raw/`

Official files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

No source CSV was modified during testing.

---

# 3. Source Availability and Counts

| Test | Expected | Observed | Status |
|---|---:|---:|---|
| Product source available | Yes | Yes | PASS |
| Inventory source available | Yes | Yes | PASS |
| Purchase Order source available | Yes | Yes | PASS |
| Quality source available | Yes | Yes | PASS |
| Product rows | ~300 | 300 | PASS |
| Inventory History rows | ~280,800 | 280,800 | PASS |
| Purchase Order rows | ~30,000 | 29,666 | PASS |
| Quality Incident rows | Hundreds | 368 | PASS |

---

# 4. Source Schema Validation

## parts_master.csv

Validated fields:

- `part_id`
- `part_family`
- `criticality_class`
- `unit_cost`
- `lead_time_days`
- `supplier_id_primary`
- `supplier_risk_class`
- `is_repairable`
- `shelf_life_days`

Field count:

9

Status:

PASS

## supply_chain_history.csv

Validated fields:

- `date`
- `site_id`
- `part_id`
- `planned_maintenance`
- `consumption_qty`
- `on_hand_qty`
- `backorder_qty`
- `blocked_qty`
- `forecast_qty`
- `forecast_type`
- `forecast_uplift_pct`

Field count:

11

Status:

PASS

## purchase_orders.csv

Validated fields:

- `po_id`
- `supplier_id`
- `site_id`
- `part_id`
- `order_date`
- `promised_date`
- `receipt_date`
- `ordered_qty`
- `received_qty`

Field count:

9

Status:

PASS

## quality_incidents.csv

Validated fields:

- `incident_id`
- `incident_date`
- `part_id`
- `supplier_id`
- `site_id`
- `defect_severity`
- `defect_type`
- `scrap_qty`

Field count:

8

Status:

PASS

---

# 5. Key Validation

| Test | Observed | Status |
|---|---:|---|
| Product rows | 300 | PASS |
| Distinct Product IDs | 300 | PASS |
| Blank Product IDs | 0 | PASS |
| Purchase Order rows | 29,666 | PASS |
| Distinct Purchase Order IDs | 29,666 | PASS |
| Blank Purchase Order IDs | 0 | PASS |
| Quality Incident rows | 368 | PASS |
| Distinct Quality Incident IDs | 368 | PASS |
| Blank Quality Incident IDs | 0 | PASS |
| Inventory composite-key rows | 280,800 | PASS |
| Distinct `date + site_id + part_id` keys | 280,800 | PASS |
| Duplicate Inventory composite-key rows | 0 | PASS |

Validated Primary Keys:

- Product → `part_id`
- Purchase Order → `po_id`
- Quality Incident → `incident_id`

Validated Inventory composite key:

`date + site_id + part_id`

---

# 6. Entity Population Validation

| Entity | Observed Count | Status |
|---|---:|---|
| Products | 300 | PASS |
| Sites | 6 | PASS |
| Suppliers | 40 | PASS |
| Product × Site combinations | 1,800 | PASS |

---

# 7. Referential Integrity Tests

## Product References

| Test | Invalid References | Status |
|---|---:|---|
| Inventory → Product | 0 | PASS |
| Purchase Orders → Product | 0 | PASS |
| Quality → Product | 0 | PASS |

## Site References

| Test | Invalid References | Status |
|---|---:|---|
| Purchase Orders → Site | 0 | PASS |
| Quality → Site | 0 | PASS |

## Supplier References

Validated Supplier counts:

- Product source: 40
- Purchase Orders: 40
- Quality: 40

| Test | Invalid Distinct Suppliers | Status |
|---|---:|---|
| PO Suppliers outside Product Supplier population | 0 | PASS |
| Quality Suppliers outside Product Supplier population | 0 | PASS |

## Blank Relationship Identifiers

The following blank counts were validated:

- Inventory Product: 0
- Inventory Site: 0
- Purchase Order Product: 0
- Purchase Order Site: 0
- Purchase Order Supplier: 0
- Quality Product: 0
- Quality Site: 0
- Quality Supplier: 0

Status:

PASS

---

# 8. Product-to-Supplier Consistency

Every Product has one Primary Supplier in `parts_master.csv`.

Purchase Order Supplier was compared with Product Primary Supplier.

Mismatches:

0

Quality Incident Supplier was compared with Product Primary Supplier.

Mismatches:

0

Status:

PASS

---

# 9. Supplier Attribute Dependency

Test:

Determine whether one Supplier can have multiple `supplier_risk_class` values.

Observed:

- Suppliers: 40
- Suppliers with exactly one Risk Class: 40
- Suppliers with multiple Risk Classes: 0

Conclusion:

`supplier_risk_class` functionally depends on Supplier under the validated source.

Status:

PASS

---

# 10. Supplier-to-Product Cardinality

Observed:

- Suppliers: 40
- minimum Products per Supplier: 4
- maximum Products per Supplier: 14
- Suppliers with exactly one Product: 0
- Suppliers with multiple Products: 40

Validated relationship:

1 Supplier → many Products

Status:

PASS

---

# 11. Inventory Grain and Historical Completeness

Observed Product × Site combinations:

1,800

Expected from population:

300 Products × 6 Sites = 1,800

Observed weeks per Product × Site:

- minimum: 156
- maximum: 156

Distinct historical dates:

156

Historical range:

`2022-01-03` through `2024-12-23`

Date interval validation:

- intervals: 155
- minimum interval: 7 days
- maximum interval: 7 days
- intervals different from 7 days: 0

Weekday validation:

- historical dates: 156
- Monday dates: 156
- non-Monday dates: 0

Conclusion:

Inventory History grain is validated as:

1 Product × 1 Site × 1 Week

with no missing weekly period in the validated source range.

Status:

PASS

---

# 12. Null / Blank Validation

Only one source field contains blank values:

`parts_master.shelf_life_days`

Observed:

- non-null: 26
- null / blank: 274

All other inspected source fields contain no blank values.

Interpretation:

`shelf_life_days` is an optional Product attribute.

Status:

PASS

---

# 13. Shelf-Life Behavior

Observed non-null Shelf Life range:

437 through 1,063 days

Non-null Shelf Life appears only in:

- Electrical
- Hydraulics

Electrical:

- total Products: 41
- with Shelf Life: 15
- without Shelf Life: 26

Hydraulics:

- total Products: 33
- with Shelf Life: 11
- without Shelf Life: 22

Repairability among Products with Shelf Life:

- No: 16
- Yes: 10

Conclusion:

No rule is supported requiring all Products of a specific family or repairability state to have Shelf Life.

Status:

PASS

---

# 14. Categorical Domain Validation

Validated current domains:

## Part Family

- Avionics
- Cabin
- Electrical
- Engine
- Fasteners
- Hydraulics
- LandingGear
- Structure

## Criticality

- A
- B
- C

## Supplier Risk

- High
- Low
- Medium

## Is Repairable

- Yes
- No

## Planned Maintenance

- True
- False

## Forecast Type

- Adjusted
- Baseline

## Defect Severity

- Critical
- Major
- Minor

## Defect Type

- Certification
- Dimensional
- Documentation
- Material
- Packaging
- Surface finish

Status:

PASS

---

# 15. Identifier Format Validation

Observed required formats:

- Product ID → `P#####`
- Supplier ID → `SUP###`
- Site ID → `SITE##`
- Purchase Order ID → `PO######`
- Quality Incident ID → `QI######`

Invalid values observed:

- Product IDs: 0
- Supplier IDs: 0
- Site IDs: 0
- Purchase Order IDs: 0
- Quality Incident IDs: 0

Status:

PASS

---

# 16. Product Numeric Validation

Observed ranges:

| Field | Minimum | Maximum |
|---|---:|---:|
| Unit Cost | 118.18 | 18,478 |
| Master Lead Time Days | 12 | 100 |
| Shelf Life Days, non-null | 437 | 1,063 |

Invalid conditions tested:

- negative Unit Cost: 0
- non-positive Master Lead Time: 0
- negative Shelf Life: 0

Status:

PASS

---

# 17. Inventory Numeric Validation

Observed ranges:

| Field | Minimum | Maximum |
|---|---:|---:|
| Consumption Qty | 0 | 73 |
| On-Hand Qty | 0 | 442 |
| Backorder Qty | 0 | 23 |
| Blocked Qty | 0 | 61 |
| Forecast Qty | 0 | 36 |

Invalid conditions tested:

- negative Consumption: 0
- negative On-Hand: 0
- negative Backorder: 0
- negative Blocked Qty: 0
- Blocked Qty > On-Hand Qty: 0
- fractional Forecast Qty rows: 0

Status:

PASS

---

# 18. Forecast Semantics Validation

Observed:

Baseline:

- rows: 227,059
- uplift range: 0.00 through 0.00

Adjusted:

- rows: 53,741
- uplift range: -0.30 through 0.42

Cross-field validation:

- Baseline with non-zero uplift: 0
- Adjusted with zero uplift: 0

Current validated source rule:

- Baseline → uplift = 0
- Adjusted → uplift <> 0

Status:

PASS

---

# 19. Purchase Order Quantity Validation

Observed:

| Metric | Value |
|---|---:|
| Ordered Qty minimum | 1 |
| Ordered Qty maximum | 263 |
| Received Qty minimum | 1 |
| Received Qty maximum | 263 |
| Full receipts | 26,311 |
| Partial receipts | 3,355 |
| Zero receipts | 0 |

Invalid conditions tested:

- non-positive Ordered Qty: 0
- negative Received Qty: 0
- Received Qty > Ordered Qty: 0

Partial receipts are valid business behavior and are not classified as failures.

Status:

PASS

---

# 20. Purchase Order Chronological Validation

Observed date ranges:

Order Date:

`2022-01-03` through `2024-12-23`

Promised Date:

`2022-01-15` through `2025-04-05`

Receipt Date:

`2022-01-12` through `2025-04-14`

Tests:

| Test | Observed | Status |
|---|---:|---|
| Promised Date before Order Date | 0 | PASS |
| Receipt Date before Order Date | 0 | PASS |
| Receipt Date after Promised Date | 16,568 | WARNING |

`ReceiptDate > PromisedDate` represents a late receipt rather than invalid source data.

It is valid supplier-performance information.

---

# 21. Purchase Order Lead-Time Validation

Promised Lead Time:

- minimum: 10 days
- maximum: 123 days

Actual Lead Time:

- minimum: 8 days
- maximum: 131 days
- source-wide average: 42.97 days

No negative Actual Lead Time exists.

Product-master Lead Time and Purchase Order Actual Lead Time are explicitly treated as different concepts.

Status:

PASS

---

# 22. Quality Incident Validation

Observed:

- Quality Incident rows: 368
- Scrap Qty minimum: 1
- Scrap Qty maximum: 13
- zero Scrap Qty incidents: 0
- negative Scrap Qty: 0

Status:

PASS

---

# 23. Date-Domain Validation

Observed minimum required date:

`2022-01-03`

Observed maximum required date:

`2025-04-14`

Proposed continuous DimDate range:

`2022-01-03` through `2025-04-14`

Inclusive calendar-day count:

1,198

Inventory historical maximum:

`2024-12-23`

The complete Date dimension may extend beyond the valid Inventory Reporting Date because procurement and quality facts contain later dates.

Status:

PASS

---

# 24. Logical Data-Type Validation

Phase 1 inspected source fields for compatibility with the planned logical data types.

Validated logical categories include:

- Text identifiers;
- Date;
- Whole Number;
- Decimal;
- Logical / Boolean;
- Percentage-style decimal.

No known type-conversion anomaly remains unresolved from Phase 1 source inspection.

Status:

PASS

---

# 25. Phase 1 Data-Design Test Summary

Current result:

PASS

No critical source-data ambiguity is currently known.

Warnings / expected valid behavior:

1. `shelf_life_days` is blank for 274 Products and is intentionally nullable.
2. 16,568 Purchase Orders were received after Promised Date and represent legitimate late-delivery observations.
3. 3,355 Purchase Orders are partial receipts and represent legitimate procurement behavior.

These observations must not be silently cleaned away.

---

# 26. Tests Not Yet Run

The following are NOT RUN because implementation has not reached their corresponding phases:

- configuration validation;
- formula validation;
- replenishment calculation validation;
- supplier-performance formula validation;
- Quality Control worksheet behavior;
- PivotTable reconciliation;
- PivotTable refresh;
- VBA automation;
- VBA failure handling;
- workbook protection;
- workbook performance;
- navigation;
- end-to-end testing;
- User Acceptance Testing.

These tests must not be marked PASS until their actual implementation exists.

---

# 27. Mandatory User Acceptance Scenarios

## Replenishment

Status:

NOT RUN

An Inventory Analyst must eventually be able to identify which Products require action at a selected Site and understand the recommended quantity.

## Supplier Performance

Status:

NOT RUN

A Procurement Manager must eventually be able to identify Suppliers with delivery, Lead-Time or quality issues.

## Refresh

Status:

NOT RUN

A valid source update must eventually be processed through the documented refresh workflow without manual copy/paste.

## Data Quality Failure

Status:

NOT RUN

A critical data-quality problem must eventually be visible before results are treated as reliable.

---

# 28. Phase 1 Evidence Boundary

Phase 1 testing establishes confidence in:

- source structure;
- source semantics;
- grains;
- keys;
- relationships;
- nullability;
- domains;
- ranges;
- logical data types;
- source-supported integrity rules.

It does not prove implementation of any Excel, Power Query or VBA component.

---

# Phase 2 — Workbook Foundation Test Evidence

## Phase 2 Test Environment

Target application:

Microsoft Excel 365 Desktop for Windows.

Workbook:

`workbook/ProcureFlow.xlsm`

Phase branch:

`phase/02-workbook-foundation`

Validation method:

Manual workbook inspection and interaction in Microsoft Excel.

The tests below validate only Workbook Foundation functionality implemented during Phase 2.

They do not imply that Power Query, operational calculations, PivotTables, VBA, reporting logic or dashboard functionality has been implemented.

---

## Workbook Foundation Tests

| Test ID | Test | Expected Result | Observed Result | Status |
|---|---|---|---|---|
| P2-WB-001 | Workbook opens normally | Workbook opens without errors or repair prompts | Workbook opened normally | PASS |
| P2-WB-002 | Worksheet count | 17 worksheets | 17 worksheets | PASS |
| P2-WB-003 | Worksheet architecture and order | Approved Phase 2 worksheet structure and order exists | Structure and order validated | PASS |
| P2-WB-004 | HOME to Configuration navigation | Link opens `01_CONFIG` | Navigation successful | PASS |
| P2-WB-005 | HOME to Quality Control navigation | Link opens `02_CONTROL` | Navigation successful | PASS |
| P2-WB-006 | Configuration to HOME navigation | Link returns to `00_HOME` | Navigation successful | PASS |
| P2-WB-007 | Reporting Date exposure | `00_HOME` displays configured Reporting Date | `23-Dec-2024` displayed | PASS |
| P2-WB-008 | Demand History validation | Values below 1 are rejected | `0` rejected | PASS |
| P2-WB-009 | Service Level validation | Values above 100% are rejected | `105%` rejected | PASS |
| P2-WB-010 | Locked configuration labels | Locked cells cannot be edited | `01_CONFIG!A6` protected | PASS |
| P2-WB-011 | Editable configuration inputs | `B6:B12` remain editable | Inputs editable | PASS |
| P2-WB-012 | Configuration Defined Names | Seven approved `cfg_*` workbook names exist | Names validated | PASS |
| P2-WB-013 | Workbook visual conventions | Approved typography, palette and layer tab colors are applied | Visual foundation validated | PASS |
| P2-WB-014 | Technical-sheet implementation boundary | Future-phase sheets contain no premature business implementation | Only Phase 2 placeholders present | PASS |
| P2-WB-015 | Workbook-wide HOME navigation | Every non-HOME worksheet provides a return link to `00_HOME` | Return navigation validated across all 16 worksheets | PASS |
| P2-WB-016 | Incremental protection boundary | Only worksheets with defined editable areas are protected | `01_CONFIG` protected; unfinished layers remain unprotected | PASS |
| P2-WB-017 | Physical Date worksheet boundary | `16_DATA_Date` exists structurally without premature Date-dimension implementation | Placeholder exists; no `tblDate` or Power Query implementation present | PASS |
| P2-WB-018 | Workbook default typography | Workbook Normal style uses Aptos 11 as the default font | Normal style configured as Aptos 11 | PASS |
---

## Phase 2 Workbook Foundation Evidence

Implemented and manually validated so far:

- first physical `ProcureFlow.xlsm` workbook;
- 17-sheet workbook architecture;
- base HOME navigation;
- central configuration worksheet;
- seven controlled business parameters;
- Data Validation for configuration inputs;
- seven workbook-scoped `cfg_*` Defined Names;
- Reporting Date exposure through `cfg_ReportingDate`;
- initial Quality Control worksheet structure;
- technical worksheet placeholders;
- approved workbook visual design system;
- layer-based worksheet tab colors;
- editable-input visual convention;
- initial worksheet protection for `01_CONFIG`.
- workbook-wide return navigation to `00_HOME`;
- confirmed permanent `16_DATA_Date` physical worksheet;
- incremental workbook-protection strategy.
- Aptos 11 configured through the workbook Normal cell style.

Not implemented in this evidence:

- Power Query ingestion;
- final Excel data tables;
- physical data model outputs;
- operational calculation engine;
- Quality Control logic;
- PivotTables;
- PivotCharts;
- VBA;
- refresh automation;
- replenishment reporting logic;
- management dashboard logic.

These remain assigned to their respective roadmap phases.

---

# Phase 3 — Power Query Pipeline Test Evidence

## Phase 3 Test Environment

Target application:

Microsoft Excel 365 Desktop for Windows.

Workbook:

`workbook/ProcureFlow.xlsm`

Phase branch:

`phase/03-power-query-pipeline`

Validation methods:

- Power Query Refresh All;
- loaded Excel Table inspection;
- workbook Quality Control formulas;
- row-count reconciliation against the validated Phase 1 source baseline;
- key-count reconciliation;
- Date-dimension range and continuity checks;
- derived Purchase Order flag reconciliation;
- manual Power Query load-behavior inspection.

Raw source files remained unchanged during testing.

---

## Power Query Pipeline Validation

| Test ID | Test | Expected | Observed | Status |
|---|---|---:|---:|---|
| P3-PQ-001 | Products row count | 300 | 300 | PASS |
| P3-PQ-002 | Sites row count | 6 | 6 | PASS |
| P3-PQ-003 | Suppliers row count | 40 | 40 | PASS |
| P3-PQ-004 | Inventory History row count | 280,800 | 280,800 | PASS |
| P3-PQ-005 | Purchase Orders row count | 29,666 | 29,666 | PASS |
| P3-PQ-006 | Quality Incidents row count | 368 | 368 | PASS |
| P3-PQ-007 | Date Dimension row count | 1,198 | 1,198 | PASS |
| P3-PQ-008 | Distinct Product IDs | 300 | 300 | PASS |
| P3-PQ-009 | Distinct Supplier IDs | 40 | 40 | PASS |
| P3-PQ-010 | Distinct Purchase Order IDs | 29,666 | 29,666 | PASS |
| P3-PQ-011 | Distinct Quality Incident IDs | 368 | 368 | PASS |
| P3-PQ-012 | Date Dimension minimum | 2022-01-03 | 2022-01-03 | PASS |
| P3-PQ-013 | Date Dimension maximum | 2025-04-14 | 2025-04-14 | PASS |
| P3-PQ-014 | Date Dimension continuous daily coverage | TRUE | TRUE | PASS |
| P3-PQ-015 | Late receipt count | 16,568 | 16,568 | PASS |
| P3-PQ-016 | Partial receipt count | 3,355 | 3,355 | PASS |
| P3-PQ-017 | Blank Shelf Life count | 274 | 274 | PASS |

All workbook Quality Control exceptions for `PQ-001` through `PQ-017` were zero.

---

## Refresh Validation

A full Excel `Refresh All` was executed after the final Phase 3 Power Query tables and workbook controls were implemented.

Observed result:

- source queries refreshed without reported error;
- staging queries refreshed without reported error;
- final dimension and fact queries refreshed without reported error;
- loaded Excel Tables remained available after Refresh;
- Quality Control results remained PASS after Refresh.

Status:

PASS

---

## Query Load Behavior

Intermediate queries use Connection Only behavior:

- `src_PartsMaster`
- `src_SupplyChainHistory`
- `src_PurchaseOrders`
- `src_QualityIncidents`
- `stg_Products`
- `stg_InventoryHistory`
- `stg_PurchaseOrders`
- `stg_QualityIncidents`

Final outputs are loaded to structured Excel Tables:

| Power Query Output | Excel Table |
|---|---|
| `dim_Product` | `tblProducts` |
| `dim_Site` | `tblSites` |
| `dim_Supplier` | `tblSuppliers` |
| `fact_InventoryWeekly` | `tblInventoryHistory` |
| `fact_PurchaseOrders` | `tblPurchaseOrders` |
| `fact_QualityIncidents` | `tblQualityIncidents` |
| `dim_Date` | `tblDate` |

The Power Query outputs are not loaded to the Excel Data Model.

Status:

PASS

---

## Phase 3 Transformation Reconciliation

The Phase 3 pipeline preserves validated source behavior:

- 274 Product Shelf Life values remain blank and intentionally nullable;
- 16,568 Purchase Orders are identified as late receipts;
- 3,355 Purchase Orders are identified as partial receipts;
- Inventory History remains at 280,800 rows;
- Date Dimension contains continuous daily coverage from `2022-01-03` through `2025-04-14`;
- raw CSV files remain unchanged.

Status:

PASS

---

## Phase 3 Testing Result

Current result:

PASS

No critical Power Query pipeline defect is currently known.

This evidence validates Phase 3 ingestion, staging, dimension/fact preparation, final table loading, Refresh behavior and baseline reconciliation.

It does not validate later-phase operational formulas, replenishment calculations, supplier-performance calculations, PivotTables, VBA, reporting or dashboard functionality.
---

## Phase 4 — Replenishment Foundation Validation

### Status

PASS

### Scope

Validation of the initial `tblReplenishment` structural and source-input implementation in `20_CALC_Replenishment`.

No Phase 5 replenishment business formulas are included in this validation.

### Operational Grain

| Test | Expected | Result |
|---|---:|---:|
| Product population | 300 | 300 |
| Site population | 6 | 6 |
| Product × Site population | 1,800 | 1,800 |
| `tblReplenishment` rows | 1,800 | 1,800 |
| Unique `ProductSiteKey` values | 1,800 | 1,800 |

Result:

PASS

### Master Attribute Validation

| Test | Expected | Result |
|---|---:|---:|
| Master-data formula errors | 0 | 0 |
| Distinct Primary Suppliers | 40 | 40 |
| Rows for Product `P00001` | 6 | 6 |

Result:

PASS

### Reporting Date and Inventory Snapshot Validation

Current restored Reporting Date:

`2024-12-23`

Current Inventory Snapshot Date:

`2024-12-23`

DEC-051 non-Monday test:

- temporary Reporting Date: `2024-12-18`
- expected Inventory Snapshot Date: `2024-12-16`
- actual Inventory Snapshot Date: `2024-12-16`
- result: PASS

The Reporting Date was restored to `2024-12-23` after the test.

### Inventory Snapshot Reconciliation

| Test | Expected | Result |
|---|---:|---:|
| Inventory source rows at snapshot | 1,800 | 1,800 |
| Product × Site source matches | 1,800 | 1,800 |
| On-Hand reconciliation | TRUE | TRUE |
| Blocked reconciliation | TRUE | TRUE |
| Backorder reconciliation | TRUE | TRUE |
| Blocked > On-Hand exceptions | 0 | 0 |

Result:

PASS

### Phase Boundary

The following remain outside this validation and remain assigned to Phase 5:

- Average Weekly Demand
- Demand variability
- Actual Lead-Time metrics
- Effective Lead Time
- Open PO Quantity
- Available Stock
- Service Level
- Safety Stock
- Reorder Point
- Inventory Position
- Target Stock
- Recommended Order Quantity
- Inventory Status
- No Recent Demand
---

## Phase 4 — Supplier Performance Foundation Validation

### Status

PASS

### Scope

Validation of the structural `tblSupplierPerformance` model in `21_CALC_SupplierPerformance`.

No Phase 5 supplier-performance metrics are included.

### Supplier Grain

| Test | Expected | Result |
|---|---:|---:|
| Supplier rows | 40 | 40 |
| Unique Supplier IDs | 40 | 40 |
| Population reconciliation | TRUE | TRUE |
| Master-data errors | 0 | 0 |
| Distinct Reporting Dates | 1 | 1 |

Result:

PASS

### Implemented Fields

- `SupplierID`
- `SupplierRiskClass`
- `ReportingDate`

### Phase 5 Boundary

The following remain unimplemented:

- Received PO Count
- On-Time PO Count
- On-Time Delivery Rate
- Late PO Count
- Late Delivery Rate
- Partial PO Count
- Partial Receipt Rate
- Average Actual Lead Time
- Lead-Time variability
- Quality Incident Count

No aggregate Supplier Score has been implemented.
---

## Phase 4 — Historical Demand Context Validation

### Status

PASS

### Scope

Validation of the completed-week historical-demand temporal context prepared in `tblReplenishment`.

No Phase 5 demand aggregation or statistical formulas are included.

### Results

| Test | Expected | Result |
|---|---:|---:|
| Distinct Demand History Week configurations | 1 | 1 |
| Rows using 26-week configuration | 1,800 | 1,800 |
| Distinct Demand History Start Dates | 1 | 1 |
| Distinct Demand History End Dates | 1 | 1 |
| Window length equals configured weeks | TRUE | TRUE |
| Window ends one week before Inventory Snapshot | TRUE | TRUE |
| Distinct source weeks in historical window | 26 | 26 |
| Inventory History source rows in window | 46,800 | 46,800 |

Result:

PASS

### Current Validated Window

- Inventory Snapshot Date: `2024-12-23`
- Demand History Weeks: 26
- Demand History Start Date: `2024-06-24`
- Demand History End Date: `2024-12-16`

The validated interval contains exactly:

26 weeks × 1,800 Product × Site combinations = 46,800 source observations.

### Phase Boundary

The following remain assigned to Phase 5:

- historical Consumption aggregation
- Average Weekly Demand
- demand variability
- `NO_RECENT_DEMAND`
- replenishment business calculations
---

## Phase 4 — Final Refresh and Exit Validation

### Status

PASS

### Refresh Validation

A full Excel `Refresh All` was executed after the Phase 4 operational-model implementation.

Observed result:

PASS

No critical Power Query refresh error was observed.

### Post-Refresh Operational Validation

| Test | Expected | Result |
|---|---:|---:|
| `tblReplenishment` rows | 1,800 | 1,800 |
| Unique `ProductSiteKey` values | 1,800 | 1,800 |
| Replenishment master/input formula errors | 0 | 0 |
| Reporting Date relationship | TRUE | TRUE |
| Historical-demand source rows | 46,800 | 46,800 |
| `tblSupplierPerformance` rows | 40 | 40 |
| Unique Supplier IDs | 40 | 40 |
| Supplier structural formula errors | 0 | 0 |

Result:

PASS

### Phase 4 Exit Criteria

#### Operational grains are correct

PASS

Validated grains:

- `tblReplenishment`: 1 Product × 1 Site
- `tblSupplierPerformance`: 1 Supplier

#### Required inputs are available

PASS

Prepared areas include:

- current inventory;
- blocked stock;
- backorders;
- Reporting Date;
- Inventory Snapshot Date;
- completed historical-demand window;
- Product and Supplier attributes;
- master Lead Time;
- Purchase Order transactional inputs;
- historical Actual Lead Time input.

#### No unresolved critical structural issue remains

PASS

No unresolved critical Phase 4 structural defect is currently known.

#### Model ready for business-rule formulas

PASS

The operational model is ready for Phase 5 — Business Logic & Advanced Formulas.

### Phase Boundary

Phase 5 calculations remain unimplemented.

## Phase 5 — Business Logic Validation

Status: PASS

Phase 5 business formulas were validated against the complete operational model.

### Replenishment Validation

All planned reconciliation tests returned their expected result:

- Historical Demand global reconciliation: PASS
- NoRecentDemand reconciliation: PASS
- Available Stock reconciliation: PASS
- Open PO Quantity global reconciliation: PASS
- Inventory Position reconciliation: PASS
- Reorder Point reconciliation: PASS
- Target Stock reconciliation: PASS
- negative Recommended Order Quantity count: 0
- non-integer Recommended Order Quantity count: 0
- NoRecentDemand with no backorder generating a purchase recommendation: 0
- invalid Inventory Status count: 0
- Phase 5 formula error count: 0

### Inventory Status Distribution

- STOCKOUT: 2
- CRITICAL: 86
- REORDER: 310
- ATTENTION: 103
- EXCESS: 967
- HEALTHY: 332
- Total: 1,800

The distribution reconciles exactly to the Product × Site operational grain.

The comparatively high EXCESS population is retained as an observed business-model result and is not treated as a formula defect without further business evidence.

### Additional Replenishment Evidence

- NoRecentDemand rows: 0
- rows using Master Lead Time fallback: 0
- rows with positive Recommended Order Quantity: 392
- total Recommended Order Quantity: 2,109 units
- Open PO Quantity at Reporting Date: 20,146 units
- Phase 5 formula errors: 0

### Supplier Performance Validation

All planned supplier-performance reconciliations returned their expected result:

- Received PO reconciliation: PASS
- On-Time PO + Late PO = Received PO: PASS
- On-Time Delivery Rate + Late Delivery Rate = 100% where received POs exist: PASS
- Partial PO reconciliation: PASS
- Late PO reconciliation: PASS
- Quality Incident reconciliation: PASS

### Refresh and Calculation Performance

Refresh All completed successfully with the full dataset.

Observed end-to-end Refresh All duration was approximately 10 minutes.

A separate full Excel formula recalculation was measured using `Application.CalculateFull()`.

Measured full recalculation duration:

- 51.4639623 seconds

Interpretation:

- Phase 5 formula recalculation completes in under one minute on the development environment.
- formula performance is accepted for the Phase 5 exit criterion with a documented performance observation;
- the approximately 10-minute end-to-end Refresh All duration is not attributable solely to Phase 5 formula recalculation and is retained as a future performance observation rather than treated as a Phase 5 formula defect.

No business rules were changed solely to improve performance.

---

# Phase 6 — Quality Control System Test Evidence

## Status

PASS — TECHNICAL IMPLEMENTATION AND CONTROLLED VALIDATION COMPLETE

GitHub publication gate remains pending.

## Environment

Application:

Microsoft Excel 365 Desktop for Windows

Workbook:

`workbook/ProcureFlow.xlsm`

Phase branch:

`phase/06-quality-control-system`

## Operational Quality Control Baseline

Implemented table:

`tblQualityControl`

Defined controls:

35

Final valid-baseline results:

| Metric | Observed |
|---|---:|
| Defined controls | 35 |
| Applicable controls | 34 |
| PASS | 34 |
| WARNING | 0 |
| FAIL | 0 |
| N/A | 1 |
| Total exceptions | 0 |
| Overall Quality Status | PASS |

`QC-032 — PivotTable refresh status` is intentionally N/A because PivotTables are assigned to Phase 7.

## Power Query Technical Health Feed

Implemented query:

`qc_PipelineHealth`

Loaded table:

`tblQCPipelineHealth`

Validated technical conditions:

- all 4 official source files available;
- all required source columns present;
- no staging / type-error rows reported;
- Products source and loaded population reconcile at 300 rows;
- Inventory History source and loaded population reconcile at 280,800 rows;
- Purchase Orders source and loaded population reconcile at 29,666 rows;
- Quality Incidents source and loaded population reconcile at 368 rows;
- current technical evaluation is valid;
- no Power Query execution errors reported.

Result:

PASS

## Integrity and Business-Rule Validation

Validated controls include:

- duplicate Product IDs;
- duplicate Purchase Order IDs;
- duplicate Quality Incident IDs;
- invalid Product references;
- Supplier consistency;
- Site consistency;
- negative inventory checks;
- Purchase Order quantity checks;
- Blocked Quantity consistency;
- Purchase Order chronological checks;
- Reporting Date range;
- Safety Stock;
- Reorder Point;
- Recommended Order Quantity;
- Inventory Status domain;
- replenishment-rule reconciliation;
- configuration validity;
- duplicate Inventory composite keys;
- Date dimension continuity;
- Excel calculation formula errors.

Final valid-baseline exception count:

0

## Critical FAIL Path Test

Controlled test:

- `QC-005` ExceptionCount temporarily set to `1`;
- Severity remained `Critical`.

Observed:

- `QC-005` Status → FAIL;
- Overall Quality Status → FAIL;
- FAIL control count → 1;
- Total Exceptions → 1.

The original formula was restored after the test.

Result:

PASS

## WARNING Path Test

Controlled test:

- `QC-005` Severity temporarily changed to `Warning`;
- ExceptionCount temporarily set to `1`.

Observed:

- `QC-005` Status → WARNING;
- Overall Quality Status → WARNING;
- WARNING control count → 1;
- FAIL control count → 0.

The original Critical severity and formula were restored after the test.

Result:

PASS

## Incomplete-Evaluation Safeguard Test

Controlled test:

- `QC-005` ExceptionCount temporarily cleared.

Observed:

- `QC-005` Status became blank;
- Overall Quality Status became WARNING rather than PASS.

The original formula was restored after the test.

Result:

PASS

## Restored Final State

After all controlled tests:

- Overall Quality Status: PASS;
- PASS controls: 34;
- WARNING controls: 0;
- FAIL controls: 0;
- Applicable controls: 34;
- N/A controls: 1;
- Total Exceptions: 0;
- `QC-032`: N/A.

Result:

PASS

## Refresh-State Limitation

Phase 6 validates the current Power Query technical evaluation timestamp through `QC-030`.

The phase does not persist the timestamp of a previous successful refresh after a later failed refresh attempt.

Persistent successful-refresh state remains assigned to later workflow automation.

This limitation is documented and is not treated as implemented functionality.

## Phase 6 Exit-Criteria Evidence

Mandatory controls implemented:

PASS

Critical FAIL behavior:

PASS

Warning behavior:

PASS

Incomplete-evaluation safeguard:

PASS

Reconciliations on the valid source baseline:

PASS

Critical data-quality failure visibility:

PASS

No critical Quality Control issue is currently known.
