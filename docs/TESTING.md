# ProcureFlow — Testing

## Document Status

Status: COMPLETED — PHASE 1 TEST EVIDENCE

Current Phase:

Phase 1 — Data Design

Target Version:

`v0.2.0`

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

- Power Query source ingestion;
- Power Query refresh;
- loaded row-count reconciliation;
- workbook Excel Table validation;
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
