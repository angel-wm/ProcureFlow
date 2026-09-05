# ProcureFlow — Data Dictionary

## Document Status

Status: IN PROGRESS — PHASE 1 DATA DESIGN

Phase: Phase 1 — Data Design

Target Version: `v0.2.0`

This document defines the validated source-data specification and logical ProcureFlow data model.

The definitions in this document describe approved data design.

They must not be interpreted as physically implemented Excel Tables, Power Query queries or workbook structures until implementation evidence exists in later phases.

---

# 1. Official Dataset

Dataset:

Aerospace Supply Chain Performance & Forecasting

Official source files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

Raw source location:

`data/raw/`

Raw files remain immutable and are excluded from Git.

---

# 2. Validated Source Summary

| Source File | Rows | Fields | Validated Grain |
|---|---:|---:|---|
| `parts_master.csv` | 300 | 9 | 1 row = 1 Product |
| `supply_chain_history.csv` | 280,800 | 11 | 1 Product × 1 Site × 1 Week |
| `purchase_orders.csv` | 29,666 | 9 | 1 row = 1 Purchase Order record for 1 Product |
| `quality_incidents.csv` | 368 | 8 | 1 row = 1 Quality Incident |

Validated entity populations:

- Products: 300
- Sites: 6
- Suppliers: 40
- Product × Site combinations: 1,800
- historical weeks: 156

---

# 3. Validated Keys

## 3.1 Product

Primary Key:

`part_id`

Validation:

- total rows: 300
- distinct `part_id`: 300
- blank `part_id`: 0

Observed format:

`P#####`

Example:

`P00001`

---

## 3.2 Inventory History

Composite Primary Key:

`date + site_id + part_id`

Validation:

- total rows: 280,800
- distinct composite keys: 280,800
- duplicate composite-key rows: 0

---

## 3.3 Purchase Orders

Primary Key:

`po_id`

Validation:

- total rows: 29,666
- distinct `po_id`: 29,666
- blank `po_id`: 0

Observed format:

`PO######`

Example:

`PO000001`

---

## 3.4 Quality Incidents

Primary Key:

`incident_id`

Validation:

- total rows: 368
- distinct `incident_id`: 368
- blank `incident_id`: 0

Observed format:

`QI######`

Example:

`QI000001`

---

# 4. Identifier Domains

## Product ID

Observed format:

`P#####`

Invalid observed values:

0

## Supplier ID

Observed format:

`SUP###`

Invalid observed values:

0

## Site ID

Observed format:

`SITE##`

Validated Sites:

- `SITE01`
- `SITE02`
- `SITE03`
- `SITE04`
- `SITE05`
- `SITE06`

Invalid observed values:

0

## Purchase Order ID

Observed format:

`PO######`

Invalid observed values:

0

## Quality Incident ID

Observed format:

`QI######`

Invalid observed values:

0

---

# 5. Source Referential Integrity

All validated Product references in:

- Inventory History;
- Purchase Orders;
- Quality Incidents

exist in `parts_master.csv`.

Invalid Product references:

0

All Site references in:

- Purchase Orders;
- Quality Incidents

exist in the validated Site population derived from Inventory History.

Invalid Site references:

0

The validated Supplier population contains 40 Suppliers.

All 40 Suppliers appear in:

- Product master;
- Purchase Orders;
- Quality Incidents.

Purchase Order Suppliers outside Product-master Supplier population:

0

Quality Suppliers outside Product-master Supplier population:

0

Blank Product, Site or Supplier identifiers in relational source fields:

0

---

# 6. Logical ProcureFlow Data Model

## Dimensions

- `DimProduct`
- `DimSite`
- `DimSupplier`
- `DimDate`

## Facts

- `FactInventoryWeekly`
- `FactPurchaseOrders`
- `FactQualityIncidents`

These names describe the approved logical data design.

Physical Power Query and Excel Table implementation occurs in later phases.

---

# 7. Dimension Specification

## 7.1 DimProduct

Grain:

1 row = 1 Product

Primary Key:

`ProductID`

Source:

`parts_master.csv`

Planned attributes:

- `ProductID`
- `PartFamily`
- `CriticalityClass`
- `UnitCost`
- `MasterLeadTimeDays`
- `PrimarySupplierID`
- `IsRepairable`
- `ShelfLifeDays`

`SupplierRiskClass` is physically present in `parts_master.csv` but logically belongs to `DimSupplier`.

---

## 7.2 DimSite

Grain:

1 row = 1 Site

Primary Key:

`SiteID`

Source:

unique `site_id` values derived from official sources.

Validated population:

6 Sites

No additional Site attributes will be invented because the official dataset does not provide them.

---

## 7.3 DimSupplier

Grain:

1 row = 1 Supplier

Primary Key:

`SupplierID`

Source:

Supplier identifiers and Supplier Risk information available through official source data.

Planned attributes:

- `SupplierID`
- `SupplierRiskClass`

Validated Supplier population:

40

Validated dependency:

1 Supplier → 1 Supplier Risk Class

All 40 Suppliers have exactly one observed Risk Class.

Supplier-to-Product cardinality:

1 Supplier → many Products

Observed Products per Supplier:

- minimum: 4
- maximum: 14

All 40 Suppliers supply multiple Products.

---

## 7.4 DimDate

Grain:

1 row = 1 calendar date

Primary Key:

`Date`

Required continuous source-supported range:

`2022-01-03` through `2025-04-14`

Calendar rows required for that inclusive range:

1,198

Planned attributes:

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

Weekly calendar convention is ISO 8601, with Monday as Week Start.

The Inventory History Reporting Date range is narrower than the complete Date dimension range.

Maximum Inventory History date:

`2024-12-23`

The later Purchase Order and Quality dates do not extend the inventory historical Reporting Date beyond available inventory evidence.

---

# 8. Fact Specification

## 8.1 FactInventoryWeekly

Source:

`supply_chain_history.csv`

Grain:

1 Product × 1 Site × 1 Week

Composite Key:

- `WeekStartDate`
- `SiteID`
- `ProductID`

Validated structure:

- 300 Products
- 6 Sites
- 1,800 Product × Site combinations
- 156 weeks per Product × Site
- 280,800 rows

Historical date range:

`2022-01-03` through `2024-12-23`

All 156 historical dates are Mondays.

All consecutive historical dates are exactly 7 days apart.

No Product × Site combination has a missing historical week within the validated source period.

---

## 8.2 FactPurchaseOrders

Source:

`purchase_orders.csv`

Grain:

1 row = 1 Purchase Order record for 1 Product

Primary Key:

`PurchaseOrderID`

The source does not support a traditional Purchase Order Header → Purchase Order Lines model.

ProcureFlow must not invent that structure.

Validated record count:

29,666

---

## 8.3 FactQualityIncidents

Source:

`quality_incidents.csv`

Grain:

1 row = 1 Quality Incident

Primary Key:

`QualityIncidentID`

Validated record count:

368

---

# 9. Field-Level Data Dictionary

## 9.1 parts_master.csv

### part_id

Logical target:

`DimProduct.ProductID`

Logical type:

Text

Required:

Yes

Role:

Primary Key

Observed format:

`P#####`

Unique:

Yes

Nulls:

0

---

### part_family

Logical target:

`DimProduct.PartFamily`

Logical type:

Text

Required:

Yes

Observed domain:

- Avionics
- Cabin
- Electrical
- Engine
- Fasteners
- Hydraulics
- LandingGear
- Structure

---

### criticality_class

Logical target:

`DimProduct.CriticalityClass`

Logical type:

Text

Required:

Yes

Observed domain:

- A
- B
- C

---

### unit_cost

Logical target:

`DimProduct.UnitCost`

Logical type:

Decimal

Required:

Yes

Observed range:

118.18 through 18,478

Validated condition:

`UnitCost >= 0`

Negative values observed:

0

---

### lead_time_days

Logical target:

`DimProduct.MasterLeadTimeDays`

Logical type:

Whole Number

Required:

Yes

Observed range:

12 through 100 days

Validated condition:

`MasterLeadTimeDays > 0`

This is a Product-master Lead Time attribute.

It must not be confused with Purchase Order Actual Lead Time derived from transaction dates.

---

### supplier_id_primary

Logical target:

`DimProduct.PrimarySupplierID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimSupplier.SupplierID`

Observed format:

`SUP###`

All Products have one Primary Supplier in the current dataset.

Purchase Order and Quality Incident Supplier values for each Product were validated against the Product Primary Supplier.

Mismatches observed:

0

---

### supplier_risk_class

Logical target:

`DimSupplier.SupplierRiskClass`

Logical type:

Text

Required:

Yes

Observed domain:

- High
- Low
- Medium

Although physically repeated in `parts_master.csv`, this attribute functionally depends on Supplier.

All 40 Suppliers have exactly one observed Supplier Risk Class.

---

### is_repairable

Logical target:

`DimProduct.IsRepairable`

Logical type:

Logical / Boolean

Required:

Yes

Observed source domain:

- Yes
- No

Power Query may standardize this field into a logical TRUE/FALSE representation if appropriate.

---

### shelf_life_days

Logical target:

`DimProduct.ShelfLifeDays`

Logical type:

Whole Number

Required:

No

Nullable:

Yes

Non-null rows:

26

Null rows:

274

Observed non-null range:

437 through 1,063 days

Non-null values occur only in the observed:

- Electrical
- Hydraulics

families.

However, Shelf Life remains selective even within those families:

Electrical:

- total Products: 41
- with Shelf Life: 15
- without Shelf Life: 26

Hydraulics:

- total Products: 33
- with Shelf Life: 11
- without Shelf Life: 22

No unsupported rule will be created requiring all Electrical or Hydraulics Products to have Shelf Life.

---

## 9.2 supply_chain_history.csv

### date

Logical target:

`FactInventoryWeekly.WeekStartDate`

Logical type:

Date

Required:

Yes

Role:

Composite Key component / Foreign Key to `DimDate`

Observed range:

`2022-01-03` through `2024-12-23`

Distinct dates:

156

Validated semantics:

weekly period start date

All observed values are Mondays.

---

### site_id

Logical target:

`FactInventoryWeekly.SiteID`

Logical type:

Text

Required:

Yes

Role:

Composite Key component / Foreign Key to `DimSite`

Observed format:

`SITE##`

---

### part_id

Logical target:

`FactInventoryWeekly.ProductID`

Logical type:

Text

Required:

Yes

Role:

Composite Key component / Foreign Key to `DimProduct`

Observed format:

`P#####`

---

### planned_maintenance

Logical target:

`FactInventoryWeekly.PlannedMaintenance`

Logical type:

Logical / Boolean

Required:

Yes

Observed domain:

- True
- False

---

### consumption_qty

Logical target:

`FactInventoryWeekly.ConsumptionQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

0 through 73

Validated condition:

`ConsumptionQty >= 0`

---

### on_hand_qty

Logical target:

`FactInventoryWeekly.OnHandQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

0 through 442

Validated condition:

`OnHandQty >= 0`

---

### backorder_qty

Logical target:

`FactInventoryWeekly.BackorderQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

0 through 23

Validated condition:

`BackorderQty >= 0`

---

### blocked_qty

Logical target:

`FactInventoryWeekly.BlockedQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

0 through 61

Validated conditions:

- `BlockedQty >= 0`
- `BlockedQty <= OnHandQty`

Violations observed:

0

---

### forecast_qty

Logical target:

`FactInventoryWeekly.ForecastQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

0 through 36

Rows containing fractional values:

0

Forecast is retained for analysis but does not drive the initial core replenishment engine.

---

### forecast_type

Logical target:

`FactInventoryWeekly.ForecastType`

Logical type:

Text

Required:

Yes

Observed domain:

- Adjusted
- Baseline

Observed rows:

- Adjusted: 53,741
- Baseline: 227,059

---

### forecast_uplift_pct

Logical target:

`FactInventoryWeekly.ForecastUpliftPct`

Logical type:

Decimal / Percentage

Required:

Yes

Baseline observed range:

0.00 through 0.00

Adjusted observed range:

-0.30 through 0.42

Validated source relationship:

- `ForecastType = Baseline` → `ForecastUpliftPct = 0`
- `ForecastType = Adjusted` → `ForecastUpliftPct <> 0`

Violations observed:

0

---

## 9.3 purchase_orders.csv

### po_id

Logical target:

`FactPurchaseOrders.PurchaseOrderID`

Logical type:

Text

Required:

Yes

Role:

Primary Key

Observed format:

`PO######`

Unique:

Yes

Nulls:

0

---

### supplier_id

Logical target:

`FactPurchaseOrders.SupplierID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimSupplier`

Observed format:

`SUP###`

---

### site_id

Logical target:

`FactPurchaseOrders.SiteID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimSite`

Observed format:

`SITE##`

---

### part_id

Logical target:

`FactPurchaseOrders.ProductID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimProduct`

Observed format:

`P#####`

---

### order_date

Logical target:

`FactPurchaseOrders.OrderDate`

Logical type:

Date

Required:

Yes

Role:

Foreign Key → `DimDate`

Observed range:

`2022-01-03` through `2024-12-23`

---

### promised_date

Logical target:

`FactPurchaseOrders.PromisedDate`

Logical type:

Date

Required:

Yes

Role:

Foreign Key → `DimDate`

Observed range:

`2022-01-15` through `2025-04-05`

Validated condition:

`PromisedDate >= OrderDate`

Violations observed:

0

---

### receipt_date

Logical target:

`FactPurchaseOrders.ReceiptDate`

Logical type:

Date

Required:

Yes

Role:

Foreign Key → `DimDate`

Observed range:

`2022-01-12` through `2025-04-14`

Validated condition:

`ReceiptDate >= OrderDate`

Violations observed:

0

Receipt after Promised Date is not an invalid condition.

Late receipts observed:

16,568

---

### ordered_qty

Logical target:

`FactPurchaseOrders.OrderedQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

1 through 263

Validated condition:

`OrderedQty > 0`

---

### received_qty

Logical target:

`FactPurchaseOrders.ReceivedQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

1 through 263

Validated conditions:

- `ReceivedQty > 0`
- `ReceivedQty <= OrderedQty`

Full receipts:

26,311

Partial receipts:

3,355

Zero receipts:

0

Received greater than Ordered Quantity:

0

---

# 10. Purchase Order Derived Fields

The following fields are not raw source fields.

They are objective derivations supported by source semantics.

## PromisedLeadTimeDays

Definition:

`PromisedDate - OrderDate`

Observed range:

10 through 123 days

Candidate implementation layer:

Power Query

---

## ActualLeadTimeDays

Definition:

`ReceiptDate - OrderDate`

Observed range:

8 through 131 days

Observed source-wide average:

42.97 days

Candidate implementation layer:

Power Query

This is distinct from Product-master `MasterLeadTimeDays`.

---

## IsLateReceipt

Definition:

`ReceiptDate > PromisedDate`

Candidate implementation layer:

Power Query

Late receipts observed in source:

16,568

---

## IsPartialReceipt

Definition:

`ReceivedQty < OrderedQty`

Candidate implementation layer:

Power Query

Partial receipts observed:

3,355

---

## IsOpenPO

This field must not be stored as a permanent source-level state.

Its meaning depends on the configured Reporting Date.

Business rule:

`OrderDate <= ReportingDate < ReceiptDate`

Candidate implementation layer:

Excel operational model / Reporting-Date-dependent calculation.

---

# 11. quality_incidents.csv

### incident_id

Logical target:

`FactQualityIncidents.QualityIncidentID`

Logical type:

Text

Required:

Yes

Role:

Primary Key

Observed format:

`QI######`

Unique:

Yes

Nulls:

0

---

### incident_date

Logical target:

`FactQualityIncidents.IncidentDate`

Logical type:

Date

Required:

Yes

Role:

Foreign Key → `DimDate`

Observed range:

`2022-01-26` through `2025-02-22`

---

### part_id

Logical target:

`FactQualityIncidents.ProductID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimProduct`

---

### supplier_id

Logical target:

`FactQualityIncidents.SupplierID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimSupplier`

---

### site_id

Logical target:

`FactQualityIncidents.SiteID`

Logical type:

Text

Required:

Yes

Role:

Foreign Key → `DimSite`

---

### defect_severity

Logical target:

`FactQualityIncidents.DefectSeverity`

Logical type:

Text

Required:

Yes

Observed domain:

- Critical
- Major
- Minor

---

### defect_type

Logical target:

`FactQualityIncidents.DefectType`

Logical type:

Text

Required:

Yes

Observed domain:

- Certification
- Dimensional
- Documentation
- Material
- Packaging
- Surface finish

---

### scrap_qty

Logical target:

`FactQualityIncidents.ScrapQty`

Logical type:

Whole Number

Required:

Yes

Observed range:

1 through 13

Validated condition:

`ScrapQty >= 0`

Negative values observed:

0

Zero-scrap incidents observed:

0

---

# 12. Nullability Summary

Only one source field currently contains null / blank values:

`parts_master.shelf_life_days`

Blank rows:

274

All other validated source fields contain no blank values.

`shelf_life_days` is therefore explicitly classified as optional.

No missing-value imputation will be invented for Shelf Life.

---

# 13. Validated Categorical Domains

## Product Family

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
- Medium
- Low

## Repairable

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

Observed domains describe the current official dataset.

They must not automatically be treated as permanently exhaustive if future valid source versions introduce documented new values.

---

# 14. Validated Relationships

Logical relationships:

`DimProduct.ProductID`
→ `FactInventoryWeekly.ProductID`

`DimProduct.ProductID`
→ `FactPurchaseOrders.ProductID`

`DimProduct.ProductID`
→ `FactQualityIncidents.ProductID`

`DimSite.SiteID`
→ `FactInventoryWeekly.SiteID`

`DimSite.SiteID`
→ `FactPurchaseOrders.SiteID`

`DimSite.SiteID`
→ `FactQualityIncidents.SiteID`

`DimSupplier.SupplierID`
→ `DimProduct.PrimarySupplierID`

`DimSupplier.SupplierID`
→ `FactPurchaseOrders.SupplierID`

`DimSupplier.SupplierID`
→ `FactQualityIncidents.SupplierID`

`DimDate.Date`
→ Inventory, Purchase Order and Quality date fields as applicable.

Validated source behavior:

1 Supplier → many Products

1 Product → many historical inventory rows

1 Site → many historical inventory rows

1 Product → many Purchase Orders

1 Supplier → many Purchase Orders

1 Site → many Purchase Orders

1 Product → many Quality Incidents where incidents exist

1 Supplier → many Quality Incidents where incidents exist

1 Site → many Quality Incidents where incidents exist

---

# 15. Core Data-Quality Rules Derived from Phase 1

The following rules are supported by validated source semantics.

## Structural

- all four official source files must exist;
- required source columns must exist;
- required fields must be non-null;
- source fields must be convertible to their approved logical data types.

## Keys

- ProductID must be unique;
- PurchaseOrderID must be unique;
- QualityIncidentID must be unique;
- WeekStartDate + SiteID + ProductID must be unique.

## Referential Integrity

- all fact Product IDs must exist in DimProduct;
- all fact Site IDs must exist in DimSite;
- all Supplier IDs must exist in DimSupplier.

## Inventory

- ConsumptionQty >= 0;
- OnHandQty >= 0;
- BackorderQty >= 0;
- BlockedQty >= 0;
- BlockedQty <= OnHandQty;
- WeekStartDate must represent the expected weekly structure.

## Product

- UnitCost >= 0;
- MasterLeadTimeDays > 0;
- ShelfLifeDays may be null;
- non-null ShelfLifeDays must not be negative.

## Purchase Orders

- OrderedQty > 0;
- ReceivedQty >= 0;
- ReceivedQty <= OrderedQty;
- PromisedDate >= OrderDate;
- ReceiptDate >= OrderDate.

ReceiptDate > PromisedDate represents late delivery and is valid business information rather than a source error.

## Forecast

- ForecastQty >= 0;
- Baseline records require ForecastUpliftPct = 0;
- Adjusted records require ForecastUpliftPct <> 0 under the current validated source semantics.

## Quality

- ScrapQty >= 0;
- required Quality identifiers must be valid;
- IncidentDate must be a valid date.

---

# 16. Observed Ranges vs Validation Rules

Observed minimum and maximum values document the current official dataset.

They must not automatically become hard validation limits.

Examples:

- current UnitCost maximum = 18,478;
- current Actual Lead Time maximum = 131;
- current Consumption maximum = 73.

A future valid source value above one of these observed maxima must not automatically be rejected solely because it exceeds the current dataset's historical range.

Hard validation should be based on business or semantic impossibility, not merely current observed extrema.

---

# 17. Source-to-Target Summary

## Product Source

`parts_master.csv`

produces information for:

- `DimProduct`
- `DimSupplier`

## Inventory Source

`supply_chain_history.csv`

produces:

- `FactInventoryWeekly`
- `DimSite` identifiers
- Date-domain evidence

## Purchase Order Source

`purchase_orders.csv`

produces:

- `FactPurchaseOrders`
- Supplier/Site relationship validation
- procurement Lead-Time derivations
- Date-domain evidence

## Quality Source

`quality_incidents.csv`

produces:

- `FactQualityIncidents`
- Supplier/Site relationship validation
- Date-domain evidence

## Date Dimension

`DimDate` is derived rather than copied from one source file.

It must provide a continuous calendar spanning all required source-supported dates.

---

# 18. Phase 1 Data-Design Resolution

The material logical data-design questions identified during source validation are resolved.

Confirmed:

- Supplier Risk belongs logically to `DimSupplier`;
- Inventory History `date` represents Monday-based `WeekStartDate`;
- `DimDate` has daily continuous grain;
- the current Date range is `2022-01-03` through `2025-04-14`;
- Inventory Reporting Date remains bounded by Inventory History, currently through `2024-12-23`;
- weekly calendar attributes use ISO 8601 semantics;
- objective Purchase Order derivations are specified;
- logical fact/dimension relationships are documented in `ARCHITECTURE.md`.

Exact implementation details such as Power Query query names or final Excel physical column placement may be refined during their implementation phases without changing this logical data specification.

No unresolved critical data-design ambiguity is currently known.

---
# 19. Implementation Status

[CONFIRMADO]

- source schemas inspected;
- grains validated;
- keys validated;
- referential integrity validated;
- logical dimensions and facts specified;
- field-level logical types specified;
- null behavior documented;
- categorical domains documented;
- source-supported quality rules identified.

[PENDIENTE]

- Power Query implementation;
- Excel Tables;
- workbook data sheets;
- physical Date dimension;
- Quality Control implementation;
- formulas;
- PivotTables;
- VBA.

Those items belong to later roadmap phases.


