# Phase 10 — Reporting & Dashboard Design

## Status

[PROPUESTO]

Phase:

`Phase 10 — Reporting & Dashboard`

Target Version:

`v0.10.0`

Phase Branch:

`phase/10-reporting-dashboard`

Baseline:

`v0.9.2`

## Purpose

Define the Phase 10 reporting and dashboard design before physical workbook implementation.

This document is a phase-level design artifact.

It does not replace the canonical project documents:

- `PROJECT_SPEC.md`
- `ROADMAP.md`
- `ARCHITECTURE.md`
- `DECISIONS.md`
- `CURRENT_STATE.md`

Nothing in this document is considered implemented until supported by workbook and validation evidence.

---

# 1. Reporting Principles

[PROPUESTO]

Phase 10 will preserve the established ProcureFlow architecture:

Power Query
→ structured Excel Tables
→ operational calculations
→ Quality Control
→ analytical aggregation
→ VBA orchestration
→ operational reporting and management dashboard

Reporting and dashboard worksheets will present validated upstream information.

They will not become a new hidden business-calculation engine.

Metrics described as being valid at the Reporting Date must not use future information.

Important meaning must not depend on color alone.

---

# 2. Operational Replenishment Report

Worksheet:

`40_RPT_Replenishment`

Primary user:

Inventory Analyst / Buyer

Primary business question:

Which Product × Site combinations require operational attention, why do they require attention, and what quantity is currently recommended for replenishment?

Primary source:

`tblReplenishment`

## 2.1 Proposed Fields

The operational report should expose the business context required to understand and act on replenishment recommendations.

Candidate fields:

- ProductID
- SiteID
- PartFamily
- CriticalityClass
- PrimarySupplierID
- SupplierRiskClass
- InventoryStatus
- AvailableStock
- BackorderQty
- OpenPOQty
- InventoryPosition
- SafetyStock
- ReorderPoint
- TargetStock
- RecommendedOrderQty
- NoRecentDemand

Additional fields may be included only when they materially improve operational decision-making.

## 2.2 Status Priority

The existing Inventory Status priority remains authoritative:

1. STOCKOUT
2. CRITICAL
3. REORDER
4. ATTENTION
5. EXCESS
6. HEALTHY

Phase 10 will not create a competing inventory-status classification.

## 2.3 Formula Strategy

Dynamic Array formulas are preferred where they provide a transparent and maintainable operational view.

Candidate functions include:

- FILTER
- SORTBY
- LET

The final formula design must be explained and validated before being considered implemented.

---

# 3. Management Dashboard

Worksheet:

`41_DASH_Management`

Primary user:

Procurement / Operations Manager

Primary business question:

Where are the most important current inventory, procurement and supplier risks requiring management attention?

The dashboard is a presentation layer.

Complex business rules must remain upstream.

## 3.1 Status and Context

Candidate management context:

- Reporting Date
- Inventory Snapshot Date
- Last Successful Refresh
- Overall Quality Status

## 3.2 Inventory and Replenishment

Candidate KPIs:

- Product-Site Count
- STOCKOUT count
- CRITICAL count
- REORDER count
- Backorder Qty
- Available Stock
- Open PO Qty
- Recommended Order Qty

Candidate visual analysis:

- Inventory Status distribution
- Inventory risk by Site
- replenishment requirement by Site or other justified operational dimension

Primary source:

`tblReplenishment`

## 3.3 Procurement and Supplier Performance

Candidate KPIs and measures:

- Received PO Count
- On-Time Delivery Rate
- Late Delivery Rate
- Partial Receipt Rate
- Avg Actual Lead Time Days
- Lead Time variability
- Supplier population by Supplier Risk Class

Primary Reporting-Date-safe source:

`tblSupplierPerformance`

Historical Purchase Order analysis may use `tblPurchaseOrders` only when its date context is explicitly appropriate and does not introduce future information into an as-of KPI.

No arbitrary aggregate Supplier Score will be introduced.

## 3.4 Supplier Quality

Candidate measures:

- Quality Incident Count
- incidents by Supplier or Supplier Risk Class
- incidents by Defect Severity or Defect Type where useful

Quality metrics must clearly identify their date context.

---

# 4. Interactive Filtering

[PROPUESTO]

ProcureFlow will not force a universal dashboard slicer across incompatible PivotCaches or source tables.

Interactive controls will be introduced only where their scope is clear.

Candidate filters for replenishment-oriented analysis:

- SiteID
- PartFamily
- CriticalityClass
- SupplierRiskClass

Candidate supplier-oriented filters:

- SupplierID
- SupplierRiskClass

Candidate historical analysis controls:

- appropriate Timeline fields where the underlying analysis uses a date-based PivotTable.

A filter must not visually imply that it controls metrics to which it is not actually connected.

---

# 5. Existing Phase 7 Analytical Assets

The following validated Phase 7 analytical worksheets already exist:

- `30_PVT_Inventory`
- `31_PVT_Procurement`
- `32_PVT_Suppliers`

Existing PivotTables, PivotCharts, Slicers and Timelines may be reused where their semantic meaning and filter context match Phase 10 requirements.

They must not be copied into the dashboard merely because they already exist.

In particular, full-source Purchase Order analysis must not automatically be presented as an as-of Reporting Date management KPI.

---

# 6. Automation Boundary

The existing Phase 9 production workflow remains authoritative.

Current production automation refreshes and validates the established analytical PivotTables.

If Phase 10 creates additional PivotTables or other refresh-dependent reporting artifacts, their integration into production automation must be explicitly designed, implemented and validated.

Phase 10 must not silently break the validated Phase 9 PivotTable-count or PivotCache contracts.

---

# 7. Accessibility and Visual Design

The existing ProcureFlow visual system remains authoritative.

Important requirements include:

- Aptos as the default workbook typeface;
- established layer and status palette;
- clear visual hierarchy;
- meaningful labels;
- no critical meaning conveyed solely through color;
- avoidance of unnecessary visual clutter;
- user-facing worksheets designed as finished business outputs rather than technical calculation sheets.

---

# 8. Reconciliation Requirement

Every management KPI must have an identifiable source and a reconciliation method.

Dashboard totals must reconcile to their approved upstream tables, formulas or validated analytical outputs.

A visually correct dashboard is not considered valid if its displayed metrics cannot be independently reconciled.

---

# 9. Phase 10 Design State

Current state:

[PROPUESTO]

No Phase 10 reporting or dashboard implementation has yet been validated by this document.

The design must be finalized before physical workbook implementation begins.