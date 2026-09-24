# Phase 10 — Reporting & Dashboard Design

## Status

[CONFIRMADO] [IMPLEMENTADO]

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

[CONFIRMADO]

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

[CONFIRMADO]

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

[IMPLEMENTADO]

Phase 10 reporting and dashboard implementation has now been validated; physical implementation evidence is recorded in Section 19 and the canonical project documentation.

The design was finalized before physical workbook implementation began.

---

# 10. Physical Implementation Specification

Status:

[IMPLEMENTADO]

This section defines the physical Phase 10 implementation baseline before workbook construction.

The specifications in this section are now supported by workbook and validation evidence.

---

## 10.1 Operational Replenishment Report Layout

Worksheet:

`40_RPT_Replenishment`

Primary source:

`tblReplenishment`

Primary user:

Inventory Analyst / Buyer

### User-facing layout

Rows 1–2:

- report title;
- short report purpose / context.

Rows 4–6:

- Reporting Date;
- Inventory Snapshot Date;
- Last Successful Refresh;
- Overall Quality Status;
- operational filter controls.

Row 8:

report headers.

Row 9 downward:

dynamic report output.

The primary report area will use columns:

`A:P`

### Report columns

The operational output will expose the following columns in this order:

1. ProductID
2. SiteID
3. PartFamily
4. CriticalityClass
5. PrimarySupplierID
6. SupplierRiskClass
7. InventoryStatus
8. AvailableStock
9. BackorderQty
10. OpenPOQty
11. InventoryPosition
12. SafetyStock
13. ReorderPoint
14. TargetStock
15. RecommendedOrderQty
16. NoRecentDemand

These values must originate from the validated `tblReplenishment` operational model.

The reporting sheet must not recalculate the underlying replenishment business rules.

### Operational filters

The report will provide user-facing filters for:

- SiteID
- PartFamily
- CriticalityClass
- SupplierRiskClass
- InventoryStatus

These are report controls, not business-configuration parameters.

They therefore do not replace the configuration inputs maintained on `01_CONFIG`.

Default values:

- SiteID → ALL
- PartFamily → ALL
- CriticalityClass → ALL
- SupplierRiskClass → ALL
- InventoryStatus → ACTIONABLE

Inventory Status filter choices must support:

- ACTIONABLE
- ALL
- STOCKOUT
- CRITICAL
- REORDER
- ATTENTION
- EXCESS
- HEALTHY

`ACTIONABLE` means all statuses except `HEALTHY`.

This does not create a new Inventory Status classification.

It is only a reporting filter over the existing authoritative statuses.

### Default report population

The initial operational view will exclude `HEALTHY` positions.

Default included statuses:

1. STOCKOUT
2. CRITICAL
3. REORDER
4. ATTENTION
5. EXCESS

The user may explicitly select `HEALTHY` or `ALL` when required.

### Sorting

The default operational sort will apply:

1. Inventory Status priority;
2. Criticality Class;
3. Recommended Order Qty descending;
4. Backorder Qty descending;
5. SiteID;
6. ProductID.

Authoritative Inventory Status priority:

1. STOCKOUT
2. CRITICAL
3. REORDER
4. ATTENTION
5. EXCESS
6. HEALTHY

Criticality sorting:

1. A
2. B
3. C

### Formula approach

The report is expected to use Dynamic Array formulas.

Candidate functions include:

- LET
- FILTER
- SORTBY
- CHOOSECOLS
- XMATCH

Any new function introduced during implementation must be explained before use and its formula must be supplied using English Excel function names.

The final production formula must be validated against `tblReplenishment`.

### Output behavior

The spilled report output will be read-only from the user's perspective.

Users interact through report-filter controls rather than editing output rows.

Conditional Formatting may reinforce Inventory Status and exceptional conditions.

Text values must remain visible so status meaning does not depend only on color.

---

# 11. Management Dashboard Physical Specification

Worksheet:

`41_DASH_Management`

Primary user:

Procurement / Operations Manager

Primary display area:

`A:P`

The dashboard is intended to function as a concise management snapshot rather than a technical analytical worksheet.

---

## 11.1 Dashboard Context Area

Approximate area:

Rows 1–5.

Four management context cards will expose:

1. Reporting Date
2. Inventory Snapshot Date
3. Last Successful Refresh
4. Overall Quality Status

Approved source responsibilities:

Reporting Date:

`cfg_ReportingDate`

Inventory Snapshot Date:

validated `tblReplenishment` snapshot context.

Last Successful Refresh:

persistent Phase 9 automation state represented by `LastSuccessfulRefresh`.

Overall Quality Status:

existing Quality Control framework.

No duplicate workflow-state engine will be created on the dashboard.

---

## 11.2 Inventory and Replenishment KPI Area

Approximate area:

Rows 7–14.

The dashboard will expose these eight primary KPIs:

1. STOCKOUT Positions
2. CRITICAL Positions
3. REORDER Positions
4. Recommended Order Qty
5. Backorder Qty
6. Open PO Qty
7. On-Time Delivery Rate
8. Quality Incident Count

### KPI sources

STOCKOUT Positions:

`tblReplenishment[InventoryStatus]`

CRITICAL Positions:

`tblReplenishment[InventoryStatus]`

REORDER Positions:

`tblReplenishment[InventoryStatus]`

Recommended Order Qty:

`tblReplenishment[RecommendedOrderQty]`

Backorder Qty:

`tblReplenishment[BackorderQty]`

Open PO Qty:

`tblReplenishment[OpenPOQty]`

On-Time Delivery Rate:

Reporting-Date-safe Supplier Performance data.

Quality Incident Count:

Reporting-Date-safe Supplier Performance data.

---

## 11.3 Overall On-Time Delivery Calculation

The management On-Time Delivery Rate must not be calculated as a simple arithmetic average of individual Supplier percentages.

The approved overall interpretation is:

total On-Time received Purchase Orders
divided by
total received Purchase Orders

Source columns:

- `tblSupplierPerformance[OnTimePOCount]`
- `tblSupplierPerformance[ReceivedPOCount]`

Conceptually:

`SUM(OnTimePOCount) / SUM(ReceivedPOCount)`

This preserves the correct Purchase-Order-weighted denominator and remains Reporting-Date-safe.

The final Excel formula will be documented and validated during implementation.

---

## 11.4 Quality Incident KPI

Management Quality Incident Count will use:

`SUM(tblSupplierPerformance[QualityIncidentCount])`

because the Supplier Performance calculation layer already applies:

`IncidentDate <= ReportingDate`

This prevents future Quality Incident information from leaking into an earlier management snapshot.

---

# 12. Dashboard Visual Analysis

The baseline dashboard will contain two primary charts.

No new PivotTable is required for the Phase 10 baseline.

---

## 12.1 Inventory Status by Site

Source analytical object:

`pvtInvStatusBySite`

Worksheet:

`30_PVT_Inventory`

Validated source:

`tblReplenishment`

The dashboard chart will visualize Inventory Status distribution by Site.

The chart should allow management to identify where STOCKOUT, CRITICAL, REORDER and other statuses are concentrated.

A standard Excel chart referencing the validated PivotTable output is preferred.

Creating a new PivotTable merely to support this chart is not required.

---

## 12.2 Supplier Delivery Performance by Risk Class

Source analytical object:

`pvtSupplierRiskPerformance`

Worksheet:

`32_PVT_Suppliers`

Validated source:

`tblSupplierPerformance`

The dashboard chart will compare Supplier delivery-performance measures by Supplier Risk Class.

The primary displayed percentage measures should remain directly interpretable, such as:

- Average On-Time Delivery Rate
- Average Late Delivery Rate

The chart title and labels must make clear that these are risk-class analytical averages rather than the weighted overall management On-Time Delivery KPI.

A standard Excel chart referencing the validated PivotTable output is preferred.

---

# 13. Dashboard Interaction Strategy

The Phase 10 baseline will not introduce a universal dashboard slicer architecture.

Reason:

the dashboard combines management information whose validated analytical sources use different PivotCaches and source tables.

A control must not visually imply that it filters KPIs or charts that it does not actually control.

Interactive exploration remains available through the validated analytical worksheets:

- `30_PVT_Inventory`
- `31_PVT_Procurement`
- `32_PVT_Suppliers`

These worksheets already contain validated:

- Slicers
- Timelines
- Report Connections
- PivotCharts
- analytical PivotTables

The dashboard will provide clear navigation to these analytical views.

The operational replenishment report will provide its own dedicated report filters.

---

# 14. Dashboard Navigation

The management dashboard will provide user-facing navigation to:

- `00_HOME`
- `01_CONFIG`
- `40_RPT_Replenishment`
- `30_PVT_Inventory`
- `31_PVT_Procurement`
- `32_PVT_Suppliers`

The operational report will provide return navigation to:

- `00_HOME`
- `41_DASH_Management`

Navigation must be validated functionally before Phase 10 completion.

---

# 15. PivotTable and VBA Impact

The Phase 10 baseline does not require new PivotTables.

The existing Phase 9 production PivotTable contract currently validates:

- 10 analytical PivotTables;
- 5 unique PivotCaches;
- worksheets `30_PVT_Inventory`, `31_PVT_Procurement` and `32_PVT_Suppliers`.

Because no new PivotTable is planned in the baseline:

`modPivotRefresh`

should not require modification merely to build the Phase 10 presentation layer.

This remains subject to implementation validation.

If workbook implementation demonstrates that an additional PivotTable is genuinely necessary, the change must be explicitly designed and the Phase 9 refresh contract updated and revalidated.

---

# 16. Automation Integration

The existing production entry point remains:

`RefreshProcureFlow`

The Phase 10 baseline relies on the existing workflow to refresh upstream Power Query outputs, calculate Excel formulas, refresh analytical PivotTables and evaluate Quality Control.

Formula-driven report and dashboard outputs should update through the existing Excel calculation stages.

No additional production VBA procedure will be introduced unless implementation evidence demonstrates a real need.

After Phase 10 implementation, validation must prove that:

- the operational report reflects refreshed upstream calculations;
- dashboard KPI formulas reflect refreshed data;
- dashboard charts reflect refreshed analytical PivotTables;
- failed production refreshes cannot appear current;
- Last Successful Refresh remains truthful;
- Overall Quality Status remains visible.

---

# 17. Dashboard Visual Layout

The initial dashboard grid will use approximately columns:

`A:P`

### Context cards

Rows 3–5:

- Reporting Date
- Inventory Snapshot Date
- Last Successful Refresh
- Overall Quality Status

### KPI cards — row group 1

Rows 7–10:

- STOCKOUT Positions
- CRITICAL Positions
- REORDER Positions
- Recommended Order Qty

### KPI cards — row group 2

Rows 11–14:

- Backorder Qty
- Open PO Qty
- On-Time Delivery Rate
- Quality Incident Count

### Visual area

Approximately rows 16–29:

Left:

Inventory Status by Site.

Right:

Supplier Delivery Performance by Risk Class.

### Navigation area

Approximately rows 31–34.

The exact cell dimensions may be refined during workbook construction for readability and professional presentation without changing the approved information architecture.

---

# 18. Reporting Reconciliation Baseline

Phase 10 validation must reconcile dashboard and operational-report results to upstream validated sources.

Required baseline reconciliations include:

- operational report rows to `tblReplenishment`;
- STOCKOUT count to `tblReplenishment`;
- CRITICAL count to `tblReplenishment`;
- REORDER count to `tblReplenishment`;
- Backorder Qty to `tblReplenishment`;
- Open PO Qty to `tblReplenishment`;
- Recommended Order Qty to `tblReplenishment`;
- On-Time Delivery Rate to Supplier Performance counts;
- Quality Incident Count to `tblSupplierPerformance`;
- Inventory Status chart to `pvtInvStatusBySite`;
- Supplier delivery chart to `pvtSupplierRiskPerformance`;
- Reporting Date displayed on outputs to `cfg_ReportingDate`.

Historical Phase 7 baseline totals may be used as reference evidence where configuration and Reporting Date remain unchanged.

They must not be treated as permanent constants.

---

# 19. Physical Design Baseline State

Current Phase 10 implementation state:

Operational replenishment report:

[IMPLEMENTED] [VALIDATED]

Management dashboard:

[IMPLEMENTED] [VALIDATED]

The approved `40_RPT_Replenishment` and `41_DASH_Management` designs have now been physically implemented and manually validated.

Validated implementation evidence is recorded in:

- `docs/FORMULAS.md`
- `docs/TESTING.md`
- `docs/CURRENT_STATE.md`
- `docs/ARCHITECTURE.md`

Phase 10 technical implementation and validation are complete.

Formal Phase 10 completion remains pending only on the GitHub publication gate.

## 19.1 Implemented Chart Object Types

Inventory Status by Site:

- implemented as a PivotChart;
- reuses the existing validated `pvtInvStatusBySite`;
- reuses the established analytical PivotCache architecture;
- PivotChart field buttons are hidden in the management presentation;
- no new PivotTable or PivotCache was introduced.

Average Delivery Performance by Risk Class:

- implemented as a standard Excel 100% Stacked Column chart;
- references validated output from `pvtSupplierRiskPerformance`;
- displays Average On-Time Delivery Rate and Average Late Delivery Rate;
- does not introduce a new PivotTable or PivotCache.

The implemented chart-object choices preserve the approved Phase 10 analytical contract of 10 PivotTables and 5 PivotCaches.