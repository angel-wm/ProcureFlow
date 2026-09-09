# ProcureFlow — Phase 4 Closeout

## Phase

Phase 4 — Operational Model

## Status

COMPLETED

## Target Version

`v0.5.0`

## Objective

Build the structured operational calculation model required before implementing final business formulas.

## Completion Summary

Phase 4 has completed its technical implementation and validation work.

The operational layer now provides the structural grains, source-supported attributes, temporal context and prepared inputs required by Phase 5 — Business Logic & Advanced Formulas.

Implemented Phase 4 capabilities include:

- validated Product × Site operational grain;
- physical `tblReplenishment`;
- validated Supplier operational grain;
- physical `tblSupplierPerformance`;
- Product and Supplier master attributes;
- configurable Reporting Date exposure;
- weekly Inventory Snapshot Date resolution;
- current On-Hand Quantity exposure;
- Blocked Quantity exposure;
- Backorder Quantity exposure;
- completed historical-demand window preparation;
- master Lead-Time exposure;
- Purchase Order inputs available for Phase 5;
- historical Actual Lead-Time input available for Phase 5;
- formal Phase 4 formula catalog;
- workbook Refresh validation;
- Phase 4 structural test evidence.

No Phase 5 replenishment or supplier-performance business formula has been implemented prematurely.

---

## Workbook

Executable workbook:

`workbook/ProcureFlow.xlsm`

Target application:

Microsoft Excel 365 Desktop for Windows.

---

## Replenishment Operational Model

Worksheet:

`20_CALC_Replenishment`

Excel Table:

`tblReplenishment`

Operational grain:

1 Product × 1 Site

Validated source population:

- Products: 300
- Sites: 6
- Product × Site combinations: 1,800

Validated table population:

- rows: 1,800
- unique `ProductSiteKey` values: 1,800

### Implemented Phase 4 Fields

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
- `DemandHistoryWeeks`
- `DemandHistoryStartDate`
- `DemandHistoryEndDate`
- `OnHandQty`
- `BlockedQty`
- `BackorderQty`

### Reserved Phase 5 Fields

The following business-calculation areas remain assigned to Phase 5:

- Average Weekly Demand
- demand variability
- average Actual Lead Time
- Lead-Time variability
- Effective Lead Time
- Open Purchase Order Quantity
- Available Stock
- Service Level
- Safety Stock
- Reorder Point
- Inventory Position
- Target Stock
- Recommended Order Quantity
- Inventory Status
- `NO_RECENT_DEMAND`

The existence of a destination column does not imply implementation.

---

## Supplier Performance Operational Model

Worksheet:

`21_CALC_SupplierPerformance`

Excel Table:

`tblSupplierPerformance`

Operational grain:

1 row = 1 Supplier

Validated population:

- Suppliers: 40
- table rows: 40
- unique Supplier IDs: 40

### Implemented Phase 4 Fields

- `SupplierID`
- `SupplierRiskClass`
- `ReportingDate`

### Reserved Phase 5 Metric Areas

The structural model reserves later metric areas including:

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

## Reporting Date Architecture

Phase 4 implements the distinction between:

- `ReportingDate`
- `InventorySnapshotDate`

`ReportingDate` is driven by:

`cfg_ReportingDate`

Inventory History remains weekly at Monday-based `WeekStartDate`.

The approved inventory snapshot rule is:

`InventorySnapshotDate = latest WeekStartDate <= ReportingDate`

The operational implementation resolves this relationship through `tblDate`.

This behavior is governed by:

`DEC-051 — Inventory Snapshot Date`

### Validation Evidence

A temporary non-Monday Reporting Date was tested:

- Reporting Date: `2024-12-18`
- expected Inventory Snapshot Date: `2024-12-16`
- actual Inventory Snapshot Date: `2024-12-16`

Result:

PASS

The Reporting Date was restored to:

`2024-12-23`

---

## Current Inventory Inputs

Current operational inventory is exposed by Product × Site using:

- `OnHandQty`
- `BlockedQty`
- `BackorderQty`

The current validated Inventory Snapshot Date is:

`2024-12-23`

Validated snapshot source rows:

1,800

Validated Product × Site source matches:

1,800

Reconciliations:

- On-Hand Quantity: PASS
- Blocked Quantity: PASS
- Backorder Quantity: PASS

Blocked Quantity greater than On-Hand Quantity exceptions:

0

---

## Completed Historical Demand Window

Phase 4 prepares, but does not calculate, historical demand.

Implemented temporal fields:

- `DemandHistoryWeeks`
- `DemandHistoryStartDate`
- `DemandHistoryEndDate`

Current configured window:

- Demand History Weeks: 26
- Demand History Start Date: `2024-06-24`
- Demand History End Date: `2024-12-16`

The historical-demand interval uses completed weekly periods preceding the Inventory Snapshot Date.

Validated source coverage:

- distinct weekly periods: 26
- Product × Site combinations per week: 1,800
- total source observations: 46,800

This behavior is governed by:

`DEC-052 — Completed Demand History Window`

No Average Demand, demand variability or `NO_RECENT_DEMAND` calculation is implemented in Phase 4.

---

## Lead-Time Inputs

Phase 4 prepares the inputs required for later Lead-Time calculations.

Available Product-level input:

- `MasterLeadTimeDays`

Available Purchase Order fields from the Phase 3 pipeline include:

- `SupplierID`
- `SiteID`
- `ProductID`
- `OrderDate`
- `PromisedDate`
- `ReceiptDate`
- `OrderedQty`
- `ReceivedQty`
- `PromisedLeadTimeDays`
- `ActualLeadTimeDays`
- `IsLateReceipt`
- `IsPartialReceipt`

Phase 5 remains responsible for:

- historical Lead-Time aggregation;
- average Actual Lead Time;
- Lead-Time variability;
- Effective Lead-Time fallback logic.

---

## Open Purchase Order Inputs

Phase 4 confirms that the transactional inputs required for later Reporting-Date-dependent Open PO calculations are available.

Required fields already available in `tblPurchaseOrders` include:

- `ProductID`
- `SiteID`
- `OrderDate`
- `ReceiptDate`
- `OrderedQty`

The approved Open PO business rule remains:

Order Date <= Reporting Date < Receipt Date

The Phase 3 pipeline intentionally does not persist `IsOpenPO`.

Phase 5 will calculate Reporting-Date-dependent Open Purchase Order Quantity.

---

## Formula Catalog

Phase 4 introduced:

`docs/FORMULAS.md`

The catalog maintains the text-versioned representation of implemented Excel formulas because the `.xlsm` workbook is a binary Git artifact.

The catalog contains:

- production formulas;
- purpose and source references;
- validation formulas;
- expected validation results;
- phase boundaries.

Excel formulas are documented using English function names.

---

## Phase 4 Decisions

Confirmed during Phase 4:

- `DEC-051` — Inventory Snapshot Date
- `DEC-052` — Completed Demand History Window

Earlier confirmed decisions remain applicable.

No decision is currently superseded.

---

## Testing Evidence

Formal evidence is documented in:

`docs/TESTING.md`

Validated Phase 4 areas include:

- Product × Site grain;
- unique Product-Site keys;
- master attributes;
- Supplier population;
- Reporting Date behavior;
- non-Monday inventory snapshot behavior;
- current inventory reconciliation;
- completed historical-demand window;
- historical source-row coverage;
- Supplier structural model;
- full workbook Refresh.

Current technical Phase 4 testing result:

PASS

---

## Final Refresh Validation

A full Excel `Refresh All` was executed after the Phase 4 implementation.

Refresh result:

PASS

Post-refresh validation results:

| Test | Expected | Result |
|---|---:|---:|
| `tblReplenishment` rows | 1,800 | 1,800 |
| Unique `ProductSiteKey` values | 1,800 | 1,800 |
| Replenishment master/input errors | 0 | 0 |
| Reporting Date relationship | TRUE | TRUE |
| Historical-demand source rows | 46,800 | 46,800 |
| `tblSupplierPerformance` rows | 40 | 40 |
| Unique Supplier IDs | 40 | 40 |
| Supplier structural errors | 0 | 0 |

Result:

PASS

No critical Phase 4 structural defect is currently known.

---

## Exit Criteria Review

### Operational grains are correct

Status:

PASS

Validated grains:

- `tblReplenishment`: 1 Product × 1 Site
- `tblSupplierPerformance`: 1 Supplier

### Required inputs are available

Status:

PASS

Prepared areas include:

- current inventory;
- blocked stock;
- backorders;
- Reporting Date;
- Inventory Snapshot Date;
- historical-demand temporal context;
- Product attributes;
- Supplier attributes;
- master Lead Time;
- Purchase Order inputs;
- historical Actual Lead-Time inputs.

### No unresolved critical structural issue remains

Status:

PASS

No unresolved critical structural defect is currently known.

### Model is ready for business-rule formulas

Status:

PASS

The operational structure is ready for Phase 5.

---

## Phase Boundary

Phase 4 does not implement:

- final historical-demand calculations;
- Safety Stock;
- Reorder Point;
- Open PO Quantity;
- Available Stock;
- Inventory Position;
- Target Stock;
- Recommended Order Quantity;
- Inventory Status;
- `NO_RECENT_DEMAND`;
- final supplier-performance metrics;
- PivotTables;
- VBA automation;
- final reporting;
- management dashboard.

These remain assigned to later roadmap phases.

---

## Technical Phase Result

Phase 4 technical implementation:

COMPLETE

Phase 4 technical validation:

PASS

Formal project status:

COMPLETED

---

## Git Evidence

Phase branch:

`phase/04-operational-model`

Target branch:

`main`

Target version:

`v0.5.0`

Phase Pull Request:

`#5 — Phase 4 — Operational Model`

Pull Request status:

MERGED

Merge commit:

`7d3f11fdb069436dd767fd1c371be169c31b2804`

Phase completion tag:

``phase-4-complete``

Version tag:

``v0.5.0``

---

## GitHub Gate

Status:

COMPLETED

Completed evidence:

1. Phase 4 technical implementation completed;
2. Phase 4 validation completed with PASS result;
3. Phase 4 closeout documentation prepared and reviewed;
4. `phase/04-operational-model` published;
5. Pull Request `#5 — Phase 4 — Operational Model` created;
6. complete Pull Request diff reviewed;
7. review findings corrected and re-reviewed;
8. Pull Request `#5` merged into `main`;
9. merge commit confirmed as `7d3f11fdb069436dd767fd1c371be169c31b2804`;
10. canonical Phase 4 completion state prepared;
11. completion tags defined as `phase-4-complete` and `v0.5.0`.

The final release commit and both completion tags have been published and verified on GitHub. Phase 4 is formally COMPLETED.
---

## Next Phase

Phase 5 — Business Logic & Advanced Formulas

Status:

NOT STARTED

Target Version:

`v0.6.0`

Phase 5 is authorized to begin through its dedicated phase workflow.
