# ProcureFlow — Formula Catalog

## Purpose

This document maintains the text-versioned representation of important Excel formulas implemented in ProcureFlow.

The executable formulas live in:

`workbook/ProcureFlow.xlsm`

This document exists because the workbook is a binary Git artifact and its formulas cannot be reviewed meaningfully through normal line-by-line Git diffs.

Only formulas that have actually been implemented are recorded as implemented.

Formula language standard:

English Excel function names.

---

# Phase 4 — Operational Model

## Worksheet

`20_CALC_Replenishment`

## Table

`tblReplenishment`

## Grain

1 Product × 1 Site

Validated population:

1,800 rows

---

## Production Formulas

### ProductID

**Status:** IMPLEMENTED

**Purpose:**

Generate the Product component of the Product × Site operational population.

**Formula:**

    =INDEX(tblProducts[ProductID],QUOTIENT(ROW()-4,ROWS(tblSites[SiteID]))+1)

**Sources:**

- `tblProducts[ProductID]`
- `tblSites[SiteID]`

**Validated behavior:**

Each Product appears once for every Site.

---

### SiteID

**Status:** IMPLEMENTED

**Purpose:**

Cycle through all Sites for each Product.

**Formula:**

    =INDEX(tblSites[SiteID],MOD(ROW()-4,ROWS(tblSites[SiteID]))+1)

**Sources:**

- `tblSites[SiteID]`

**Validated behavior:**

Each Product is associated with all six validated Sites.

---

### ProductSiteKey

**Status:** IMPLEMENTED

**Purpose:**

Create a deterministic technical key for the Product × Site operational grain.

**Formula:**

    =[@ProductID]&"|"&[@SiteID]

**Validated behavior:**

1,800 rows and 1,800 unique keys.

---

### PartFamily

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@ProductID],tblProducts[ProductID],tblProducts[PartFamily])

**Purpose:**

Expose the Product-owned Part Family attribute.

---

### CriticalityClass

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@ProductID],tblProducts[ProductID],tblProducts[CriticalityClass])

**Purpose:**

Expose the Product-owned Criticality Class.

---

### PrimarySupplierID

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@ProductID],tblProducts[ProductID],tblProducts[PrimarySupplierID])

**Purpose:**

Expose the Product's validated Primary Supplier.

---

### SupplierRiskClass

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@PrimarySupplierID],tblSuppliers[SupplierID],tblSuppliers[SupplierRiskClass])

**Purpose:**

Expose Supplier Risk from the Supplier dimension.

**Architecture note:**

`SupplierRiskClass` belongs logically to the Supplier and is therefore retrieved from `tblSuppliers`, not treated as a Product-owned attribute.

---

### UnitCost

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@ProductID],tblProducts[ProductID],tblProducts[UnitCost])

**Purpose:**

Expose Product Unit Cost.

---

### MasterLeadTimeDays

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@ProductID],tblProducts[ProductID],tblProducts[MasterLeadTimeDays])

**Purpose:**

Expose the master Lead-Time input for later Phase 5 business logic.

---

### ReportingDate

**Status:** IMPLEMENTED

**Formula:**

    =cfg_ReportingDate

**Purpose:**

Expose the configurable business Reporting Date to every Product × Site operational row.

---

### InventorySnapshotDate

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@ReportingDate],tblDate[Date],tblDate[WeekStartDate])

**Purpose:**

Resolve the weekly inventory snapshot corresponding to the configured Reporting Date.

**Decision:**

`DEC-051 — Inventory Snapshot Date`

**Validated behavior:**

A non-Monday Reporting Date of `2024-12-18` resolved to Inventory Snapshot Date `2024-12-16`.

---

### OnHandQty

**Status:** IMPLEMENTED

**Formula:**

    =SUMIFS(tblInventoryHistory[OnHandQty],tblInventoryHistory[WeekStartDate],[@InventorySnapshotDate],tblInventoryHistory[ProductID],[@ProductID],tblInventoryHistory[SiteID],[@SiteID])

**Purpose:**

Expose On-Hand Quantity for the Product × Site snapshot.

---

### BlockedQty

**Status:** IMPLEMENTED

**Formula:**

    =SUMIFS(tblInventoryHistory[BlockedQty],tblInventoryHistory[WeekStartDate],[@InventorySnapshotDate],tblInventoryHistory[ProductID],[@ProductID],tblInventoryHistory[SiteID],[@SiteID])

**Purpose:**

Expose Blocked Quantity for the Product × Site snapshot.

---

### BackorderQty

**Status:** IMPLEMENTED

**Formula:**

    =SUMIFS(tblInventoryHistory[BackorderQty],tblInventoryHistory[WeekStartDate],[@InventorySnapshotDate],tblInventoryHistory[ProductID],[@ProductID],tblInventoryHistory[SiteID],[@SiteID])

**Purpose:**

Expose Backorder Quantity for the Product × Site snapshot.

---

# Phase 4 Validation Formulas

The following formulas were used as implementation evidence and are not operational output columns.

## Operational Row Count

    =ROWS(tblReplenishment[ProductSiteKey])

Expected:

`1800`

---

## Expected Product × Site Population

    =ROWS(tblProducts[ProductID])*ROWS(tblSites[SiteID])

Expected:

`1800`

---

## Unique ProductSiteKey Count

    =ROWS(UNIQUE(tblReplenishment[ProductSiteKey]))

Expected:

`1800`

---

## Master-Attribute Formula Errors

    =SUMPRODUCT(--ISERROR(tblReplenishment[[PartFamily]:[MasterLeadTimeDays]]))

Expected:

`0`

---

## Distinct Primary Suppliers

    =ROWS(UNIQUE(tblReplenishment[PrimarySupplierID]))

Expected:

`40`

---

## Product Site Coverage Sample

    =COUNTIF(tblReplenishment[ProductID],"P00001")

Expected:

`6`

---

## Distinct Reporting Dates

    =ROWS(UNIQUE(tblReplenishment[ReportingDate]))

Expected:

`1`

---

## Distinct Inventory Snapshot Dates

    =ROWS(UNIQUE(tblReplenishment[InventorySnapshotDate]))

Expected:

`1`

---

## Snapshot Causality

    =MAX(tblReplenishment[InventorySnapshotDate])<=MIN(tblReplenishment[ReportingDate])

Expected:

`TRUE`

---

## Snapshot-Date Relationship

    =INDEX(tblReplenishment[InventorySnapshotDate],1)=XLOOKUP(cfg_ReportingDate,tblDate[Date],tblDate[WeekStartDate])

Expected:

`TRUE`

---

## DEC-051 Non-Monday Test

    =INDEX(tblReplenishment[InventorySnapshotDate],1)=DATE(2024,12,16)

Used temporarily with:

`cfg_ReportingDate = 2024-12-18`

Expected:

`TRUE`

---

## Inventory Snapshot Source Rows

    =COUNTIF(tblInventoryHistory[WeekStartDate],INDEX(tblReplenishment[InventorySnapshotDate],1))

Expected:

`1800`

---

## Product × Site Snapshot Matches

    =SUMPRODUCT(COUNTIFS(tblInventoryHistory[WeekStartDate],tblReplenishment[InventorySnapshotDate],tblInventoryHistory[ProductID],tblReplenishment[ProductID],tblInventoryHistory[SiteID],tblReplenishment[SiteID]))

Expected:

`1800`

---

## On-Hand Reconciliation

    =SUM(tblReplenishment[OnHandQty])=SUMIFS(tblInventoryHistory[OnHandQty],tblInventoryHistory[WeekStartDate],INDEX(tblReplenishment[InventorySnapshotDate],1))

Expected:

`TRUE`

---

## Blocked Quantity Reconciliation

    =SUM(tblReplenishment[BlockedQty])=SUMIFS(tblInventoryHistory[BlockedQty],tblInventoryHistory[WeekStartDate],INDEX(tblReplenishment[InventorySnapshotDate],1))

Expected:

`TRUE`

---

## Backorder Reconciliation

    =SUM(tblReplenishment[BackorderQty])=SUMIFS(tblInventoryHistory[BackorderQty],tblInventoryHistory[WeekStartDate],INDEX(tblReplenishment[InventorySnapshotDate],1))

Expected:

`TRUE`

---

## Blocked Quantity Integrity

    =SUMPRODUCT(--(tblReplenishment[BlockedQty]>tblReplenishment[OnHandQty]))

Expected:

`0`

---

# Phase Boundary

Phase 5 formulas are intentionally not documented as implemented yet.

They will be added to this catalog only after actual workbook implementation and validation.

Future Phase 5 formula areas include:

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
- Supplier-performance business metrics