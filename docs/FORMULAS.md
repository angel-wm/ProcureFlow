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

### DemandHistoryWeeks

**Status:** IMPLEMENTED

**Formula:**

    =cfg_DemandHistoryWeeks

**Purpose:**

Expose the configurable number of completed historical-demand weeks used by later Phase 5 calculations.

---

### DemandHistoryStartDate

**Status:** IMPLEMENTED

**Formula:**

    =[@InventorySnapshotDate]-7*[@DemandHistoryWeeks]

**Purpose:**

Resolve the first completed weekly period included in the historical-demand window.

---

### DemandHistoryEndDate

**Status:** IMPLEMENTED

**Formula:**

    =[@InventorySnapshotDate]-7

**Purpose:**

Resolve the final completed weekly period immediately preceding the current Inventory Snapshot Date.

**Decision:**

`DEC-052 — Completed Demand History Window`

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
---

# Phase 4 — Supplier Performance Structural Model

## Worksheet — 21_CALC_SupplierPerformance

## Table

`tblSupplierPerformance`

## Grain

1 row = 1 Supplier

Validated population:

40 rows

---

### SupplierID

**Status:** IMPLEMENTED

**Formula:**

    =INDEX(tblSuppliers[SupplierID],ROW()-3)

**Purpose:**

Populate the structural Supplier grain from `tblSuppliers`.

---

### SupplierRiskClass

**Status:** IMPLEMENTED

**Formula:**

    =XLOOKUP([@SupplierID],tblSuppliers[SupplierID],tblSuppliers[SupplierRiskClass])

**Purpose:**

Expose the validated Supplier Risk classification.

---

### ReportingDate

**Status:** IMPLEMENTED

**Formula:**

    =cfg_ReportingDate

**Purpose:**

Expose the configurable Reporting Date to the Supplier Performance operational structure.

---

## Validation Formulas

### Supplier Row Count

    =ROWS(tblSupplierPerformance[SupplierID])

Expected:

`40`

---

### Unique Supplier Count

    =ROWS(UNIQUE(tblSupplierPerformance[SupplierID]))

Expected:

`40`

---

### Supplier Population Reconciliation

    =ROWS(tblSupplierPerformance[SupplierID])=ROWS(tblSuppliers[SupplierID])

Expected:

`TRUE`

---

### Master-Data Errors

    =SUMPRODUCT(--ISERROR(tblSupplierPerformance[[SupplierID]:[SupplierRiskClass]]))

Expected:

`0`

---

### Distinct Reporting Dates

    =ROWS(UNIQUE(tblSupplierPerformance[ReportingDate]))

Expected:

`1`
## Historical Demand Window Validation

### Distinct Demand History Week Configurations

    =ROWS(UNIQUE(tblReplenishment[DemandHistoryWeeks]))

Expected:

`1`

---

### Rows Using 26 Weeks

    =COUNTIF(tblReplenishment[DemandHistoryWeeks],26)

Expected:

`1800`

---

### Distinct Demand History Start Dates

    =ROWS(UNIQUE(tblReplenishment[DemandHistoryStartDate]))

Expected:

`1`

---

### Distinct Demand History End Dates

    =ROWS(UNIQUE(tblReplenishment[DemandHistoryEndDate]))

Expected:

`1`

---

### Configured Window Length

    =(INDEX(tblReplenishment[DemandHistoryEndDate],1)-INDEX(tblReplenishment[DemandHistoryStartDate],1))/7+1=cfg_DemandHistoryWeeks

Expected:

`TRUE`

---

### Completed-Week Boundary

    =INDEX(tblReplenishment[DemandHistoryEndDate],1)=INDEX(tblReplenishment[InventorySnapshotDate],1)-7

Expected:

`TRUE`

---

### Distinct Source Weeks

    =ROWS(UNIQUE(FILTER(tblInventoryHistory[WeekStartDate],(tblInventoryHistory[WeekStartDate]>=INDEX(tblReplenishment[DemandHistoryStartDate],1))*(tblInventoryHistory[WeekStartDate]<=INDEX(tblReplenishment[DemandHistoryEndDate],1)))))

Expected:

`26`

---

### Historical Source Row Count

    =COUNTIFS(tblInventoryHistory[WeekStartDate],">="&INDEX(tblReplenishment[DemandHistoryStartDate],1),tblInventoryHistory[WeekStartDate],"<="&INDEX(tblReplenishment[DemandHistoryEndDate],1))

Expected:

`46800`
---

## Final Phase 4 Refresh Validation

### Replenishment Row Count

    =ROWS(tblReplenishment[ProductSiteKey])

Expected:

`1800`

---

### Unique Product-Site Keys

    =ROWS(UNIQUE(tblReplenishment[ProductSiteKey]))

Expected:

`1800`

---

### Replenishment Master/Input Errors

    =SUMPRODUCT(--ISERROR(tblReplenishment[[PartFamily]:[BackorderQty]]))

Expected:

`0`

---

### Reporting Date Relationship

    =INDEX(tblReplenishment[ReportingDate],1)=cfg_ReportingDate

Expected:

`TRUE`

---

### Historical-Demand Source Coverage

    =COUNTIFS(tblInventoryHistory[WeekStartDate],">="&INDEX(tblReplenishment[DemandHistoryStartDate],1),tblInventoryHistory[WeekStartDate],"<="&INDEX(tblReplenishment[DemandHistoryEndDate],1))

Expected:

`46800`

---

### Supplier Performance Row Count

    =ROWS(tblSupplierPerformance[SupplierID])

Expected:

`40`

---

### Unique Supplier Count

    =ROWS(UNIQUE(tblSupplierPerformance[SupplierID]))

Expected:

`40`

---

### Supplier Structural Errors

    =SUMPRODUCT(--ISERROR(tblSupplierPerformance[[SupplierID]:[SupplierRiskClass]]))

Expected:

`0`

## Phase 5 — Business Logic & Advanced Formulas

Status: IMPLEMENTED AND VALIDATED

### Replenishment Engine

The following formulas are implemented in `tblReplenishment` on `20_CALC_Replenishment`.

#### HistoricalDemandQty

    =SUMIFS(tblInventoryHistory[ConsumptionQty],tblInventoryHistory[ProductID],[@ProductID],tblInventoryHistory[SiteID],[@SiteID],tblInventoryHistory[WeekStartDate],">="&[@DemandHistoryStartDate],tblInventoryHistory[WeekStartDate],"<="&[@DemandHistoryEndDate])

Purpose: total historical consumption for the configured Product × Site demand-history window.

#### AverageWeeklyDemand

    =[@HistoricalDemandQty]/[@DemandHistoryWeeks]

Purpose: average weekly consumption using the complete configured history window, including zero-consumption weeks.

#### DemandStdDev

    =STDEV.S(FILTER(tblInventoryHistory[ConsumptionQty],(tblInventoryHistory[ProductID]=[@ProductID])*(tblInventoryHistory[SiteID]=[@SiteID])*(tblInventoryHistory[WeekStartDate]>=[@DemandHistoryStartDate])*(tblInventoryHistory[WeekStartDate]<=[@DemandHistoryEndDate])))

Purpose: sample standard deviation of weekly demand for the configured historical window.

#### AvgActualLeadTimeDays

    =IF(COUNTIFS(tblPurchaseOrders[ProductID],[@ProductID],tblPurchaseOrders[SiteID],[@SiteID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate])=0,"",AVERAGEIFS(tblPurchaseOrders[ActualLeadTimeDays],tblPurchaseOrders[ProductID],[@ProductID],tblPurchaseOrders[SiteID],[@SiteID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate]))

Purpose: average historical Actual Lead Time using only receipts known by Reporting Date.

#### LeadTimeStdDevDays

    =IF(COUNTIFS(tblPurchaseOrders[ProductID],[@ProductID],tblPurchaseOrders[SiteID],[@SiteID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate])<2,0,STDEV.S(FILTER(tblPurchaseOrders[ActualLeadTimeDays],(tblPurchaseOrders[ProductID]=[@ProductID])*(tblPurchaseOrders[SiteID]=[@SiteID])*(tblPurchaseOrders[ReceiptDate]<=[@ReportingDate]))))

Purpose: sample standard deviation of historical Actual Lead Time. Fewer than two observations returns zero variability.

#### EffectiveLeadTimeDays

    =IF([@AvgActualLeadTimeDays]="",[@MasterLeadTimeDays],[@AvgActualLeadTimeDays])

Purpose: use observed historical Lead Time when available and master Lead Time as fallback otherwise.

#### OpenPOQty

    =SUMIFS(tblPurchaseOrders[OrderedQty],tblPurchaseOrders[ProductID],[@ProductID],tblPurchaseOrders[SiteID],[@SiteID],tblPurchaseOrders[OrderDate],"<="&[@ReportingDate],tblPurchaseOrders[ReceiptDate],">"&[@ReportingDate])

Purpose: incoming quantity from Purchase Orders that were open at Reporting Date according to `OrderDate <= ReportingDate < ReceiptDate`.

#### AvailableStock

    =MAX(0,[@OnHandQty]-[@BlockedQty])

Purpose: usable inventory after blocked stock, with zero as the minimum.

#### ServiceLevel

    =IFS([@CriticalityClass]="A",cfg_ServiceLevelA,[@CriticalityClass]="B",cfg_ServiceLevelB,[@CriticalityClass]="C",cfg_ServiceLevelC)

Purpose: retrieve the configured Service Level for each Criticality Class without hidden constants.

#### SafetyStock

    =NORM.S.INV([@ServiceLevel])*SQRT(([@EffectiveLeadTimeDays]/7)*([@DemandStdDev]^2)+([@AverageWeeklyDemand]^2)*(([@LeadTimeStdDevDays]/7)^2))

Purpose: Safety Stock incorporating both demand variability and Lead-Time variability with consistent weekly time units.

#### ReorderPoint

    =([@AverageWeeklyDemand]*([@EffectiveLeadTimeDays]/7))+[@SafetyStock]

Purpose: expected demand during Lead Time plus Safety Stock.

#### InventoryPosition

    =[@AvailableStock]+[@OpenPOQty]-[@BackorderQty]

Purpose: net replenishment position including usable stock, incoming open POs and backorders.

#### TargetStock

    =[@AverageWeeklyDemand]*(([@EffectiveLeadTimeDays]/7)+cfg_ReviewPeriodWeeks)+[@SafetyStock]

Purpose: stock target covering Lead Time, configured Review Period and Safety Stock.

#### RecommendedOrderQty

    =IF(AND([@NoRecentDemand],[@BackorderQty]=0),0,IF([@InventoryPosition]<=[@ReorderPoint],MAX(0,ROUNDUP([@TargetStock]-[@InventoryPosition],0)),0))

Purpose: recommend whole units up to Target Stock only when replenishment is required. No-recent-demand rows without a backorder do not receive an automatic statistical purchase recommendation.

#### InventoryStatus

    =IFS(AND([@AvailableStock]=0,[@BackorderQty]>0),"STOCKOUT",OR([@BackorderQty]>0,[@AvailableStock]<[@SafetyStock]),"CRITICAL",[@InventoryPosition]<=[@ReorderPoint],"REORDER",[@InventoryPosition]<=[@ReorderPoint]+[@AverageWeeklyDemand],"ATTENTION",[@InventoryPosition]>[@TargetStock]+([@AverageWeeklyDemand]*cfg_ExcessBufferWeeks),"EXCESS",TRUE,"HEALTHY")

Priority order: STOCKOUT, CRITICAL, REORDER, ATTENTION, EXCESS, HEALTHY.

#### NoRecentDemand

    =[@HistoricalDemandQty]=0

Purpose: Boolean flag identifying Product × Site combinations with zero consumption across the configured historical window.

### Supplier Performance

The following formulas are implemented in `tblSupplierPerformance` on `21_CALC_SupplierPerformance`.

#### ReceivedPOCount

    =COUNTIFS(tblPurchaseOrders[SupplierID],[@SupplierID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate])

#### OnTimePOCount

    =COUNTIFS(tblPurchaseOrders[SupplierID],[@SupplierID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate],tblPurchaseOrders[IsLateReceipt],FALSE)

#### OnTimeDeliveryRate

    =IF([@ReceivedPOCount]=0,"",[@OnTimePOCount]/[@ReceivedPOCount])

#### LatePOCount

    =COUNTIFS(tblPurchaseOrders[SupplierID],[@SupplierID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate],tblPurchaseOrders[IsLateReceipt],TRUE)

#### LateDeliveryRate

    =IF([@ReceivedPOCount]=0,"",[@LatePOCount]/[@ReceivedPOCount])

#### PartialPOCount

    =COUNTIFS(tblPurchaseOrders[SupplierID],[@SupplierID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate],tblPurchaseOrders[IsPartialReceipt],TRUE)

#### PartialReceiptRate

    =IF([@ReceivedPOCount]=0,"",[@PartialPOCount]/[@ReceivedPOCount])

#### AvgActualLeadTimeDays

    =IF([@ReceivedPOCount]=0,"",AVERAGEIFS(tblPurchaseOrders[ActualLeadTimeDays],tblPurchaseOrders[SupplierID],[@SupplierID],tblPurchaseOrders[ReceiptDate],"<="&[@ReportingDate]))

#### LeadTimeStdDevDays

    =IF([@ReceivedPOCount]<2,0,STDEV.S(FILTER(tblPurchaseOrders[ActualLeadTimeDays],(tblPurchaseOrders[SupplierID]=[@SupplierID])*(tblPurchaseOrders[ReceiptDate]<=[@ReportingDate]))))

#### QualityIncidentCount

    =COUNTIFS(tblQualityIncidents[SupplierID],[@SupplierID],tblQualityIncidents[IncidentDate],"<="&[@ReportingDate])

No aggregate Supplier Score is introduced in Phase 5.
