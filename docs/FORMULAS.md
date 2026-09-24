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

---

# Phase 6 — Quality Control System Formulas

Status:

IMPLEMENTED AND VALIDATED

Worksheet:

`02_CONTROL`

Primary table:

`tblQualityControl`

## Status Formula

Implemented in `tblQualityControl[Status]`:

    =IF([@Severity]="N/A","N/A",IF([@ExceptionCount]="","",IF([@ExceptionCount]=0,"PASS",IF([@Severity]="Warning","WARNING","FAIL"))))

Purpose:

- preserve N/A controls;
- leave unevaluated applicable controls blank;
- return PASS for zero exceptions;
- return WARNING or FAIL according to configured severity.

## Overall Quality Status

Implemented summary formula:

    =LET(StatusRange,tblQualityControl[Status],Applicable,ROWS(tblQualityControl[ControlID])-COUNTIF(tblQualityControl[Severity],"N/A"),Evaluated,COUNTIF(StatusRange,"PASS")+COUNTIF(StatusRange,"WARNING")+COUNTIF(StatusRange,"FAIL"),IF(COUNTIF(StatusRange,"FAIL")>0,"FAIL",IF(Evaluated<Applicable,"WARNING",IF(COUNTIF(StatusRange,"WARNING")>0,"WARNING","PASS"))))

Purpose:

Apply the precedence:

`FAIL` > `WARNING` > `PASS`

while preventing an incomplete set of applicable controls from reporting PASS.

## Summary Metrics

PASS controls:

    =COUNTIF(tblQualityControl[Status],"PASS")

WARNING controls:

    =COUNTIF(tblQualityControl[Status],"WARNING")

FAIL controls:

    =COUNTIF(tblQualityControl[Status],"FAIL")

Applicable controls:

    =ROWS(tblQualityControl[ControlID])-COUNTIF(tblQualityControl[Severity],"N/A")

N/A controls:

    =COUNTIF(tblQualityControl[Status],"N/A")

Total exceptions:

    =SUM(tblQualityControl[ExceptionCount])

Phase 6 current valid technical refresh (historical baseline):

    =XLOOKUP("QC-030",tblQualityControl[ControlID],tblQualityControl[Result],"")

## Technical Power Query Feed Lookups — Phase 6 Baseline

For `QC-001`, `QC-002`, `QC-003`, `QC-004`, `QC-030` and `QC-031`:

Result:

    =XLOOKUP([@ControlID],tblQCPipelineHealth[ControlID],tblQCPipelineHealth[Result],"")

Expected:

    =XLOOKUP([@ControlID],tblQCPipelineHealth[ControlID],tblQCPipelineHealth[Expected],"")

ExceptionCount:

    =XLOOKUP([@ControlID],tblQCPipelineHealth[ControlID],tblQCPipelineHealth[ExceptionCount],"")

## Representative Integrity Controls

Duplicate Product IDs:

    =ROWS(tblProducts[ProductID])-ROWS(UNIQUE(tblProducts[ProductID]))

Duplicate Purchase Order IDs:

    =ROWS(tblPurchaseOrders[PurchaseOrderID])-ROWS(UNIQUE(tblPurchaseOrders[PurchaseOrderID]))

Duplicate Quality Incident IDs:

    =ROWS(tblQualityIncidents[QualityIncidentID])-ROWS(UNIQUE(tblQualityIncidents[QualityIncidentID]))

Invalid Inventory Product references:

    =SUMPRODUCT(--ISNA(MATCH(tblInventoryHistory[ProductID],tblProducts[ProductID],0)))

Invalid Purchase Order Product references:

    =SUMPRODUCT(--ISNA(MATCH(tblPurchaseOrders[ProductID],tblProducts[ProductID],0)))

Invalid Quality Product references:

    =SUMPRODUCT(--ISNA(MATCH(tblQualityIncidents[ProductID],tblProducts[ProductID],0)))

## Representative Business-Rule Controls

Non-positive Ordered Quantity:

    =COUNTIF(tblPurchaseOrders[OrderedQty],"<=0")

Received Quantity greater than Ordered Quantity:

    =SUMPRODUCT(--(tblPurchaseOrders[ReceivedQty]>tblPurchaseOrders[OrderedQty]))

Negative Blocked Quantity:

    =COUNTIF(tblInventoryHistory[BlockedQty],"<0")

Blocked Quantity greater than On-Hand Quantity:

    =SUMPRODUCT(--(tblInventoryHistory[BlockedQty]>tblInventoryHistory[OnHandQty]))

Receipt Date before Order Date:

    =SUMPRODUCT(--(tblPurchaseOrders[ReceiptDate]<tblPurchaseOrders[OrderDate]))

Promised Date before Order Date:

    =SUMPRODUCT(--(tblPurchaseOrders[PromisedDate]<tblPurchaseOrders[OrderDate]))

## Configuration Controls

Demand History Weeks:

    =--NOT(AND(ISNUMBER(cfg_DemandHistoryWeeks),cfg_DemandHistoryWeeks>0,cfg_DemandHistoryWeeks=INT(cfg_DemandHistoryWeeks)))

Service Levels:

    =--OR(cfg_ServiceLevelA<=0,cfg_ServiceLevelA>=1)+--OR(cfg_ServiceLevelB<=0,cfg_ServiceLevelB>=1)+--OR(cfg_ServiceLevelC<=0,cfg_ServiceLevelC>=1)

Review Period:

    =--NOT(AND(ISNUMBER(cfg_ReviewPeriodWeeks),cfg_ReviewPeriodWeeks>0,cfg_ReviewPeriodWeeks=INT(cfg_ReviewPeriodWeeks)))

Excess Buffer:

    =--NOT(AND(ISNUMBER(cfg_ExcessBufferWeeks),cfg_ExcessBufferWeeks>=0,cfg_ExcessBufferWeeks=INT(cfg_ExcessBufferWeeks)))

The Service Level Quality Control uses the stricter operational requirement:

`0 < Service Level < 1`

because `NORM.S.INV()` cannot safely operate at exactly 0% or 100%.

## Extended Phase 6 Controls

Duplicate Inventory composite keys:

    =ROWS(tblInventoryHistory[ProductID])-ROWS(UNIQUE(tblInventoryHistory[WeekStartDate]&"|"&tblInventoryHistory[SiteID]&"|"&tblInventoryHistory[ProductID]))

Date continuity exceptions:

    =ABS((MAX(tblDate[Date])-MIN(tblDate[Date])+1)-ROWS(UNIQUE(tblDate[Date])))+(ROWS(tblDate[Date])-ROWS(UNIQUE(tblDate[Date])))

Calculation formula errors:

    =SUMPRODUCT(--ISERROR(tblReplenishment[#Data]))+SUMPRODUCT(--ISERROR(tblSupplierPerformance[#Data]))

## Phase 6 Validation Baseline

Validated final baseline:

- 35 defined controls;
- 34 applicable controls;
- 34 PASS;
- 0 WARNING;
- 0 FAIL;
- 1 N/A;
- total exceptions: 0;
- Overall Quality Status: PASS.

`QC-032 — PivotTable refresh status` remained N/A at the Phase 6 baseline pending the analytical and automation layers. This statement is retained as historical Phase 6 evidence.

## Phase 9 — Automation-Controlled Quality Control Formulas

Phase 9 migrated the automation-dependent Quality Control logic to persistent state in `tblAutomationState`.

`QC-030` now reads `LastSuccessfulRefresh` from persistent automation state rather than using the current Power Query evaluation timestamp.

`QC-032` now reads the persistent `PivotRefreshStatus` maintained by the production automation workflow.

### QC-030 — Last Successful Refresh

Result:

    =LET(last,XLOOKUP("LastSuccessfulRefresh",tblAutomationState[StateKey],tblAutomationState[StateValue],""),IF(OR(last="",last=0),"Not available",last))

Expected:

`Accepted production refresh`

ExceptionCount:

    =LET(s,XLOOKUP("WorkflowStatus",tblAutomationState[StateKey],tblAutomationState[StateValue],""),last,XLOOKUP("LastSuccessfulRefresh",tblAutomationState[StateKey],tblAutomationState[StateValue],""),IF(OR(s="",s="NOT_RUN",s="RUNNING"),"",--NOT(AND(OR(s="SUCCESS",s="WARNING"),last<>"",last<>0))))

Status:

    =IF([@Severity]="N/A","N/A",IF([@ExceptionCount]="","",IF([@ExceptionCount]=0,"PASS",IF([@Severity]="Warning","WARNING","FAIL"))))

### QC-032 — PivotTable Refresh Status

Result:

    =XLOOKUP("PivotRefreshStatus",tblAutomationState[StateKey],tblAutomationState[StateValue],"NOT_EVALUATED")

Expected:

`PASS`

ExceptionCount:

    =LET(s,XLOOKUP("PivotRefreshStatus",tblAutomationState[StateKey],tblAutomationState[StateValue],"NOT_EVALUATED"),IF(OR(s="",s="NOT_EVALUATED"),"",--(s<>"PASS")))

Status:

    =IF([@Severity]="N/A","N/A",IF([@ExceptionCount]="","",IF([@ExceptionCount]=0,"PASS",IF([@Severity]="Warning","WARNING","FAIL"))))

# Phase 10 — Operational Replenishment Report Formulas

Status:

[IMPLEMENTED] [VALIDATED]

Worksheet:

`40_RPT_Replenishment`

## Reporting Date

    =cfg_ReportingDate

Purpose:

Expose the configured business Reporting Date without duplicating business logic.

## Inventory Snapshot Date

    =INDEX(tblReplenishment[InventorySnapshotDate],1)

Purpose:

Expose the validated common Inventory Snapshot Date used by the operational Product × Site model.

## Last Successful Refresh

    =LET(last,XLOOKUP("LastSuccessfulRefresh",tblAutomationState[StateKey],tblAutomationState[StateValue],""),IF(OR(last="",last=0),"Not available",last))

Purpose:

Expose persistent accepted production-refresh state maintained by Phase 9 automation.

## Overall Quality Status

    =LET(StatusRange,tblQualityControl[Status],Applicable,ROWS(tblQualityControl[ControlID])-COUNTIF(tblQualityControl[Severity],"N/A"),Evaluated,COUNTIF(StatusRange,"PASS")+COUNTIF(StatusRange,"WARNING")+COUNTIF(StatusRange,"FAIL"),IF(COUNTIF(StatusRange,"FAIL")>0,"FAIL",IF(Evaluated<Applicable,"WARNING",IF(COUNTIF(StatusRange,"WARNING")>0,"WARNING","PASS"))))

Purpose:

Expose the established Quality Control precedence:

`FAIL > WARNING > PASS`

without creating a second Quality Control engine.

## Data Validation Helper Lists

Site:

    =VSTACK("ALL",SORT(UNIQUE(tblReplenishment[SiteID])))

Part Family:

    =VSTACK("ALL",SORT(UNIQUE(tblReplenishment[PartFamily])))

Criticality:

    =VSTACK("ALL",SORT(UNIQUE(tblReplenishment[CriticalityClass])))

Supplier Risk:

    =VSTACK("ALL",SORT(UNIQUE(tblReplenishment[SupplierRiskClass])))

Inventory Status:

    =VSTACK("ACTIONABLE","ALL","STOCKOUT","CRITICAL","REORDER","ATTENTION","EXCESS","HEALTHY")

The helper Dynamic Arrays feed Data Validation through spilled-range references.

## Operational Dynamic Array

Implemented from:

`A11`

Formula:

    =LET(Headers,{"ProductID","SiteID","PartFamily","CriticalityClass","PrimarySupplierID","SupplierRiskClass","InventoryStatus","AvailableStock","BackorderQty","OpenPOQty","InventoryPosition","SafetyStock","ReorderPoint","TargetStock","RecommendedOrderQty","NoRecentDemand"},Data,CHOOSECOLS(tblReplenishment,XMATCH(Headers,tblReplenishment[#Headers])),AllRows,SEQUENCE(ROWS(tblReplenishment[ProductID]))>0,SiteMask,IF($B$8="ALL",AllRows,tblReplenishment[SiteID]=$B$8),PartMask,IF($E$8="ALL",AllRows,tblReplenishment[PartFamily]=$E$8),CriticalityMask,IF($H$8="ALL",AllRows,tblReplenishment[CriticalityClass]=$H$8),SupplierRiskMask,IF($K$8="ALL",AllRows,tblReplenishment[SupplierRiskClass]=$K$8),StatusMask,IF($N$8="ALL",AllRows,IF($N$8="ACTIONABLE",tblReplenishment[InventoryStatus]<>"HEALTHY",tblReplenishment[InventoryStatus]=$N$8)),Keep,SiteMask*PartMask*CriticalityMask*SupplierRiskMask*StatusMask,MatchCount,SUM(--Keep),IF(MatchCount=0,"No matching records",LET(Filtered,FILTER(Data,Keep),StatusRank,XMATCH(CHOOSECOLS(Filtered,7),{"STOCKOUT","CRITICAL","REORDER","ATTENTION","EXCESS","HEALTHY"}),CriticalityRank,XMATCH(CHOOSECOLS(Filtered,4),{"A","B","C"}),SORTBY(Filtered,StatusRank,1,CriticalityRank,1,CHOOSECOLS(Filtered,15),-1,CHOOSECOLS(Filtered,9),-1,CHOOSECOLS(Filtered,2),1,CHOOSECOLS(Filtered,1),1))))

Purpose:

Produce a dynamic operational report from `tblReplenishment` while preserving the validated upstream business logic.

The formula:

1. selects the approved reporting columns by header name;
2. builds full-row masks for ALL selections;
3. applies Site, Part Family, Criticality, Supplier Risk and Inventory Status filters;
4. treats ACTIONABLE as all statuses except HEALTHY;
5. handles zero-match conditions explicitly;
6. filters the source rows;
7. converts Inventory Status to the approved operational priority;
8. converts Criticality to A/B/C priority;
9. sorts by status, Criticality, Recommended Order Quantity, Backorder Quantity, Site and Product.

The `AllRows` array is required so an ALL selection preserves the same row dimensionality as the source table.

Validated counts:

- ALL: 1,800;
- ACTIONABLE: 1,468;
- STOCKOUT: 2;
- HEALTHY: 332.

# Phase 10 — Management Dashboard Formulas

Status:

[IMPLEMENTED] [VALIDATED]

Worksheet:

`41_DASH_Management`

## STOCKOUT Positions

    =COUNTIF(tblReplenishment[InventoryStatus],"STOCKOUT")

Validated baseline:

2

## CRITICAL Positions

    =COUNTIF(tblReplenishment[InventoryStatus],"CRITICAL")

Validated baseline:

86

## REORDER Positions

    =COUNTIF(tblReplenishment[InventoryStatus],"REORDER")

Validated baseline:

310

## Recommended Order Qty

    =SUM(tblReplenishment[RecommendedOrderQty])

Validated baseline:

2,109

## Backorder Qty

    =SUM(tblReplenishment[BackorderQty])

Validated baseline:

27

## Open PO Qty

    =SUM(tblReplenishment[OpenPOQty])

Validated baseline:

20,146

## Weighted On-Time Delivery Rate

    =LET(OnTime,SUM(tblSupplierPerformance[OnTimePOCount]),Received,SUM(tblSupplierPerformance[ReceivedPOCount]),IF(Received=0,"",OnTime/Received))

Purpose:

Calculate the global On-Time Delivery Rate using Purchase Order counts.

The KPI intentionally uses:

SUM(OnTimePOCount) / SUM(ReceivedPOCount)

rather than:

AVERAGE(tblSupplierPerformance[OnTimeDeliveryRate])

because Suppliers with different Purchase Order volumes must not receive equal weighting in the global management KPI.

Validated baseline:

44.4%

## Quality Incident Count

    =SUM(tblSupplierPerformance[QualityIncidentCount])

Purpose:

Expose Reporting-Date-safe quality incidents already calculated in the Supplier Performance operational model.

Validated baseline:

350

## Reporting Context

Reporting Date:

    =cfg_ReportingDate

Inventory Snapshot Date:

    =INDEX(tblReplenishment[InventorySnapshotDate],1)

Last Successful Refresh:

    =LET(last,XLOOKUP("LastSuccessfulRefresh",tblAutomationState[StateKey],tblAutomationState[StateValue],""),IF(OR(last="",last=0),"Not available",last))

Overall Quality Status:

    =LET(StatusRange,tblQualityControl[Status],Applicable,ROWS(tblQualityControl[ControlID])-COUNTIF(tblQualityControl[Severity],"N/A"),Evaluated,COUNTIF(StatusRange,"PASS")+COUNTIF(StatusRange,"WARNING")+COUNTIF(StatusRange,"FAIL"),IF(COUNTIF(StatusRange,"FAIL")>0,"FAIL",IF(Evaluated<Applicable,"WARNING",IF(COUNTIF(StatusRange,"WARNING")>0,"WARNING","PASS"))))