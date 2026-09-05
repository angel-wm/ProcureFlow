# ProcureFlow — Data Dictionary

## Status

PENDING DETAILED COMPLETION IN PHASE 1 — DATA DESIGN

## Official Source Files

### parts_master.csv

Expected grain:

1 row = 1 product / part.

Primary role:

Product master and product-level planning attributes.

### supply_chain_history.csv

Expected grain:

1 row = 1 Product × 1 Site × 1 Week.

Primary role:

Historical inventory, consumption, backorders, blocked stock and forecast information.

### purchase_orders.csv

Expected grain:

1 row = 1 purchase-order record for one product.

Primary role:

Procurement, quantities, supplier delivery and lead-time analysis.

Important limitation:

The source does not provide a traditional Purchase Order Header → multiple Purchase Order Lines structure.

### quality_incidents.csv

Expected grain:

1 row = 1 quality incident.

Primary role:

Supplier-quality analysis.

## Planned ProcureFlow Tables

- `tblProducts`
- `tblSites`
- `tblSuppliers`
- `tblInventoryHistory`
- `tblPurchaseOrders`
- `tblQualityIncidents`
- `tblReplenishment`
- `tblSupplierPerformance`

A full field-level dictionary will be produced and validated during Phase 1.
