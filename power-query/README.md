# ProcureFlow — Power Query Source

This directory contains the text-source representation of the Power Query M code implemented in `workbook/ProcureFlow.xlsm`.

## Purpose

The Excel workbook is a binary `.xlsm` file, so Git cannot provide meaningful line-by-line diffs for embedded Power Query queries.

The `.pq` files in this directory provide:

- inspectable Power Query M source;
- meaningful Git diffs;
- technical documentation;
- implementation traceability;
- easier code review.

## Query Layers

### `src/`

Source-access queries.

Responsibilities:

- resolve the configured raw-data folder;
- open official CSV files;
- parse CSV content;
- promote source headers.

No business cleansing or business-rule logic belongs here.

### `stg/`

Staging and preparation queries.

Responsibilities:

- rename source fields to ProcureFlow logical names;
- apply explicit data types;
- standardize objective representations;
- surface invalid technical values.

### `dim/`

Final dimension-query definitions.

Implemented dimensions:

- `dim_Product`
- `dim_Site`
- `dim_Supplier`
- `dim_Date`

### `fact/`

Final fact-query definitions.

Implemented facts:

- `fact_InventoryWeekly`
- `fact_PurchaseOrders`
- `fact_QualityIncidents`

## Source of Truth

The production implementation runs inside `workbook/ProcureFlow.xlsm`.

The `.pq` files must be kept synchronized with the corresponding workbook queries whenever Power Query code changes.

## Language

Power Query queries are written in the Power Query M language.

M is separate from Visual Basic for Applications (VBA).

Power Query M is primarily responsible for ingestion and reproducible data preparation.

VBA will be introduced in a later ProcureFlow phase for workflow orchestration and automation.
