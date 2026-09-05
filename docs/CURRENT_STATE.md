# ProcureFlow — Current State

## Project Status

IN DEVELOPMENT

## Current Version

v0.1.0

## Current Phase

Phase 1 — Data Design

Status: IN PROGRESS

Target Version: `v0.2.0`

Phase Branch:

`phase/01-data-design`

## Last Completed Phase

Phase 0 — Project Design

Version: `v0.1.0`

GitHub Pull Request: `#1 — Phase 0 — Project Design`

Merge Commit: `4a30476`

Final Phase 0 Commit: `3cbab55`

Phase Tag: `phase-0-complete`

Version Tag: `v0.1.0`

## Implemented

Project governance and design baseline only.

Implemented project assets:

- local and GitHub repository structure;
- canonical Markdown documentation;
- dataset policy;
- Git/GitHub workflow;
- roadmap;
- architecture baseline;
- requirements baseline;
- business-rule baseline;
- acceptance criteria;
- Definition of Done.

No Excel workbook functionality has been implemented yet.

No Power Query queries, formulas, PivotTables, VBA procedures, automation, operational reports or dashboards are considered implemented.

## Phase 1 Progress

Source inspection and data-design validation are currently in progress.

Validated locally so far:

- official source-file availability;
- source row counts;
- source column schemas;
- primary-key uniqueness for Products, Purchase Orders and Quality Incidents;
- composite Inventory History grain;
- Product, Site and Supplier populations;
- Product, Site and Supplier referential integrity;
- mandatory identifier completeness;
- source null behavior;
- observed categorical domains;
- observed numeric ranges;
- date ranges;
- weekly historical continuity;
- Product × Site historical completeness;
- Purchase Order chronological consistency;
- partial-receipt behavior;
- Lead-Time ranges;
- Supplier-to-Product cardinality;
- Supplier Risk dependency;
- forecast type / uplift consistency;
- source identifier formats;
- candidate logical data types.

These validations are data-design evidence only.

No workbook, Power Query pipeline or operational functionality is considered implemented.

## Official Dataset

Aerospace Supply Chain Performance & Forecasting

Official local source files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

The files are stored locally in:

`data/raw/`

and are excluded from Git.

## Confirmed Source Counts

- Products: 300
- Inventory History: 280,800
- Purchase Orders: 29,666
- Quality Incidents: 368
- Sites: 6
- Suppliers: 40

## Confirmed Historical Grain

Inventory History:

1 Product × 1 Site × 1 Week

Validated population:

- 1,800 Product × Site combinations;
- 156 weeks per Product × Site;
- 156 distinct weekly dates;
- all weekly dates are Mondays;
- historical range: 2022-01-03 through 2024-12-23.

## Known Issues

None currently classified as critical.

`shelf_life_days` is nullable in 274 of 300 Product records and is being treated as an optional source attribute.

## Relevant Decisions

Confirmed Phase 0 decisions:

`DEC-001` through `DEC-041`

See:

`docs/DECISIONS.md`

## Next Immediate Step

Complete the formal Phase 1 data specification.

Immediate work:

- complete `docs/DATA_DICTIONARY.md`;
- document source-to-target mappings;
- finalize dimension and fact specifications;
- finalize Date dimension design;
- document relationships and cardinalities;
- record Phase 1 test evidence.

## Next Phase

Phase 2 — Workbook Foundation

Status: NOT STARTED

Target Version: `v0.3.0`

Phase 2 must not begin until Phase 1 satisfies its exit criteria and GitHub gate.
