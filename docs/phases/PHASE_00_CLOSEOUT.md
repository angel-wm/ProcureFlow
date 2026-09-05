# ProcureFlow — Phase 0 Closeout

## Phase

Phase 0 — Project Design

## Status

IN PROGRESS — GITHUB GATE PENDING

All local Phase 0 design and documentation work has been completed.

The phase will only become COMPLETED after the approved Phase 0 state is merged into `main`, published to GitHub and tagged.

## Objective

Design ProcureFlow completely before implementation begins.

## Confirmed Deliverables

- Project scope
- Business problem
- Users and use cases
- Official dataset selection
- Functional requirements
- Non-functional requirements
- Conceptual data model
- Business rules
- Replenishment policy
- Workbook architecture
- Power Query strategy
- Formula strategy
- PivotTable strategy
- VBA strategy
- Dashboard/reporting strategy
- Quality-control strategy
- Git/GitHub workflow
- Repository structure
- Roadmap
- Acceptance criteria
- Definition of Done

## Official Dataset

Aerospace Supply Chain Performance & Forecasting

Official source files:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

## Key Confirmed Design Decisions

- Excel 365 Desktop for Windows is the target platform.
- Raw source CSV files remain unchanged.
- Power Query handles ingestion and preparation.
- Excel formulas implement the core operational business rules.
- PivotTables provide analytical aggregation.
- VBA provides automation and orchestration rather than replacing the core model.
- Power Pivot is deferred unless later justified.
- Replenishment is evaluated at Product × Site level.
- The default demand-history window is 26 weeks.
- The review period is 1 week.
- Service levels are initially 99% / 97% / 95% for criticality classes A / B / C.
- Safety Stock considers demand variability and Lead Time variability.
- Reorder Point equals expected Lead-Time demand plus Safety Stock.
- Inventory Position includes available stock, open purchase orders and backorders.
- GitHub is the official handoff mechanism between phases.
- Every completed phase must be documented and pushed before the next phase starts.

## Implementation Status

No workbook implementation has started.

No formulas, Power Query queries, PivotTables, macros, VBA or dashboard components are considered implemented.

## Local Completion Evidence

Completed locally:

- repository initialized;
- repository structure created;
- raw-data exclusion configured;
- official source CSV files placed locally in `data/raw/`;
- canonical documentation created;
- Phase 0 design baseline committed;
- dedicated Phase 0 branch created.

Current branch:

`phase/00-project-design`

## Remaining Phase 0 Tasks

- Create GitHub repository
- Publish repository
- Open Phase 0 Pull Request
- Review and merge into `main`
- Push final `main`
- Create and push tag `phase-0-complete`
- Create and push tag `v0.1.0`
- Update this document to COMPLETED
- Update `CURRENT_STATE.md` for transition to Phase 1
- Update `ROADMAP.md` Phase 0 status to COMPLETED

## Next Phase

Phase 1 — Data Design

