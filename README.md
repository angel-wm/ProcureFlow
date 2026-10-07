# ProcureFlow

ProcureFlow is an Excel-based Procurement & Inventory Management System for inventory monitoring, replenishment decisions, procurement analysis, supplier performance, Quality Control, and management reporting.

It is built as a professional end-to-end Excel solution rather than an isolated dashboard: source data is ingested through Power Query, business rules remain auditable in Excel, Visual Basic for Applications (VBA) orchestrates the supported refresh workflow, and reporting is backed by validation controls and documented test evidence.

## Is ProcureFlow relevant to you?

ProcureFlow is designed for:

- procurement and inventory analysts who refresh data, monitor stock, identify replenishment needs, and investigate exceptions;
- buyers who need recommended order quantities, site context, open-purchase visibility, and prioritization;
- procurement or operations managers who need inventory-risk, supplier-delivery, quality, and procurement KPIs;
- reviewers who want to inspect a versioned, portfolio-ready Excel implementation with Power Query, formulas, PivotTables, VBA, testing evidence, and technical documentation.

### Important requirements and limitations

Before using the supported production workflow, note that:

- the supported platform is **Microsoft Excel 365 Desktop for Windows**;
- the production refresh workflow depends on VBA and therefore requires macro execution to be permitted by the user's environment;
- four source CSV files are required under `data/raw/` and are intentionally excluded from Git;
- the relative folder relationship between `workbook/` and `data/raw/` must be preserved;
- the validated executable workbook remains the `v1.0.0` artifact; `v1.0.2` is the current documentation-readability and repository-QA maintenance release.

For detailed setup and operating instructions, see the [ProcureFlow User Guide](docs/USER_GUIDE.md).

## Quick Start

1. Clone or download this repository.
2. Place the four required source files in `data/raw/`:
   - `parts_master.csv`
   - `supply_chain_history.csv`
   - `purchase_orders.csv`
   - `quality_incidents.csv`
3. Open [`workbook/ProcureFlow.xlsm`](workbook/ProcureFlow.xlsm) in Microsoft Excel 365 Desktop for Windows.
4. Allow VBA execution only if permitted by your environment's security policy.
5. Review the business configuration on `01_CONFIG`.
6. Return to `00_HOME` and run **Refresh ProcureFlow**.
7. Confirm the workflow state and Overall Quality Status before relying on refreshed outputs.
8. Use `40_RPT_Replenishment` for operational replenishment actions and `41_DASH_Management` for management reporting.

A successful production refresh should leave the workflow in an accepted state such as:

- `WorkflowStatus = SUCCESS`;
- `PivotRefreshStatus = PASS`;
- Overall Quality Status = `PASS`;
- `LastSuccessfulRefresh` advanced to the accepted refresh timestamp.

If refresh fails, do not treat changed worksheet values as accepted output. Use the [troubleshooting guidance](docs/USER_GUIDE.md#20-troubleshooting) before retrying.

## Portfolio Preview

### Management Dashboard

![ProcureFlow Management Dashboard](screenshots/procureflow-management-dashboard.png)

The management dashboard provides a reconciled view of inventory risk, replenishment, open purchase orders, backorders, supplier delivery performance, and quality incidents.

### Replenishment Report

![ProcureFlow Replenishment Report](screenshots/procureflow-replenishment-report.png)

The operational replenishment report provides prioritized Product × Site actions with configurable filtering and validated replenishment outputs.

### Home and Navigation

![ProcureFlow Home](screenshots/procureflow-home.png)

`00_HOME` provides the primary navigation and operating context, including Reporting Date, Last Successful Refresh, and Overall Quality Status.

### Quality Control

![ProcureFlow Quality Control](screenshots/procureflow-quality-control.png)

The centralized Quality Control layer exposes structural, reconciliation, integrity, and business-rule controls with explicit PASS / WARNING / FAIL states.

Additional technical evidence, including the Power Query dependency view, is available in [`screenshots/`](screenshots/).

## What ProcureFlow Includes

The validated workbook contains:

- reproducible Power Query ingestion and preparation;
- structured Product, Site, Supplier, Inventory, Purchase Order, Quality, and Date data;
- a 1,800-row Product × Site replenishment engine;
- a 40-row Supplier Performance model;
- Safety Stock, Reorder Point, Inventory Position, Target Stock, and Recommended Order Quantity logic;
- centralized Quality Control and reconciliation;
- Inventory, Procurement, and Supplier analytical PivotTables;
- PivotCharts, Slicers, and Timelines;
- VBA-driven production refresh orchestration;
- an operational replenishment report;
- a management dashboard;
- versioned formula, Power Query, VBA, architecture, decision, and testing documentation.

### Implemented data layer

| Excel Table | Purpose | Validated Rows |
| --- | --- | ---: |
| `tblProducts` | Product dimension | 300 |
| `tblSites` | Site dimension | 6 |
| `tblSuppliers` | Supplier dimension | 40 |
| `tblInventoryHistory` | Weekly Product × Site inventory history | 280,800 |
| `tblPurchaseOrders` | Purchase Order fact data | 29,666 |
| `tblQualityIncidents` | Supplier-quality incidents | 368 |
| `tblDate` | Daily Date dimension | 1,198 |
| `tblReplenishment` | Product × Site replenishment calculation engine | 1,800 |
| `tblSupplierPerformance` | Supplier-performance calculation model | 40 |

## Architecture at a Glance

The supported data and operating flow is:

```text
source CSV files
    ↓
Power Query ingestion and preparation
    ↓
structured Excel Tables
    ↓
operational formulas and business rules
    ↓
Quality Control
    ↓
PivotTables and analytical outputs
    ↓
VBA-orchestrated production refresh
    ↓
replenishment report and management dashboard
```

The architecture intentionally keeps responsibilities separated: Power Query prepares data, Excel formulas calculate business decisions, PivotTables aggregate, VBA orchestrates, and reporting presents validated outputs.

For the full design and implementation evidence, see [Architecture](docs/ARCHITECTURE.md).

## Repository Map

```text
ProcureFlow/
├── workbook/       executable macro-enabled workbook
├── data/           source-data policy and local raw-data location
├── power-query/    versioned Power Query M source
├── vba/            versioned VBA source and educational examples
├── docs/           canonical project documentation
└── screenshots/    portfolio and technical visual evidence
```

The four source CSV files belong under `data/raw/`, but raw data is intentionally excluded from Git.

## Documentation Map

Use the document that matches what you are trying to do:

| Need | Document |
| --- | --- |
| Set up, refresh, operate, or troubleshoot ProcureFlow | [User Guide](docs/USER_GUIDE.md) |
| Check the authoritative current project state | [Current State](docs/CURRENT_STATE.md) |
| Understand requirements, scope, and Definition of Done | [Project Specification](docs/PROJECT_SPEC.md) |
| Understand workbook and technical architecture | [Architecture](docs/ARCHITECTURE.md) |
| Review phases, target versions, and historical progression | [Roadmap](docs/ROADMAP.md) |
| Understand material design and governance decisions | [Decision Log](docs/DECISIONS.md) |
| Review source fields, grains, keys, and logical data model | [Data Dictionary](docs/DATA_DICTIONARY.md) |
| Inspect implemented Excel business formulas | [Formula Catalog](docs/FORMULAS.md) |
| Review validation, regression, acceptance, and audit evidence | [Testing](docs/TESTING.md) |
| Review VBA learning material from Phase 8 | [VBA Foundations](docs/VBA_FOUNDATIONS.md) |
| Review phase-by-phase closeout evidence | [Phase Closeouts](docs/phases/) |
| Review versioned Power Query source conventions | [Power Query Source](power-query/README.md) |
| Review production and educational VBA source conventions | [VBA Source](vba/README.md) |
| Understand source-data handling | [Data README](data/README.md) |

The canonical documents, rather than this README alone, remain authoritative for detailed project state, requirements, decisions, implementation evidence, and validation.

## Current Project Status

| Item | Current State |
| --- | --- |
| Project status | `COMPLETED` |
| Final development phase | Phase 12 — Documentation & Portfolio Release |
| Phase 12 completion version | `v1.0.0` |
| Current maintenance version | `v1.0.2` |
| Definition of Done | `32 / 32 ACCEPTED` |
| Final Phase 12 regression | `9 / 9 PASS` |
| Quality Control at final acceptance | `35 PASS / 0 WARNING / 0 FAIL` |
| Critical known defects at final acceptance | `0` |

The `v1.0.2` maintenance release consolidates the R1–R4 documentation-readability work and final repository QA. It does not change the executable workbook, which remains the validated `v1.0.0` artifact.

For authoritative release-state details, see [Current State](docs/CURRENT_STATE.md) and [Testing](docs/TESTING.md).

## Technology Stack

ProcureFlow is intentionally centered on Excel:

- Microsoft Excel 365 Desktop for Windows;
- Excel Tables and Structured References;
- Power Query;
- Excel formulas and Dynamic Arrays;
- Named Ranges and Named Formulas;
- Data Validation and Conditional Formatting;
- PivotTables and PivotCharts;
- Slicers and Timelines;
- Macros and Visual Basic for Applications (VBA);
- Git, GitHub, and Markdown technical documentation.

Power Pivot / Excel Data Model is not required by the current implementation.

## Versioned Source

The executable implementation lives in [`workbook/ProcureFlow.xlsm`](workbook/ProcureFlow.xlsm), while text representations make otherwise-binary implementation details reviewable in Git:

- important Excel formulas are maintained in the [Formula Catalog](docs/FORMULAS.md);
- Power Query M source is maintained under [`power-query/`](power-query/);
- production VBA source is maintained under [`vba/modules/`](vba/modules/);
- educational Phase 8 VBA is kept separately under [`vba/examples/phase08/`](vba/examples/phase08/).

The workbook remains the executable source for Power Query, VBA, formulas, PivotTables, and reporting behavior. Versioned text source must stay synchronized whenever the corresponding workbook implementation changes.

## Development and Release History

ProcureFlow was developed sequentially. A phase was considered complete only after implementation, validation, documentation, Git workflow, merge, and publication criteria were satisfied.

| Phase | Scope | Completion Version |
| ---: | --- | --- |
| 0 | Project Design | `v0.1.0` |
| 1 | Data Design | `v0.2.0` |
| 2 | Workbook Foundation | `v0.3.0` |
| 3 | Power Query Pipeline | `v0.4.0` |
| 4 | Operational Model | `v0.5.0` |
| 5 | Business Logic & Advanced Formulas | `v0.6.0` |
| 6 | Quality Control System | `v0.7.0` |
| 7 | Analysis & PivotTables | `v0.8.0` |
| 8 | VBA Foundations | `v0.8.1` |
| 9 | Automation | `v0.9.0` |
| 10 | Reporting & Dashboard | `v0.10.0` |
| 11 | Testing & Hardening | `v0.11.0` |
| 12 | Documentation & Portfolio Release | `v1.0.0` |

Corrective and maintenance releases do not change the formal completion version of the phase they follow. Historical patch details are maintained in [Current State](docs/CURRENT_STATE.md), while full phase history and closeout evidence are available in the [Roadmap](docs/ROADMAP.md) and [phase closeouts](docs/phases/).

## License

ProcureFlow source code and original project documentation are licensed under the [MIT License](LICENSE).

External datasets retain their own applicable licenses and terms.
