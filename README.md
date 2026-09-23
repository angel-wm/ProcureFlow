# ProcureFlow

ProcureFlow is an Excel-based Procurement & Inventory Management System designed as a professional end-to-end workbook for procurement, inventory, replenishment and supplier-performance workflows.

The solution is being developed incrementally in Microsoft Excel with reproducible data ingestion, structured operational models, auditable formulas, validation controls and Git/GitHub documentation.

## Current Project Status

Project status:

`IN DEVELOPMENT`

Current released version:

`v0.9.2`

Current development phase:

**Phase 10 — Reporting & Dashboard**

Phase 10 status:

**NOT STARTED**

Last completed phase:

**Phase 9 — Automation**

Pull Request:

`#10 — Phase 9 — Automation`

Version:

`v0.9.0`

## Completed Phases

### Phase 0 — Project Design

Status:

`COMPLETED`

Version:

`v0.1.0`

Established the initial project specification, architecture, roadmap, decision framework and Git/GitHub governance.

### Phase 1 — Data Design

Status:

`COMPLETED`

Version:

`v0.2.0`

Established and validated the logical data model, field definitions, grains, source relationships and data-quality baseline.

### Phase 2 — Workbook Foundation

Status:

`COMPLETED`

Version:

`v0.3.0`

Created the physical macro-enabled Excel workbook, configuration layer, navigation, worksheet architecture, visual conventions, Named Ranges and initial controls.

### Phase 3 — Power Query Pipeline

Status:

`COMPLETED`

Technical release:

`v0.4.0`

Phase 3 corrective documentation release:

`v0.4.1`

Implemented the reproducible Power Query pipeline from raw CSV files through source, staging, dimension and fact queries into structured Excel Tables.

### Phase 4 — Operational Model

Status:

`COMPLETED`

Version:

`v0.5.0`

Completed implementation includes:

- Product × Site operational model;
- `tblReplenishment`;
- 1,800 validated Product × Site rows;
- Supplier operational model;
- `tblSupplierPerformance`;
- 40 validated Supplier rows;
- Product and Supplier master attributes;
- Reporting Date relationships;
- weekly Inventory Snapshot resolution;
- On-Hand Quantity;
- Blocked Quantity;
- Backorder Quantity;
- completed historical-demand window preparation;
- Lead-Time inputs;
- Purchase Order inputs prepared for later business logic;
- versioned Excel formula documentation;
- full post-implementation Refresh validation.

Phase 4 intentionally stopped before Phase 5 business calculations.

### Phase 5 — Business Logic & Advanced Formulas

Status:

`COMPLETED`

Version:

`v0.6.0`

Corrective documentation releases after Phase 5:

- `v0.6.1` — synchronized Phase 5 release metadata; commit `0f447a3599556a0444103d43cdfc49ff1eef48cb`.
- `v0.6.2` — finalized Phase 5 release synchronization; commit `1ae9dfecdba3badd7bf071a35341994b66594602`.

These corrective versions did not change the formal Phase 5 completion version, which remains `v0.6.0`.

Completed implementation includes:

- Historical Demand;
- Average Weekly Demand;
- demand variability;
- historical Actual Lead Time;
- Lead-Time variability;
- Effective Lead-Time fallback;
- Open PO Quantity;
- Available Stock;
- configurable Service Levels;
- Safety Stock;
- Reorder Point;
- Inventory Position;
- Target Stock;
- Recommended Order Quantity;
- Inventory Status;
- NoRecentDemand;
- On-Time Delivery metrics;
- Late Delivery metrics;
- partial-receipt metrics;
- supplier Lead-Time metrics;
- supplier Quality Incident counts;
- documented and validated Excel business formulas.

Phase 5 validation completed with zero formula errors. The Product × Site replenishment engine contains 1,800 validated rows and Supplier Performance contains 40 validated Suppliers.

### Phase 6 — Quality Control System

Status:

`COMPLETED`

Version:

`v0.7.0`

Implemented the centralized Quality Control framework, reconciliation controls, Power Query technical-health feed and validated PASS / WARNING / FAIL behavior.

### Phase 7 — Analysis & PivotTables

Status:

`COMPLETED`

Version:

`v0.8.0`

Implemented the Inventory, Procurement and Supplier analytical layers using PivotTables, PivotCharts, Slicers and Timelines, with source-to-Pivot reconciliation and full refresh validation.

### Phase 8 — VBA Foundations

Status:

`COMPLETED`

Version:

`v0.8.1`

Established practical VBA foundations using the real ProcureFlow workbook, validated readiness for production automation and preserved educational VBA separately from the operational workbook.

### Phase 9 — Automation

Status:

`COMPLETED`

Version:

`v0.9.0`

Implemented the production VBA automation layer, including controlled synchronous Power Query refresh, persistent automation state, staged Quality Control orchestration, PivotTable refresh coordination, user-facing refresh execution, controlled failure handling and late-stage successful-refresh rollback protection.

## Post-Phase-9 Corrective Releases

### v0.9.1

Documentation-only synchronization after the Phase 9 release. Published through Pull Request `#11` with no changes to the executable workbook, VBA, Power Query or business logic.

### v0.9.2

Documentation-only cleanup before Phase 10, covering duplicated Phase 9 test evidence, obsolete architecture-status wording, corrective-release traceability, CURRENT_STATE normalization and establishment of v0.9.2 as the Phase 10 baseline.

## Current Workbook

Executable workbook:

`workbook/ProcureFlow.xlsm`

Target application:

Microsoft Excel 365 Desktop for Windows.

The workbook currently contains the approved layered architecture for:

- configuration and control;
- structured data;
- operational calculations;
- PivotTable analysis;
- reporting;
- management dashboard development.

Later-phase worksheets may exist structurally without implying that their future functionality has already been implemented.

## Technology Stack

ProcureFlow is intentionally centered on Excel.

Current and planned technologies include:

- Microsoft Excel 365 Desktop for Windows
- Excel Tables
- Structured References
- Power Query
- Excel formulas
- Dynamic Arrays
- Named Ranges and Named Formulas
- Data Validation
- Conditional Formatting
- PivotTables and PivotCharts
- Slicers and Timelines
- Macros
- Visual Basic for Applications (VBA)
- Git
- GitHub
- Markdown technical documentation

Power Pivot / Excel Data Model is not currently required.

## Implemented Data Layer

The current Power Query pipeline loads the following structured Excel Tables:

| Excel Table | Purpose | Validated Rows |
|---|---|---:|
| `tblProducts` | Product dimension | 300 |
| `tblSites` | Site dimension | 6 |
| `tblSuppliers` | Supplier dimension | 40 |
| `tblInventoryHistory` | Weekly Product × Site inventory history | 280,800 |
| `tblPurchaseOrders` | Purchase Order fact data | 29,666 |
| `tblQualityIncidents` | Supplier-quality incidents | 368 |
| `tblDate` | Daily Date dimension | 1,198 |
| `tblReplenishment` | Product × Site replenishment calculation engine | 1,800 |
| `tblSupplierPerformance` | Supplier-performance calculation model | 40 |

## Documentation

Canonical project documentation is maintained under:

`docs/`

Important documents include:

- `docs/PROJECT_SPEC.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DECISIONS.md`
- `docs/CURRENT_STATE.md`
- `docs/DATA_DICTIONARY.md`
- `docs/TESTING.md`
- `docs/FORMULAS.md`

Phase closeout documents are maintained under:

`docs/phases/`

The canonical documents, rather than this README alone, are the authoritative source for detailed project state and implementation decisions.

## Formula Versioning

Important Excel formulas implemented in the workbook are also maintained as text in:

`docs/FORMULAS.md`

This provides inspectable Git history for formulas that otherwise live inside the binary `.xlsm` workbook.

## Power Query Source Versioning

Executable Power Query M code lives inside:

`workbook/ProcureFlow.xlsm`

A synchronized text representation is maintained under:

`power-query/`

This allows Power Query implementation to be reviewed through normal Git and GitHub diffs.

## Development Approach

ProcureFlow is developed phase by phase.

Major functionality is not marked implemented until it has been:

1. designed;
2. created in the workbook;
3. validated with actual Excel evidence;
4. documented;
5. versioned in Git;
6. reviewed through the corresponding phase workflow.

Each completed phase uses a dedicated branch, Pull Request, merge, phase-completion tag and semantic version tag.

## Next Planned Phase

**Phase 10 — Reporting & Dashboard**

Target version:

`v0.10.0`

## Portfolio Status

ProcureFlow is still under active development.

The final public-facing portfolio presentation, screenshots, polished repository documentation and complete README are planned for the later documentation and release phase.

## License

ProcureFlow source code and original project documentation are licensed under the MIT License.

External datasets retain their own applicable licenses and terms.




