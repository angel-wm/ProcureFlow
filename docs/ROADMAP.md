# ProcureFlow — Roadmap

## Document Status

Status: CONFIRMED  
Project State: PRE-PROJECT  
Current Phase: Phase 0 — Project Design  
Current Phase Status: IN PROGRESS  
Target Phase 0 Version: v0.1.0

This roadmap defines the approved sequential development plan for ProcureFlow.

A phase is only considered COMPLETED after its technical work, validation, documentation and GitHub publication gate have all been completed.

---

# 1. Roadmap Governance

## 1.1 Phase Status Values

Only the following phase statuses will be used:

- NOT STARTED
- IN PROGRESS
- BLOCKED
- COMPLETED

## 1.2 Sequential Execution

ProcureFlow will initially be developed sequentially.

A new phase must not begin until the preceding phase has completed its exit criteria and GitHub gate.

## 1.3 GitHub Completion Rule

Local completion alone does not mark a phase as COMPLETED.

A completed phase must:

1. satisfy its deliverables;
2. satisfy its validation requirements;
3. update applicable canonical documentation;
4. create its phase closeout document;
5. complete its Git workflow;
6. be merged into `main`;
7. be pushed to GitHub;
8. have its phase-complete tag published.

## 1.4 Canonical Phase Handoff

A new phase chat should reconstruct project state from GitHub, primarily through:

1. `docs/CURRENT_STATE.md`
2. previous `docs/phases/PHASE_XX_CLOSEOUT.md`
3. `docs/ROADMAP.md`
4. `docs/ARCHITECTURE.md`
5. `docs/DECISIONS.md`
6. other relevant documentation as required.

Learning or debugging chats do not automatically modify official project state.

---

# 2. Version Strategy

ProcureFlow uses Semantic Versioning concepts:

`MAJOR.MINOR.PATCH`

Planned phase versions:

| Phase | Target Version |
|---|---|
| Phase 0 | v0.1.0 |
| Phase 1 | v0.2.0 |
| Phase 2 | v0.3.0 |
| Phase 3 | v0.4.0 |
| Phase 4 | v0.5.0 |
| Phase 5 | v0.6.0 |
| Phase 6 | v0.7.0 |
| Phase 7 | v0.8.0 |
| Phase 8 | v0.8.1 |
| Phase 9 | v0.9.0 |
| Phase 10 | v0.10.0 |
| Phase 11 | v0.11.0 |
| Phase 12 | v1.0.0 |

Patch versions may be created where corrective work is required.

`v1.0.0` is reserved for the fully accepted portfolio-ready release.

---

# 3. Phase 0 — Project Design

## Status

IN PROGRESS

## Target Version

v0.1.0

## Local Work Status

COMPLETE — GITHUB GATE PENDING

## Depends On

None.

## Objective

Design ProcureFlow formally before implementation begins.

## Scope

- define business purpose;
- establish project scope;
- identify users and use cases;
- select the official dataset;
- inspect dataset suitability;
- define conceptual model;
- define functional requirements;
- define non-functional requirements;
- establish business rules;
- establish replenishment policy;
- establish supplier-performance policy;
- design workbook architecture;
- define Power Query responsibilities;
- define Excel-formula responsibilities;
- define PivotTable strategy;
- define VBA strategy;
- define quality-control strategy;
- define Git/GitHub strategy;
- define repository structure;
- define roadmap;
- define acceptance criteria;
- define Definition of Done;
- create initial canonical documentation;
- create and publish the initial repository state.

## Deliverables

- `README.md`
- `.gitignore`
- repository folder structure
- `docs/PROJECT_SPEC.md`
- `docs/ARCHITECTURE.md`
- `docs/ROADMAP.md`
- `docs/DECISIONS.md`
- `docs/CURRENT_STATE.md`
- initial `docs/DATA_DICTIONARY.md`
- initial `docs/TESTING.md`
- `docs/phases/PHASE_00_CLOSEOUT.md`

## Entry Criteria

None.

## Exit Criteria

- design baseline approved;
- official dataset confirmed;
- requirements confirmed;
- business rules confirmed;
- architecture confirmed;
- quality-control strategy confirmed;
- Git strategy confirmed;
- roadmap confirmed;
- acceptance criteria confirmed;
- Definition of Done confirmed;
- canonical documents created;
- repository structure validated;
- Phase 0 closeout completed;
- no unresolved critical Phase 0 design issue.

## GitHub Gate

- local Git repository initialized;
- canonical files committed;
- GitHub repository created;
- `main` pushed;
- Phase 0 state visible in GitHub;
- tag `phase-0-complete` created and pushed;
- tag `v0.1.0` created and pushed.

## Version

v0.1.0

---

# 4. Phase 1 — Data Design

## Status

NOT STARTED

## Target Version

v0.2.0

## Depends On

Phase 0 — Project Design.

## Objective

Transform the selected Aerospace dataset into a formally validated data specification before workbook construction.

## Scope

- inspect every source file;
- validate row grain;
- inspect every relevant source column;
- classify data types;
- identify primary keys;
- identify candidate keys;
- validate uniqueness;
- identify foreign-key relationships;
- identify required and optional fields;
- inspect null behavior;
- inspect relevant value ranges;
- identify date ranges;
- identify source anomalies;
- define Product dimension;
- define Site dimension;
- define Supplier dimension;
- define Date dimension;
- define Inventory fact;
- define Purchase Order fact;
- define Quality Incident fact;
- define source-to-ProcureFlow field mappings;
- identify data-quality rules supported by source semantics.

## Deliverables

- completed field-level `DATA_DICTIONARY.md`;
- validated grain definitions;
- key and relationship specification;
- source-to-target mapping;
- dimension and fact specification;
- Date dimension design;
- data-model diagram or equivalent documentation;
- Phase 1 test evidence;
- `PHASE_01_CLOSEOUT.md`.

## Learning Focus

- table grain;
- primary keys;
- foreign keys;
- dimensions;
- facts;
- relational thinking;
- data types;
- data dictionaries;
- source validation.

## Entry Criteria

- Phase 0 COMPLETED;
- Phase 0 documentation published;
- `main` synchronized;
- `phase-0-complete` tag available.

## Exit Criteria

- every required source field understood;
- grains validated;
- required keys validated;
- relationships documented;
- required dimensions and facts specified;
- no unresolved critical data-design ambiguity;
- `DATA_DICTIONARY.md` updated;
- test evidence documented;
- Phase 1 closeout completed.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge to `main`, push and phase tag.

## Version

v0.2.0

---

# 5. Phase 2 — Workbook Foundation

## Status

NOT STARTED

## Target Version

v0.3.0

## Depends On

Phase 1 — Data Design.

## Objective

Create the first physical ProcureFlow workbook and establish a professional workbook foundation.

## Scope

- create `ProcureFlow.xlsm`;
- create approved worksheet architecture;
- establish sheet-order conventions;
- create base navigation;
- create `00_HOME`;
- create `01_CONFIG`;
- create `02_CONTROL`;
- create technical sheet structure;
- establish styles and formatting conventions;
- create configuration inputs;
- create Data Validation;
- create initial Named Ranges / Named Formulas;
- establish Excel Table placeholders where appropriate;
- establish initial protection strategy.

## Deliverables

- first `workbook/ProcureFlow.xlsm`;
- complete physical worksheet skeleton;
- working navigation foundation;
- configuration structure;
- naming standards implemented;
- initial protection and visual conventions;
- Phase 2 test evidence;
- `PHASE_02_CLOSEOUT.md`.

## Learning Focus

- workbook architecture;
- Excel Tables;
- Structured References;
- absolute, relative and mixed references;
- Named Ranges;
- Named Formulas;
- Data Validation;
- workbook protection;
- maintainable workbook design.

## Entry Criteria

- Phase 1 COMPLETED;
- data design stable;
- physical data requirements known.

## Exit Criteria

- workbook opens without errors;
- required base sheets exist;
- configuration inputs are separated from formulas;
- navigation foundation works;
- naming conventions are applied;
- workbook is ready for Power Query integration.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.3.0

---

# 6. Phase 3 — Power Query Pipeline

## Status

NOT STARTED

## Target Version

v0.4.0

## Depends On

Phase 2 — Workbook Foundation.

## Objective

Build a reproducible Power Query ingestion and preparation pipeline from raw source files to structured ProcureFlow data tables.

## Scope

- implement `src_*` queries;
- implement `stg_*` queries;
- assign data types;
- standardize required fields;
- implement objective cleansing;
- implement dimensions;
- implement fact outputs;
- implement Date dimension;
- configure query load behavior;
- use Connection Only where appropriate;
- load required final tables;
- test Refresh;
- validate row counts and transformations;
- document source-path strategy.

## Deliverables

- functioning Power Query pipeline;
- `tblProducts`;
- `tblSites`;
- `tblSuppliers`;
- `tblInventoryHistory`;
- `tblPurchaseOrders`;
- `tblQualityIncidents`;
- Date structure as designed;
- refresh documentation;
- query validation evidence;
- `PHASE_03_CLOSEOUT.md`.

## Learning Focus

- Power Query Editor;
- query steps;
- data types;
- source queries;
- staging;
- transformations;
- merges;
- appends where justified;
- query dependencies;
- Connection Only;
- Refresh.

## Entry Criteria

- Phase 2 COMPLETED;
- workbook architecture available;
- source-to-target design confirmed.

## Exit Criteria

- valid source files refresh successfully;
- expected final tables load correctly;
- row counts reconcile;
- no critical Power Query errors remain;
- no manual cleansing of raw files is required.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.4.0

---

# 7. Phase 4 — Operational Model

## Status

NOT STARTED

## Target Version

v0.5.0

## Depends On

Phase 3 — Power Query Pipeline.

## Objective

Build the structured operational calculation model required before implementing final business formulas.

## Scope

- create Product × Site operational grain;
- design `tblReplenishment`;
- design `tblSupplierPerformance`;
- expose current inventory information;
- expose blocked stock;
- expose backorders;
- expose relevant historical demand;
- expose open-Purchase-Order inputs;
- prepare supplier attributes;
- prepare Lead-Time inputs;
- establish Reporting Date relationships;
- reduce repeated scanning of large history tables where appropriate.

## Deliverables

- `20_CALC_Replenishment` structure;
- `tblReplenishment` structural model;
- `21_CALC_SupplierPerformance` structure;
- `tblSupplierPerformance` structural model;
- validated Product × Site population;
- prepared inputs for Phase 5;
- `PHASE_04_CLOSEOUT.md`.

## Learning Focus

- operational-model design;
- grain preservation;
- structured formulas;
- intermediate calculations;
- performance-aware Excel modeling.

## Entry Criteria

- Phase 3 COMPLETED;
- all required structured data tables refresh correctly.

## Exit Criteria

- operational grains are correct;
- required inputs are available;
- no unresolved critical structural issue remains;
- model is ready for business-rule formulas.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.5.0

---

# 8. Phase 5 — Business Logic & Advanced Formulas

## Status

NOT STARTED

## Target Version

v0.6.0

## Depends On

Phase 4 — Operational Model.

## Objective

Implement and validate the core inventory, replenishment, procurement and supplier-performance business logic using auditable Excel formulas.

## Scope

Implement:

- historical demand;
- average demand;
- demand variability;
- Actual Lead Time;
- average Lead Time;
- Lead-Time variability;
- effective Lead Time fallback;
- Service Level lookup;
- Safety Stock;
- Reorder Point;
- Available Stock;
- open Purchase Order quantity;
- Inventory Position;
- Target Stock;
- Recommended Order Quantity;
- Inventory Status;
- `NO_RECENT_DEMAND`;
- On-Time Delivery;
- delay metrics;
- partial-receipt metrics;
- supplier-quality metrics.

## Formula Learning Focus

Use and understand professionally relevant functions including:

- XLOOKUP;
- VLOOKUP;
- HLOOKUP;
- INDEX;
- MATCH;
- IF;
- IFS;
- AND;
- OR;
- IFERROR;
- SUMIF;
- SUMIFS;
- COUNTIF;
- COUNTIFS;
- AVERAGEIFS;
- FILTER;
- SORT;
- SORTBY;
- UNIQUE;
- SEQUENCE;
- LET;
- LAMBDA when justified;
- date functions;
- text functions.

Legacy functions will be understood but not forced into production logic when a more maintainable modern function is preferable.

## Deliverables

- functioning replenishment engine;
- functioning supplier-performance calculations;
- documented critical formulas;
- manual calculation checks;
- formula test evidence;
- `PHASE_05_CLOSEOUT.md`.

## Entry Criteria

- Phase 4 COMPLETED;
- operational model structurally valid.

## Exit Criteria

- mandatory business rules implemented;
- critical calculations manually reconciled;
- formulas use configuration parameters correctly;
- no critical negative or contradictory outputs remain;
- formula performance is acceptable on full data.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.6.0

---

# 9. Phase 6 — Quality Control System

## Status

NOT STARTED

## Target Version

v0.7.0

## Depends On

Phase 5 — Business Logic & Advanced Formulas.

## Objective

Implement the formal ProcureFlow data-quality, logic-quality and reconciliation framework.

## Scope

- implement `02_CONTROL`;
- implement structural controls;
- implement referential controls;
- implement business-rule controls;
- implement configuration controls;
- implement calculation controls;
- implement reconciliations;
- implement PASS / WARNING / FAIL statuses;
- determine overall quality state;
- document exceptions;
- establish refresh-state controls.

## Deliverables

- operational `02_CONTROL`;
- mandatory Quality Control checks;
- reconciliation framework;
- understandable exception counts;
- documented severity;
- Quality Control test evidence;
- `PHASE_06_CLOSEOUT.md`.

## Learning Focus

- error controls;
- COUNTIFS;
- SUMIFS;
- logical validation;
- reconciliation;
- exception design;
- transparent error handling.

## Entry Criteria

- Phase 5 COMPLETED;
- business calculations available for validation.

## Exit Criteria

- mandatory controls implemented;
- critical FAIL behavior works;
- reconciliations pass for valid source data;
- no critical data-quality problem is silently hidden.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.7.0

---

# 10. Phase 7 — Analysis & PivotTables

## Status

NOT STARTED

## Target Version

v0.8.0

## Depends On

Phase 6 — Quality Control System.

## Objective

Build the analytical layer using PivotTables, PivotCharts, Slicers and Timelines.

## Scope

Develop:

- `30_PVT_Inventory`;
- `31_PVT_Procurement`;
- `32_PVT_Suppliers`;
- relevant PivotTables;
- relevant PivotCharts;
- date grouping;
- interactive filtering;
- Slicers;
- Timelines where justified;
- reconciliation with source tables.

## Deliverables

- inventory analysis;
- procurement analysis;
- supplier analysis;
- interactive filters;
- reconciled PivotTables;
- Phase 7 test evidence;
- `PHASE_07_CLOSEOUT.md`.

## Learning Focus

- PivotTables;
- PivotCharts;
- grouping;
- filters;
- Slicers;
- Timelines;
- refresh behavior.

## Entry Criteria

- Phase 6 COMPLETED;
- validated structured tables and business calculations available.

## Exit Criteria

- required analytical domains exist;
- PivotTable totals reconcile;
- interactive filters work;
- refresh behavior is validated.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.8.0

---

# 11. Phase 8 — VBA Foundations

## Status

NOT STARTED

## Target Version

v0.8.1

## Depends On

Phase 7 — Analysis & PivotTables.

## Objective

Build practical understanding of Visual Basic for Applications before implementing production automation.

## Scope

Progressively learn:

1. what a macro is;
2. Macro Recorder;
3. generated VBA inspection;
4. Visual Basic Editor;
5. `Sub`;
6. variables using `Dim`;
7. common data types;
8. Workbook object;
9. Worksheet object;
10. Range object;
11. Cells property;
12. ListObject;
13. `If`;
14. `For`;
15. `For Each`;
16. `Do While`;
17. procedures and functions;
18. basic error handling;
19. PivotTable interaction;
20. refresh concepts.

## Deliverables

- educational macros tied to relevant workbook operations;
- documented VBA learning examples where appropriate;
- demonstrated ability to read and explain generated code;
- readiness assessment for production automation;
- `PHASE_08_CLOSEOUT.md`.

## Learning Focus

Visual Basic for Applications fundamentals.

## Entry Criteria

- Phase 7 COMPLETED;
- workbook manual workflows are sufficiently stable to automate.

## Exit Criteria

- core VBA concepts understood;
- user can inspect and explain foundational code;
- production automation can proceed without treating VBA as a black box.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.8.1

---

# 12. Phase 9 — Automation

## Status

NOT STARTED

## Target Version

v0.9.0

## Depends On

Phase 8 — VBA Foundations.

## Objective

Implement professional VBA automation for stable ProcureFlow workflows.

## Scope

Potential automations include:

- refresh workflow;
- dependent PivotTable refresh;
- Quality Control orchestration;
- successful-refresh timestamp;
- navigation;
- operational-report preparation;
- report export;
- appropriate user messages;
- error handling;
- stale-output prevention.

## Deliverables

- production VBA modules;
- documented procedures;
- error-handling behavior;
- exported `.bas` modules;
- automation test evidence;
- `PHASE_09_CLOSEOUT.md`.

## Learning Focus

- modular procedures;
- workbook automation;
- ListObjects;
- PivotTables;
- refresh control;
- error handling;
- maintainable VBA structure.

## Entry Criteria

- Phase 8 COMPLETED;
- manual workflows understood and stable.

## Exit Criteria

- required automation works;
- expected failures are handled safely;
- failed refresh does not appear successful;
- relevant VBA source exported and versioned.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.9.0

---

# 13. Phase 10 — Reporting & Dashboard

## Status

NOT STARTED

## Target Version

v0.10.0

## Depends On

Phase 9 — Automation.

## Objective

Build the final operational reporting and management presentation layer on top of a stable validated backend.

## Scope

- complete `40_RPT_Replenishment`;
- complete `41_DASH_Management`;
- design management Key Performance Indicators;
- integrate inventory signals;
- integrate procurement signals;
- integrate supplier signals;
- implement useful interactive filtering;
- establish clear visual hierarchy;
- validate accessibility;
- finalize navigation;
- integrate relevant automation.

## Deliverables

- professional operational replenishment report;
- management dashboard;
- reconciled indicators;
- final user-navigation experience;
- reporting test evidence;
- `PHASE_10_CLOSEOUT.md`.

## Learning Focus

- dashboard design;
- information hierarchy;
- chart selection;
- operational vs management reporting;
- user experience;
- accessibility;
- report reconciliation.

## Entry Criteria

- Phase 9 COMPLETED;
- backend logic and automation stable.

## Exit Criteria

- operational report is actionable;
- dashboard metrics reconcile;
- navigation is functional;
- presentation is professional;
- important meaning does not depend only on color.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.10.0

---

# 14. Phase 11 — Testing & Hardening

## Status

NOT STARTED

## Target Version

v0.11.0

## Depends On

Phase 10 — Reporting & Dashboard.

## Objective

Perform complete system validation, resolve defects and harden ProcureFlow for final release.

## Scope

- end-to-end testing;
- formula tests;
- configuration tests;
- Power Query refresh tests;
- Quality Control tests;
- PivotTable tests;
- VBA tests;
- failure-path tests;
- edge cases;
- performance testing;
- protection testing;
- navigation testing;
- usability testing;
- User Acceptance Testing;
- full-dataset validation;
- defect correction;
- final reconciliation.

## Deliverables

- completed `TESTING.md`;
- resolved critical defects;
- accepted User Acceptance Testing scenarios;
- performance evidence;
- final hardening changes;
- `PHASE_11_CLOSEOUT.md`.

## Learning Focus

- structured testing;
- debugging;
- regression testing;
- edge-case reasoning;
- User Acceptance Testing;
- performance evaluation.

## Entry Criteria

- Phase 10 COMPLETED;
- complete functional system available for validation.

## Exit Criteria

- no critical known defect;
- mandatory acceptance tests pass;
- required User Acceptance Testing scenarios pass;
- performance acceptable;
- documentation reflects validated behavior;
- system is ready for portfolio release preparation.

## GitHub Gate

Complete phase branch workflow, Pull Request, merge, push and phase tag.

## Version

v0.11.0

---

# 15. Phase 12 — Documentation & Portfolio Release

## Status

NOT STARTED

## Target Version

v1.0.0

## Depends On

Phase 11 — Testing & Hardening.

## Objective

Finalize ProcureFlow as a professional, reproducible and portfolio-ready project.

## Scope

- update all canonical documentation;
- finalize `README.md`;
- complete architecture documentation;
- complete project specification;
- complete data dictionary;
- complete testing evidence;
- verify all phase closeouts;
- verify decisions;
- export final VBA source;
- produce screenshots;
- document user workflow;
- document refresh workflow;
- document installation / setup;
- review repository quality;
- final workbook cleanup;
- final Git history review;
- final acceptance against Definition of Done;
- create final tag;
- create GitHub Release.

## Deliverables

- final `ProcureFlow.xlsm`;
- final VBA source;
- final README;
- final canonical documents;
- screenshots;
- complete phase documentation;
- final test evidence;
- tag `v1.0.0`;
- GitHub Release.

## Learning Focus

- technical communication;
- portfolio presentation;
- release management;
- reproducibility;
- maintainability;
- professional project handoff.

## Entry Criteria

- Phase 11 COMPLETED;
- no critical known defects;
- mandatory acceptance evidence available.

## Exit Criteria

All applicable mandatory Acceptance Criteria and Definition of Done items are satisfied or formally superseded through an approved decision.

`CURRENT_STATE.md` reports:

Project Status: COMPLETED

Current Version: v1.0.0

## GitHub Gate

- final Pull Request merged;
- final `main` pushed;
- final phase tag pushed;
- tag `v1.0.0` pushed;
- GitHub Release created;
- repository reviewed as portfolio-ready.

## Version

v1.0.0

---

# 16. Roadmap Summary

| Phase | Name | Status | Target Version |
|---:|---|---|---|
| 0 | Project Design | IN PROGRESS | v0.1.0 |
| 1 | Data Design | NOT STARTED | v0.2.0 |
| 2 | Workbook Foundation | NOT STARTED | v0.3.0 |
| 3 | Power Query Pipeline | NOT STARTED | v0.4.0 |
| 4 | Operational Model | NOT STARTED | v0.5.0 |
| 5 | Business Logic & Advanced Formulas | NOT STARTED | v0.6.0 |
| 6 | Quality Control System | NOT STARTED | v0.7.0 |
| 7 | Analysis & PivotTables | NOT STARTED | v0.8.0 |
| 8 | VBA Foundations | NOT STARTED | v0.8.1 |
| 9 | Automation | NOT STARTED | v0.9.0 |
| 10 | Reporting & Dashboard | NOT STARTED | v0.10.0 |
| 11 | Testing & Hardening | NOT STARTED | v0.11.0 |
| 12 | Documentation & Portfolio Release | NOT STARTED | v1.0.0 |

---

# 17. Current Immediate Next Step

Complete Phase 0.9:

Canonical Documents & Phase 0 Publication.

After Phase 0 passes its GitHub gate, the next authorized phase will be:

Phase 1 — Data Design.

