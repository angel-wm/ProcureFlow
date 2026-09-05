# ProcureFlow — Decision Log

## Document Status

Status: CONFIRMED  
Project State: PRE-PROJECT  
Current Phase: Phase 0 — Project Design  
Target Phase 0 Version: v0.1.0

This document records material project decisions that affect ProcureFlow scope, architecture, business rules, implementation strategy, governance or release management.

---

# 1. Decision Statuses

Decisions may use the following statuses:

- CONFIRMED — approved and currently valid.
- SUPERSEDED — replaced by a later approved decision.
- PROPOSED — under consideration and not yet authoritative.

A material change to a CONFIRMED decision should be documented explicitly rather than silently modifying project behavior.

---

# 2. Decision Log

## DEC-001 — Official Dataset

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will use the Aerospace Supply Chain Performance & Forecasting dataset as its official base dataset.

**Rationale:**  
It provides stronger support than the evaluated alternatives for historical inventory by site, consumption, backorders, blocked stock, procurement, receipt dates, Lead Times and supplier quality.

AdventureWorks remains outside ProcureFlow and may be used in a separate future project.

---

## DEC-002 — Official Source Files

**Status:** CONFIRMED

**Decision:**  
The official source set consists of:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

**Rationale:**  
Together these files provide the minimum required procurement, inventory, supplier and quality domains for ProcureFlow.

---

## DEC-003 — Source Data Immutability

**Status:** CONFIRMED

**Decision:**  
Original source CSV files must remain unchanged.

Cleaning, typing, validation and preparation must occur reproducibly through Power Query.

**Rationale:**  
This preserves traceability and prevents undocumented manual preprocessing.

---

## DEC-004 — Core Analytical Grain

**Status:** CONFIRMED

**Decision:**  
The main operational replenishment model will use:

1 Product × 1 Site

as its calculation grain.

Historical inventory remains at:

1 Product × 1 Site × 1 Week.

**Rationale:**  
Replenishment decisions must reflect location-specific inventory conditions.

---

## DEC-005 — Do Not Invent Unsupported Business Data

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will not fabricate business attributes that are absent from the official dataset.

This includes, unless future evidence supports them:

- alternate suppliers;
- Purchase Order header/line structures;
- Minimum Order Quantities;
- order multiples;
- warehouse attributes.

**Rationale:**  
Portfolio realism must come from transparent modeling, not from invented source semantics.

---

## DEC-006 — Layered Workbook Architecture

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will use distinct workbook layers for:

- configuration;
- control;
- data;
- calculations;
- analysis;
- operational reporting;
- dashboard.

**Rationale:**  
Separation of responsibility improves maintainability, usability and auditability.

---

## DEC-007 — Raw Data Remains Outside the Workbook

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will not create a dedicated RAW worksheet containing copies of the original source files.

Raw CSV files remain external.

**Rationale:**  
This avoids redundant storage and preserves a clean ingestion architecture.

---

## DEC-008 — Power Query Responsibility

**Status:** CONFIRMED

**Decision:**  
Power Query will handle ingestion and reproducible data preparation.

Primary responsibilities include:

- source access;
- typing;
- objective cleaning;
- standardization;
- dimension preparation;
- fact preparation;
- appropriate pre-aggregation.

**Rationale:**  
Power Query is the appropriate Extract, Transform and Load layer within an Excel-centered solution.

---

## DEC-009 — Excel Formula Responsibility

**Status:** CONFIRMED

**Decision:**  
Core configurable operational business logic will primarily be implemented using auditable Excel formulas.

**Rationale:**  
Business rules such as Safety Stock, Reorder Point and recommended purchasing quantities should remain inspectable and educational rather than hidden in code.

---

## DEC-010 — VBA Responsibility

**Status:** CONFIRMED

**Decision:**  
Visual Basic for Applications will be used for orchestration and automation rather than as the hidden mathematical engine.

**Rationale:**  
VBA is most valuable for repeatable workflows such as refresh, validation, report preparation and export.

---

## DEC-011 — Power Pivot Deferred

**Status:** CONFIRMED

**Decision:**  
Power Pivot / Excel Data Model will not be included initially.

It may be added only if later evidence shows material value.

**Rationale:**  
The project should not add technology merely for demonstration purposes.

---

## DEC-012 — Operational Report Separate from Dashboard

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will maintain separate outputs for:

- operational replenishment;
- management reporting.

Planned sheets:

- `40_RPT_Replenishment`
- `41_DASH_Management`

**Rationale:**  
Operational purchasing decisions and management summaries serve different users and information needs.

---

## DEC-013 — Workbook Business Language

**Status:** CONFIRMED

**Decision:**  
The final workbook, field naming and technical nomenclature will primarily use business English.

Development explanations may be provided in Spanish.

**Rationale:**  
This improves professional portfolio value while allowing progressive learning.

---

## DEC-014 — Target Excel Platform

**Status:** CONFIRMED

**Decision:**  
ProcureFlow targets Microsoft Excel 365 Desktop for Windows.

**Rationale:**  
The planned stack depends on modern Excel capabilities including Dynamic Arrays, XLOOKUP and Power Query.

---

## DEC-015 — Inputs Must Be Separated

**Status:** CONFIRMED

**Decision:**  
User-editable business configuration must be clearly separated from formulas, refreshed data and outputs.

**Rationale:**  
This reduces accidental changes and improves usability.

---

## DEC-016 — Error Transparency

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will not indiscriminately hide errors through blanket error-handling formulas.

**Rationale:**  
Errors that affect trust in calculations must remain visible or be handled explicitly.

---

## DEC-017 — Workbook Protection

**Status:** CONFIRMED

**Decision:**  
Critical formulas and structures will be protected against accidental editing while legitimate inputs remain editable.

**Rationale:**  
Protection must support usability without obstructing auditability.

---

## DEC-018 — Repository Structure

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will use the approved structured repository containing:

- workbook;
- data;
- VBA source;
- documentation;
- phase closeouts;
- screenshots.

**Rationale:**  
The repository must represent the full engineering project rather than only the binary workbook.

---

## DEC-019 — Workbook Versioning

**Status:** CONFIRMED

**Decision:**  
`ProcureFlow.xlsm` will be versioned in Git despite being a binary file.

**Rationale:**  
The workbook is the primary executable artifact and must be preserved with each approved project state.

---

## DEC-020 — VBA Source Export

**Status:** CONFIRMED

**Decision:**  
Relevant VBA modules will be exported as text files such as `.bas`.

**Rationale:**  
Text-based VBA source enables direct inspection, review and meaningful Git diffs.

---

## DEC-021 — Raw Dataset Repository Policy

**Status:** CONFIRMED

**Decision:**  
The complete raw Aerospace dataset will not initially be committed to Git.

`data/raw/` remains ignored except for structural placeholder files.

**Rationale:**  
This reduces repository size and keeps source-distribution concerns separate from project code and documentation.

---

## DEC-022 — Canonical Documentation Format

**Status:** CONFIRMED

**Decision:**  
Canonical project documents will use Markdown rather than JSON.

**Rationale:**  
Markdown is easier for humans, GitHub and future project chats to inspect while remaining versionable.

---

## DEC-023 — GitHub Releases Instead of Release Folder

**Status:** CONFIRMED

**Decision:**  
The repository will not maintain a permanent `releases/` directory.

Formal release artifacts will use GitHub Releases when justified.

**Rationale:**  
GitHub already provides purpose-built release management.

---

## DEC-024 — Phase-Based GitHub Handoff

**Status:** CONFIRMED

**Decision:**  
GitHub will be the authoritative handoff mechanism between major project-phase chats.

**Rationale:**  
A new chat must be able to reconstruct project state without depending on conversational memory.

---

## DEC-025 — Main Contains Completed States Only

**Status:** CONFIRMED

**Decision:**  
The `main` branch represents only completed and approved phase states.

**Rationale:**  
This keeps `main` stable and authoritative.

---

## DEC-026 — Phase Branches

**Status:** CONFIRMED

**Decision:**  
Execution phases will use dedicated branches following a pattern similar to:

`phase/01-data-design`

**Rationale:**  
Phase branches isolate active work from the last approved project state.

---

## DEC-027 — Local-First Development with Mandatory Publication

**Status:** CONFIRMED

**Decision:**  
Work may remain local while a phase is in progress.

Immediately after a phase is completed, its approved state must be published to GitHub before the next phase begins.

**Rationale:**  
This supports practical local development while preserving reliable phase handoffs.

---

## DEC-028 — Pull Request per Phase

**Status:** CONFIRMED

**Decision:**  
Each execution phase should conclude through a Pull Request before merging into `main`, even for a solo project.

**Rationale:**  
Pull Requests create a professional review and historical checkpoint.

---

## DEC-029 — Phase Tags

**Status:** CONFIRMED

**Decision:**  
Each completed phase will receive an immutable phase-complete tag.

Example:

`phase-0-complete`

**Rationale:**  
Tags create permanent historical reference points independent of branch movement.

---

## DEC-030 — Semantic Versioning

**Status:** CONFIRMED

**Decision:**  
ProcureFlow will use a `MAJOR.MINOR.PATCH` version strategy.

The final accepted first release will be:

`v1.0.0`

**Rationale:**  
Explicit versions make project evolution easier to understand and present professionally.

---

## DEC-031 — GitHub Handoff Protocol

**Status:** CONFIRMED

**Decision:**  
A new phase chat will begin by reviewing, at minimum:

1. `CURRENT_STATE.md`
2. previous phase closeout;
3. `ROADMAP.md`;
4. relevant architecture and decisions.

**Rationale:**  
This establishes a reproducible context-loading procedure.

---

## DEC-032 — Sequential Roadmap

**Status:** CONFIRMED

**Decision:**  
The initial ProcureFlow roadmap will execute phases sequentially rather than in parallel.

**Rationale:**  
Later phases depend strongly on stable outputs from earlier phases.

---

## DEC-033 — Formal Phase Gates

**Status:** CONFIRMED

**Decision:**  
Every phase has explicit entry criteria, exit criteria and a GitHub gate.

**Rationale:**  
A phase must be objectively complete rather than informally abandoned.

---

## DEC-034 — Version Mapping by Phase

**Status:** CONFIRMED

**Decision:**  
The planned version progression is:

- Phase 0 → `v0.1.0`
- Phase 1 → `v0.2.0`
- Phase 2 → `v0.3.0`
- Phase 3 → `v0.4.0`
- Phase 4 → `v0.5.0`
- Phase 5 → `v0.6.0`
- Phase 6 → `v0.7.0`
- Phase 7 → `v0.8.0`
- Phase 8 → `v0.8.1`
- Phase 9 → `v0.9.0`
- Phase 10 → `v0.10.0`
- Phase 11 → `v0.11.0`
- Phase 12 → `v1.0.0`

**Rationale:**  
Versions provide visible project maturity checkpoints.

---

## DEC-035 — Meaning of v1.0.0

**Status:** CONFIRMED

**Decision:**  
`v1.0.0` is reserved for the fully accepted ProcureFlow system after all required phases and release criteria are satisfied.

**Rationale:**  
The version must signify a complete portfolio-ready system rather than an arbitrary milestone.

---

## DEC-036 — Learning and Debugging Chats Are Non-Authoritative

**Status:** CONFIRMED

**Decision:**  
Educational, experimental and debugging conversations do not automatically modify official project state.

A change only becomes authoritative when incorporated into the canonical project workflow.

**Rationale:**  
This prevents exploratory work from silently changing requirements or architecture.

---

## DEC-037 — Acceptance Criteria Policy

**Status:** CONFIRMED

**Decision:**  
Acceptance criteria classified as MUST are mandatory for the final release unless changed through a formal approved decision.

**Rationale:**  
Mandatory acceptance requirements must have release authority.

---

## DEC-038 — Definition of Done

**Status:** CONFIRMED

**Decision:**  
ProcureFlow uses a formal Definition of Done covering:

- functionality;
- data;
- Power Query;
- business logic;
- procurement;
- supplier performance;
- quality controls;
- reporting;
- PivotTables;
- VBA;
- performance;
- protection;
- testing;
- documentation;
- Git/GitHub;
- project handoff;
- portfolio readiness;
- final release.

**Rationale:**  
No single dashboard, macro or formula can by itself qualify the project as finished.

---

## DEC-039 — No Critical Known Defects at Release

**Status:** CONFIRMED

**Decision:**  
ProcureFlow `v1.0.0` cannot be released with known critical defects.

**Rationale:**  
A final release must represent a trusted and validated project state.

---

## DEC-040 — Scope Protection

**Status:** CONFIRMED

**Decision:**  
Features explicitly outside the approved v1.0.0 scope are not required to satisfy the Definition of Done.

**Rationale:**  
The project requires a controlled completion boundary and must not expand indefinitely.

---

## DEC-041 — Repository License

**Status:** CONFIRMED

**Decision:**  
ProcureFlow source code and original project documentation will be published under the MIT License.

Third-party datasets and other external materials retain their own applicable licenses, terms and ownership and are not relicensed by ProcureFlow.

**Rationale:**  
MIT provides a simple and permissive license suitable for a portfolio software project while maintaining the required copyright and license notice.

---
# 3. Confirmed Replenishment Policy

The following business-rule decisions are also part of the confirmed Phase 0 baseline.

## Reporting Date

A configurable Reporting Date determines the historical point of view used by relevant calculations.

Initial dataset maximum:

23-Dec-2024

## Historical Demand Window

Default:

26 weeks

by Product × Site.

Zero-consumption weeks remain valid observations.

## Review Period

Default:

1 week.

## Service Levels

Initial configurable values:

- Criticality A → 99%
- Criticality B → 97%
- Criticality C → 95%

## Safety Stock

Safety Stock will account for:

- demand variability;
- Lead-Time variability;
- service level.

## Reorder Point

Reorder Point consists of expected Lead-Time demand plus Safety Stock.

## Available Stock

Blocked inventory is excluded from Available Stock.

## Open Purchase Orders

An order is open when:

Order Date <= Reporting Date < Receipt Date

Future receipt knowledge must not be leaked retroactively into historical calculations.

## Inventory Position

Inventory Position includes:

- Available Stock;
- Open Purchase Orders;
- Backorders.

## Target Stock

Target Stock covers:

- Lead-Time demand;
- Review Period demand;
- Safety Stock.

## Recommended Order Quantity

A recommendation is triggered when Inventory Position is at or below Reorder Point.

The recommendation replenishes toward Target Stock and cannot be negative.

## Inventory Status Priority

Priority order:

1. STOCKOUT
2. CRITICAL
3. REORDER
4. ATTENTION
5. EXCESS
6. HEALTHY

## No Recent Demand

A separate `NO_RECENT_DEMAND` flag applies when the configured historical window contains zero total consumption.

## Excess Inventory

Initial Excess Buffer:

4 weeks of average demand.

## Supplier Delivery Performance

On-Time Delivery means:

Actual Receipt Date <= Promised Date

and uses received orders as its denominator.

## Supplier Performance Score

No arbitrary aggregate Supplier Score will be created initially.

Delivery, Lead Time, quantity and quality metrics remain individually visible.

## Supplier Risk

Source supplier-risk classification may support alerts, filtering and analysis.

It will not automatically multiply Safety Stock because doing so could double-count supply uncertainty already reflected in observed Lead-Time variability.

## Forecast

Source forecast data will be retained for analysis but will not drive the initial core replenishment engine.

---

# 4. Change-Control Rule

A new decision entry should be created when a future change materially affects:

- project scope;
- source datasets;
- business rules;
- workbook architecture;
- calculation architecture;
- Power Query responsibilities;
- VBA responsibilities;
- quality controls;
- Git/GitHub workflow;
- phase gates;
- release criteria.

Existing decision history should not be rewritten merely to hide that a change occurred.

If a decision is replaced, its status should become:

SUPERSEDED

and the replacement decision should identify what changed and why.

---

# 5. Current Decision Baseline

Decisions confirmed through Phase 0:

`DEC-001` through `DEC-040`

Current superseded decisions:

None.

Current proposed material decisions:

None.

Next major decision review:

Phase 1 — Data Design, if detailed source validation reveals evidence requiring a change to the Phase 0 design baseline.

