# ProcureFlow — Project Specification

## Document Status

Status: CONFIRMED DESIGN BASELINE  
Project State: PRE-PROJECT  
Current Phase: Phase 0 — Project Design  
Target Phase 0 Version: v0.1.0

This document defines the approved functional and business specification of ProcureFlow before implementation begins.

Nothing described as a requirement or design decision in this document should be interpreted as implemented unless the project state explicitly marks it as IMPLEMENTED.

---

# 1. Project Overview

## 1.1 Project Name

ProcureFlow

## 1.2 Project Type

Excel-based Procurement & Inventory Management System.

## 1.3 Purpose

ProcureFlow will be an end-to-end business solution built primarily in Microsoft Excel for procurement, inventory monitoring, replenishment, supplier performance, quality control, operational reporting and management reporting.

The project is intended to resemble a maintainable business tool rather than:

- an isolated dashboard;
- a technical exercise;
- a collection of unrelated Excel formulas;
- a VBA demonstration;
- a set of disconnected analytics.

## 1.4 Professional Objective

The final project must demonstrate the ability to design, build, validate, automate, document and version a professional Excel solution using:

- Microsoft Excel 365 Desktop;
- Excel Tables and Structured References;
- Power Query;
- advanced Excel formulas;
- Dynamic Arrays;
- PivotTables;
- PivotCharts;
- Slicers;
- Timelines;
- Data Validation;
- Conditional Formatting;
- Named Ranges and Named Formulas;
- Macros;
- Visual Basic for Applications;
- Git;
- GitHub;
- technical documentation.

---

# 2. Business Problem

ProcureFlow addresses a business environment in which procurement and inventory information exists across operational data sources but lacks a unified process for monitoring, replenishment and decision support.

The system must help address problems such as:

- poor visibility of inventory availability;
- risk of stockouts;
- excess inventory;
- backorders;
- blocked stock;
- manual replenishment decisions;
- limited visibility into open Purchase Orders;
- supplier delays;
- variable Lead Times;
- partial receipts;
- supplier quality incidents;
- manual reporting;
- weak traceability of calculations;
- insufficient data-quality controls.

ProcureFlow should enable business users to answer questions such as:

- Which products require replenishment?
- At which site?
- How much should be ordered?
- Which products are already in stockout or critical condition?
- Which Purchase Orders remain open as of a selected Reporting Date?
- Which suppliers deliver late?
- How variable are supplier Lead Times?
- Which suppliers generate quality incidents?
- Where does excess inventory exist?
- Can the current data and calculations be trusted?

---

# 3. Target Users

## 3.1 Primary User — Procurement / Inventory Analyst

Responsibilities supported by ProcureFlow:

- refresh data;
- review data quality;
- monitor inventory;
- identify replenishment requirements;
- review open Purchase Orders;
- investigate exceptions;
- analyze supplier performance;
- prepare operational reports.

## 3.2 Secondary User — Buyer / Purchasing Specialist

Primary needs:

- determine what requires purchase;
- determine recommended quantities;
- identify the relevant site;
- monitor outstanding purchases;
- prioritize urgent requirements.

## 3.3 Management User — Procurement / Operations Manager

Primary needs:

- review key performance indicators;
- monitor inventory risk;
- understand procurement status;
- evaluate supplier delivery and quality;
- identify operational exceptions and trends.

---

# 4. Technology Scope

## 4.1 Target Platform

Microsoft Excel 365 Desktop for Windows.

## 4.2 Approved Core Technologies

- Excel Tables;
- Structured References;
- Power Query;
- Excel formulas;
- Dynamic Arrays;
- PivotTables;
- PivotCharts;
- Slicers;
- Timelines;
- Data Validation;
- Conditional Formatting;
- Named Ranges;
- Named Formulas;
- Macros;
- Visual Basic for Applications;
- Git;
- GitHub;
- Markdown documentation.

## 4.3 Technologies Not Included in the Core Solution

ProcureFlow will not use:

- Python;
- SQL as an application dependency;
- PostgreSQL;
- Power BI;
- BigQuery;
- Snowflake;
- Databricks;
- dbt.

Power Pivot / Excel Data Model is not part of the initial architecture and will only be introduced if later evidence demonstrates a real need.

---

# 5. Official Dataset

## 5.1 Dataset

Aerospace Supply Chain Performance & Forecasting.

## 5.2 Official Source Files

ProcureFlow will use:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

## 5.3 Source Data Policy

The original CSV files are treated as immutable raw data.

They must not be manually edited as part of the normal ProcureFlow workflow.

Required cleaning, typing, transformation, validation and preparation will be performed through Power Query.

## 5.4 Dataset Characteristics

The selected dataset provides approximately:

- 300 products / parts;
- 6 operational sites;
- 40 suppliers;
- 156 weekly periods;
- approximately 280,800 inventory-history records;
- approximately 30,000 Purchase Order records;
- hundreds of supplier-quality incidents.

The data supports analysis of:

- inventory;
- consumption;
- backorders;
- blocked stock;
- forecast;
- purchasing;
- actual receipt dates;
- promised dates;
- Lead Times;
- partial receipts;
- supplier quality;
- site-level inventory behavior.

## 5.5 Known Dataset Limitations

The design explicitly accepts these limitations:

- Purchase Orders do not use a traditional Header → multiple Lines model.
- Each product effectively has one primary supplier in the source.
- There is no rich independent supplier master.
- There is no rich independent site / warehouse master.
- Minimum Order Quantity is not provided.
- Order multiples are not provided.
- The dataset is synthetic rather than confidential real-company data.

ProcureFlow must not invent unsupported business data merely to fill these gaps.

---

# 6. Project Scope

## 6.1 In Scope

ProcureFlow will cover:

- product master data;
- supplier identifiers;
- site identifiers;
- Purchase Orders;
- historical inventory;
- historical consumption;
- blocked stock;
- backorders;
- open Purchase Orders;
- inventory position;
- demand analysis;
- Lead Time analysis;
- Stock de Seguridad / Safety Stock;
- Punto de Reposición / Reorder Point;
- target inventory;
- recommended order quantity;
- inventory-status classification;
- excess inventory;
- supplier delivery performance;
- partial receipts;
- supplier quality incidents;
- Power Query ingestion;
- data-quality controls;
- PivotTable analysis;
- operational reporting;
- management dashboard;
- VBA automation;
- Git/GitHub versioning;
- technical documentation.

## 6.2 Out of Scope for v1.0.0

The following are explicitly outside the mandatory scope:

- Enterprise Resource Planning implementation;
- direct SAP integration;
- direct Oracle integration;
- direct connection to corporate ERP systems;
- accounting;
- general ledger;
- Accounts Payable;
- invoice processing;
- supplier payments;
- tax accounting;
- sales management;
- customer relationship management;
- complete warehouse-management-system functionality;
- advanced logistics optimization;
- machine-learning forecasting;
- custom web application;
- database-backed application architecture;
- multi-user concurrent transaction processing;
- Power BI;
- mandatory Power Pivot;
- artificial creation of alternate suppliers not present in the dataset;
- invented Minimum Order Quantities;
- invented purchasing constraints unsupported by the source.

---

# 7. Conceptual Data Model

ProcureFlow will use the following conceptual entities.

## 7.1 Product

Source:

`parts_master.csv`

Grain:

1 row = 1 product / part.

## 7.2 Site

Derived from unique Site identifiers in `supply_chain_history.csv`.

Grain:

1 row = 1 site.

A Site is treated as an operational inventory location functionally similar to a warehouse, but ProcureFlow will not invent location attributes not provided by the source.

## 7.3 Supplier

Derived from Supplier identifiers present in procurement and quality sources.

Grain:

1 row = 1 supplier.

## 7.4 Inventory History

Source:

`supply_chain_history.csv`

Grain:

1 Product × 1 Site × 1 Week.

## 7.5 Purchase Orders

Source:

`purchase_orders.csv`

Grain:

1 row = 1 purchase record / Purchase Order record for one product.

## 7.6 Quality Incidents

Source:

`quality_incidents.csv`

Grain:

1 row = 1 supplier-quality incident.

## 7.7 Date Dimension

A date dimension will be derived to support consistent temporal filtering, reporting and analysis.

Its exact physical implementation will be determined during implementation.

---

# 8. Functional Requirements

## 8.1 Master Data

### FR-01 — Product Master

The system must provide a validated product master based on the official source.

### FR-02 — Supplier Master

The system must maintain a unique supplier list derived from official source data.

### FR-03 — Site Master

The system must maintain a unique site list derived from official source data.

---

## 8.2 Data Ingestion

### FR-04 — Automated Ingestion

The four official source files must be ingested through Power Query.

### FR-05 — Refresh

Normal data updates must not require manual copy/paste of operational records.

### FR-06 — Explicit Data Types

Dates, identifiers, quantities, costs and other relevant fields must use appropriate data types.

---

## 8.3 Inventory Management

### FR-07 — Inventory Position Monitoring

The system must provide inventory information by Product × Site.

### FR-08 — Historical Inventory

The system must support historical inventory analysis.

### FR-09 — Consumption Analysis

The system must analyze historical consumption by Product × Site.

### FR-10 — Backorder Monitoring

The system must identify and analyze backorders.

### FR-11 — Blocked Stock Monitoring

Blocked stock must be distinguished from stock available for consumption.

---

## 8.4 Replenishment

### FR-12 — Historical Demand

The system must calculate demand indicators using approved historical consumption rules.

### FR-13 — Safety Stock

The system must calculate Safety Stock using a documented methodology.

### FR-14 — Reorder Point

The system must calculate Reorder Point using the approved business rule.

### FR-15 — Inventory Status

The system must classify Product × Site inventory positions using approved operational states.

### FR-16 — Open Purchase Orders

Open Purchase Orders must be incorporated into replenishment decisions where applicable.

### FR-17 — Recommended Order Quantity

The system must produce a documented recommended order quantity.

---

## 8.5 Procurement

### FR-18 — Purchase Order Monitoring

The system must support Purchase Order analysis.

### FR-19 — Open Purchase Orders as of Reporting Date

The system must determine which Purchase Orders were open at the selected Reporting Date.

### FR-20 — Late Purchase Orders

The system must identify late receipts.

### FR-21 — Partial Receipts

The system must identify receipts where received quantity is below ordered quantity.

### FR-22 — Procurement Spend

The system must calculate purchasing value only where source fields support a valid interpretation.

---

## 8.6 Supplier Performance

### FR-23 — Actual Lead Time

Actual Lead Time must be calculated using order and receipt dates.

### FR-24 — Lead-Time Variability

The system must analyze variation in actual Lead Times.

### FR-25 — On-Time Delivery

The system must calculate a reproducible On-Time Delivery metric.

### FR-26 — Supplier Quality

The system must incorporate supplier-quality incidents.

### FR-27 — Supplier Performance View

The system must provide transparent supplier-performance metrics.

An aggregate Supplier Score is not mandatory and must not be introduced without a documented methodology.

---

## 8.7 Data Quality

### FR-28 — Duplicate Detection

The system must identify duplicates where uniqueness is required.

### FR-29 — Referential Integrity

References between products, suppliers, sites and transactional data must be validated.

### FR-30 — Required Fields

Mandatory fields must be checked for missing values.

### FR-31 — Logical Date Validation

Impossible or inconsistent date relationships must be detected.

### FR-32 — Quantity Validation

Invalid or logically inconsistent quantities must be identified.

---

## 8.8 Analysis and Reporting

### FR-33 — Inventory Analysis

The system must provide aggregated inventory analysis.

### FR-34 — Procurement Analysis

The system must provide aggregated procurement analysis.

### FR-35 — Supplier Analysis

The system must provide supplier-performance and supplier-quality analysis.

### FR-36 — Interactive Filtering

Slicers and Timelines must be used where they provide meaningful analytical value.

### FR-37 — Operational Replenishment Report

A dedicated operational replenishment report must provide actionable purchasing recommendations.

### FR-38 — Management Dashboard

A management dashboard must summarize relevant inventory, procurement and supplier information.

---

## 8.9 Automation

### FR-39 — Refresh Automation

VBA may automate the complete refresh workflow once the manual workflow is stable.

### FR-40 — Quality-Control Automation

VBA may automate or orchestrate quality-control execution and navigation.

### FR-41 — Reporting Automation

VBA may automate report preparation and export where professionally justified.

---

# 9. Business Rules

## 9.1 Reporting Date

Reporting Date is a configurable system parameter.

The initial maximum Reporting Date for the current dataset is:

23-Dec-2024

The parameter must remain within the valid data range.

## 9.2 Demand History Window

Default historical demand window:

26 weeks.

Demand calculations are performed by Product × Site.

Weeks with zero consumption are valid observations and must not be discarded automatically.

## 9.3 Review Period

Default Review Period:

1 week.

This aligns with the weekly grain of the historical source.

## 9.4 Lead Time

Actual Lead Time:

Receipt Date − Order Date

The system will use historical actual Lead Times to calculate:

- average Lead Time;
- Lead-Time variability.

The product-master Lead Time may be used as a fallback if insufficient history exists.

## 9.5 Service Levels by Criticality

Initial configurable policy:

- Criticality A: 99%
- Criticality B: 97%
- Criticality C: 95%

These values are configuration parameters and must not be hidden constants inside formulas.

## 9.6 Available Stock

Conceptual rule:

Available Stock = On-Hand Inventory − Blocked Stock

Available Stock must not be negative; implementation will apply an appropriate minimum-zero rule where required.

## 9.7 Backorders

Backorders represent demand already pending because available inventory was insufficient.

Backorders reduce Inventory Position.

## 9.8 Safety Stock

Safety Stock must incorporate:

- demand variability;
- Lead-Time variability;
- service level.

The formula must use consistent time units.

## 9.9 Reorder Point

Conceptual rule:

Reorder Point = Expected Demand During Lead Time + Safety Stock

## 9.10 Open Purchase Order

A Purchase Order is open as of Reporting Date when:

Order Date <= Reporting Date < Receipt Date

While open, the originally ordered quantity is treated as incoming inventory.

ProcureFlow must not use future receipt information retroactively to alter what was known at the Reporting Date.

## 9.11 Inventory Position

Conceptual rule:

Inventory Position = Available Stock + Open Purchase Order Quantity − Backorders

## 9.12 Target Stock

Target Stock must cover:

- expected demand during Lead Time;
- one Review Period;
- Safety Stock.

Conceptually:

Target Stock = Average Demand × (Lead Time + Review Period) + Safety Stock

with consistent units.

## 9.13 Recommended Order Quantity

A positive order recommendation is only produced when:

Inventory Position <= Reorder Point

Conceptually:

Recommended Order Quantity = Target Stock − Inventory Position

The result must never be negative and will be rounded appropriately to whole units.

Minimum Order Quantity and order multiples are not applied because the official dataset does not provide them.

## 9.14 Inventory Status Priority

Inventory states are evaluated in this priority order:

1. STOCKOUT
2. CRITICAL
3. REORDER
4. ATTENTION
5. EXCESS
6. HEALTHY

### STOCKOUT

No usable stock exists and unmet demand / backorder exists.

### CRITICAL

A backorder exists or Available Stock is below Safety Stock.

### REORDER

Inventory Position is at or below Reorder Point.

### ATTENTION

Inventory Position is above Reorder Point but within approximately one week of average demand above the Reorder Point.

### EXCESS

Inventory Position exceeds Target Stock by more than the configured Excess Buffer.

Initial Excess Buffer:

4 weeks of average demand.

### HEALTHY

No higher-priority state applies.

## 9.15 No Recent Demand

If total consumption during the configured 26-week window is zero:

`NO_RECENT_DEMAND`

must be raised as an additional flag.

The system must not automatically create a purchasing recommendation solely because of a statistical formula in such cases.

## 9.16 On-Time Delivery

A received Purchase Order is On Time when:

Actual Receipt Date <= Promised Date

The On-Time Delivery percentage is calculated using received orders only.

## 9.17 Partial Receipt

A Purchase Order is considered partially received when:

Received Quantity < Ordered Quantity

## 9.18 Forecast

Forecast data from the source will be preserved for analysis.

Forecast will not be the primary input of the initial replenishment engine.

Forecast accuracy analysis is optional and must not block v1.0.0 unless later promoted into mandatory scope through a documented decision.

## 9.19 Supplier Risk

Supplier Risk from the source may be used for:

- filtering;
- alerts;
- reporting;
- analysis.

It will not automatically multiply Safety Stock because observed Lead-Time variability already captures supply uncertainty and double-counting risk must be avoided.

---

# 10. Configurable Business Parameters

The following parameters must be configurable rather than embedded as hidden constants:

- Reporting Date
- Demand History Weeks
- Review Period Weeks
- Service Level A
- Service Level B
- Service Level C
- Excess Buffer Weeks

Initial values:

| Parameter | Initial Value |
|---|---:|
| Demand History Weeks | 26 |
| Review Period Weeks | 1 |
| Service Level A | 99% |
| Service Level B | 97% |
| Service Level C | 95% |
| Excess Buffer Weeks | 4 |
| Reporting Date | Valid selected date |

---

# 11. Non-Functional Requirements

## NFR-01 — Maintainable Structure

The workbook must clearly separate:

- configuration;
- control;
- processed data;
- calculations;
- analysis;
- reporting.

## NFR-02 — Naming Conventions

Consistent naming must be used for:

- worksheets;
- tables;
- Power Query queries;
- defined names;
- VBA procedures;
- quality controls.

## NFR-03 — Structured References

Excel Tables and Structured References should be preferred over fragile coordinate-based references where appropriate.

## NFR-04 — No Hidden Business Constants

Business-policy numbers must be centralized in configuration.

## NFR-05 — Auditability

Core business calculations must remain inspectable and explainable.

## NFR-06 — Data Lineage

The project must distinguish between:

- source values;
- Power Query transformations;
- configurable inputs;
- Excel calculations;
- analytical outputs.

## NFR-07 — Decision Traceability

Important changes to architecture, rules or scope must be recorded in `DECISIONS.md`.

## NFR-08 — Reasonable Performance

Normal user interaction must remain practical with the complete official dataset.

## NFR-09 — Avoid Unnecessary Volatile Formulas

Volatile functions must only be used when justified.

## NFR-10 — Avoid Unnecessary Full-Column Calculations

Large-range formulas must be designed with performance in mind.

## NFR-11 — Measured Optimization

Optimization decisions must be based on observed performance rather than premature complexity.

## NFR-12 — Reproducible Refresh

The documented refresh workflow must rebuild required processed data without manual copy/paste.

## NFR-13 — Raw Data Immutability

Original CSV files remain unchanged.

## NFR-14 — No Hidden Manual Steps

Any unavoidable manual procedure must be documented.

## NFR-15 — Clear Navigation

Users must be able to reach core functional areas from `00_HOME`.

## NFR-16 — Editable Inputs Clearly Identified

Configurable cells must be visually and technically distinguishable from formulas and outputs.

## NFR-17 — Concise User Instructions

Workbook instructions must be practical and concise.

## NFR-18 — Consistent Visual Design

Formatting, typography, hierarchy and states must follow a coherent visual system.

## NFR-19 — Avoid Decorative Noise

The solution must avoid unnecessary 3D charts, excessive colors and decorative elements without functional value.

## NFR-20 — Clear Visual Hierarchy

Key performance indicators, warnings, details and navigation must be visually distinguishable.

## NFR-21 — Editable Cells Remain Editable

Protection must preserve legitimate user inputs.

## NFR-22 — Critical Formulas Protected

Critical formulas and structures should be protected against accidental modification.

## NFR-23 — Power Query Output Is Not Manual Input

Users must not treat refreshed DATA tables as manually editable operational inputs.

## NFR-24 — Error Transparency

Errors must not be indiscriminately hidden using blanket IFERROR logic.

## NFR-25 — Understandable Error Messages

User-facing errors must explain the problem and required action where possible.

## NFR-26 — Safe Automation

A VBA failure must not leave the workbook falsely marked as successfully refreshed or validated.

## NFR-27 — Explicit Critical Actions

Important automated actions must have understandable intent and behavior.

## NFR-28 — Target Platform

Microsoft Excel 365 Desktop for Windows.

## NFR-29 — Controlled Workbook Size

The workbook must avoid redundant raw-data copies, unnecessary loaded queries and excessive unused objects.

## NFR-30 — Accessibility

Status and meaning must not depend on color alone.

---

# 12. Workbook Language

The final workbook and technical naming convention will primarily use business English.

Examples:

- AvailableStock
- SafetyStock
- ReorderPoint
- RecommendedOrderQty

Development explanations may be provided in Spanish, with English terminology introduced progressively.

---

# 13. Quality-Control Requirements

ProcureFlow must provide a formal Quality Control layer using:

- PASS
- WARNING
- FAIL

Minimum control areas include:

- source-file availability;
- required columns;
- data types;
- duplicate keys;
- Product integrity;
- Supplier integrity;
- Site integrity;
- mandatory fields;
- invalid quantities;
- invalid dates;
- invalid configuration parameters;
- negative impossible calculated values;
- invalid inventory statuses;
- Power Query errors;
- key row-count reconciliations;
- last successful refresh.

A critical FAIL must invalidate the overall quality status.

Unusual but potentially valid data must not automatically be treated as invalid.

---

# 14. Reporting Requirements

## 14.1 Operational Replenishment Report

The operational report must allow a user to identify at minimum:

- Product;
- Site;
- Inventory Status;
- Available Stock;
- Open Purchase Order Quantity;
- Reorder Point;
- Recommended Order Quantity.

## 14.2 Management Dashboard

The dashboard must communicate at minimum:

- overall inventory situation;
- replenishment risk;
- procurement status;
- supplier performance;
- relevant operational exceptions.

Important dashboard indicators must be traceable to validated underlying data.

---

# 15. Automation Principles

Visual Basic for Applications will be used for orchestration and automation, not as the hidden mathematical core of ProcureFlow.

Candidate automations include:

- complete refresh workflow;
- PivotTable refresh;
- quality-control execution;
- timestamps;
- navigation;
- replenishment-report preparation;
- export.

The user must first understand the relevant manual process and VBA foundations before professional automation is introduced.

---

# 16. Acceptance Policy

Acceptance criteria use two priority levels.

## MUST

Mandatory for v1.0.0 unless formally changed through an approved decision.

## SHOULD

Desirable but not necessarily release-blocking if omission is justified and documented.

The complete executable acceptance evidence will be maintained in `TESTING.md` during implementation.

Mandatory acceptance areas include:

- dataset ingestion;
- Power Query refresh;
- data-model integrity;
- replenishment calculations;
- procurement analysis;
- supplier performance;
- configuration;
- quality controls;
- PivotTables;
- operational reporting;
- dashboard reconciliation;
- VBA behavior;
- performance;
- user experience;
- documentation;
- Git/GitHub;
- phase-to-phase handoff.

---

# 17. Definition of Done

ProcureFlow may only be declared complete and released as `v1.0.0` when all mandatory applicable acceptance criteria and all mandatory Definition of Done conditions are satisfied or have been formally changed through an approved decision.

At minimum:

1. Mandatory functional requirements are implemented.
2. Mandatory non-functional requirements are satisfied.
3. The four official Aerospace source files can feed the solution through the documented workflow.
4. Raw files remain unchanged.
5. Power Query refresh is reproducible.
6. The final data model is documented.
7. The replenishment engine implements the approved business rules.
8. Configuration parameters are externalized from formulas.
9. Procurement and supplier analysis are functional.
10. `02_CONTROL` is operational.
11. Critical quality controls pass on the valid release dataset.
12. Reconciliations pass.
13. The operational replenishment report is usable.
14. The management dashboard is usable and reconciled.
15. PivotTables and interactive filters refresh correctly.
16. VBA automation handles expected failure scenarios appropriately.
17. Relevant VBA modules are exported as versionable source files.
18. ProcureFlow operates with the complete official dataset.
19. Critical formulas and structures are protected appropriately.
20. Navigation from `00_HOME` is functional.
21. End-to-end and User Acceptance Testing have been completed.
22. No critical known defects remain open.
23. Canonical documentation reflects the actual implementation.
24. Every phase has a closeout document.
25. GitHub contains the complete approved project history.
26. A new project participant can reconstruct project state from the repository without reading prior chat conversations.
27. The public README is complete.
28. Representative screenshots exist.
29. `main` contains the final accepted state.
30. Git tag `v1.0.0` exists.
31. A final GitHub Release exists where appropriate.
32. `CURRENT_STATE.md` marks the project as COMPLETED.

---

# 18. Scope Protection for v1.0.0

The following are not mandatory for the initial Definition of Done:

- Power Pivot;
- custom forecasting model;
- SAP integration;
- Oracle integration;
- accounting;
- invoicing;
- Accounts Payable;
- payment processing;
- Power BI;
- web application;
- database application;
- invented alternate suppliers;
- invented Minimum Order Quantities;
- invented order multiples;
- capabilities not approved as part of the v1.0.0 scope.

Future versions may extend the system after v1.0.0.

---

# 19. Project Governance

## 19.1 Status Labels

The project uses:

- `[PROPUESTO]` — proposed, not yet approved;
- `[CONFIRMADO]` — approved design or decision;
- `[IMPLEMENTADO]` — exists and has evidence;
- `[PENDIENTE]` — not yet completed;
- `[SUPERADO]` — consciously replaced by a later decision.

## 19.2 Source of Truth

Canonical documents are:

- `PROJECT_SPEC.md`
- `ARCHITECTURE.md`
- `ROADMAP.md`
- `DECISIONS.md`
- `CURRENT_STATE.md`

Supporting canonical documents include:

- `DATA_DICTIONARY.md`
- `TESTING.md`
- `docs/phases/PHASE_XX_CLOSEOUT.md`

## 19.3 Phase Handoff Rule

A new execution phase must not begin until the previous phase:

- meets its exit criteria;
- is validated;
- is documented;
- is merged to `main`;
- is pushed to GitHub;
- has its phase-complete tag.

GitHub must contain enough information for a new phase chat to reconstruct project context without requiring a manual summary from the previous chat.

---

# 20. Current Implementation Status

As of the Phase 0 design baseline:

- Dataset selection: CONFIRMED
- Functional requirements: CONFIRMED
- Non-functional requirements: CONFIRMED
- Business rules: CONFIRMED
- Workbook architecture: CONFIRMED
- Technical architecture: CONFIRMED
- Quality strategy: CONFIRMED
- Git/GitHub strategy: CONFIRMED
- Roadmap: CONFIRMED
- Acceptance criteria: CONFIRMED
- Definition of Done: CONFIRMED

Implementation status:

PENDING

No workbook functionality is considered implemented at this stage.
