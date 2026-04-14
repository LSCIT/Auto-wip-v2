# Auto-WIP Schedule — Project Status & Task Plan
**Date:** March 31, 2026
**Prepared by:** Josh Garrison, IT Development
**Priority:** HIGHEST — Owner / Board / Senior Management escalation

---

## Background

The Auto-WIP Schedule is a macro-enabled Excel workbook used by Accounting and Operations to
produce the company's monthly Work-in-Progress schedule across four companies (AIC, APC, NESM,
WML). The tool connects to Viewpoint (Vista) for job cost data and provides a three-stage approval
workflow: Accounting Initial Review → Operations Review → Accounting Final Approval.

The original tool was built over 10+ years by an outside consultant (Michael Roberts, RCS Plan).
In early 2026 the decision was made to rebuild it in-house. As of March 31, the Vista data read
path is fully operational. The remaining work is the persistence database (LylesWIP on P&P),
the write path (saving overrides and approvals), the historical override data load, and
end-to-end workflow completion.

**Product platform:** Excel macro-enabled workbook. Committed to by the CEO to the board and ownership.

---

## System Architecture

The confirmed architecture uses a dedicated **LylesWIP database on the P&P server** as the
central persistence layer. This is required because Vista's SQL Server cannot be reached from
remote offices and job-site trailers, but P&P is accessible from all locations.

```
STAGE 1 — Accounting Initial Review (corporate — Vista access required)
  ├── Workbook reads live data from Vista Production (10.112.11.8)
  ├── Accounting reviews Jobs-Ops tab (read-only at this stage)
  ├── Batch created in LylesWIP on P&P → Vista data snapshot stored
  └── Start sheet → "Ready for Ops: Yes" → batch state advances

STAGE 2 — Operations Review (any location — P&P access only, no Vista required)
  ├── Workbook reads batch data from LylesWIP on P&P (no Vista connection)
  ├── Ops edits yellow override columns (revenue, cost, completion date, notes)
  ├── Double-click col H (Op's Done) → row saved to LylesWIP per job
  └── Start sheet → "Ops Final Approval: Yes" → batch state advances

STAGE 3 — Accounting Final Approval (corporate)
  ├── Workbook reads from LylesWIP — all Ops edits visible
  ├── Accounting edits GAAP override columns (Jobs-GAAP tab)
  ├── Double-click col I (GAAP Done) → row saved to LylesWIP per job
  ├── Start sheet → "Accounting Final Approval: Yes" → batch locked
  └── [Phase 2] Final values written back to Vista bJCOR / bJCOP
```

### LylesWIP Database Schema (P&P Server)

**WipBatches** — one row per company / month / division; tracks batch lifecycle state

| Column | Type | Purpose |
|--------|------|---------|
| Id | INT IDENTITY | PK |
| JCCo | TINYINT | Vista company number |
| WipMonth | DATE | First of month (e.g. 2025-12-01) |
| Department | VARCHAR(10) | Vista department code |
| BatchState | VARCHAR(20) | Open / ReadyForOps / OpsApproved / AcctApproved |
| CreatedBy | VARCHAR(100) | Windows login |
| CreatedAt | DATETIME | Batch creation timestamp |
| StateChangedBy | VARCHAR(100) | Who last advanced the state |
| StateChangedAt | DATETIME | When state last changed |

**WipJobData** — one row per job per WIP cycle; stores all overrides entered by users

| Column | Type | Purpose |
|--------|------|---------|
| Id | INT IDENTITY | PK |
| JCCo | TINYINT | Vista company number |
| Job | VARCHAR(50) | Vista job number (e.g. "51.1108") |
| WipMonth | DATE | First of month |
| OpsRevOverride | DECIMAL(15,2) | Ops revenue projection (NULL = use Vista calc) |
| OpsRevPlugged | BIT | Flagged as a manual override |
| GAAPRevOverride | DECIMAL(15,2) | GAAP revenue projection |
| GAAPRevPlugged | BIT | Flagged as a manual override |
| OpsCostOverride | DECIMAL(15,2) | Ops estimated cost to complete |
| OpsCostPlugged | BIT | Flagged as a manual override |
| GAAPCostOverride | DECIMAL(15,2) | GAAP estimated cost to complete |
| GAAPCostPlugged | BIT | Flagged as a manual override |
| BonusProfit | DECIMAL(15,2) | Bonus-basis projected profit |
| OpsRevNotes | VARCHAR(500) | Notes from Ops on revenue override |
| GAAPRevNotes | VARCHAR(500) | Notes from Accounting on GAAP revenue |
| OpsCostNotes | VARCHAR(500) | Notes from Ops on cost override |
| GAAPCostNotes | VARCHAR(500) | Notes from Accounting on GAAP cost |
| CompletionDate | DATE | Anticipated completion date |
| IsClosed | BIT | Job flagged for close |
| IsOpsDone | BIT | Ops has marked this job Done |
| IsGAAPDone | BIT | Accounting has marked this job Done |
| UserName | VARCHAR(100) | Last user to touch the row |
| UpdatedAt | DATETIME | Last update timestamp |

**WipYearEndSnapshot** — archived at December AcctApproved; source for Prior Year columns next year

**WipMonthlySnapshot** — archived at each AcctApproved; used for 6-month trend (deferred feature)

### Stored Procedures

| Procedure | Purpose |
|-----------|---------|
| LylesWIPCreateBatch | INSERT into WipBatches; return batch ID |
| LylesWIPGetBatches | SELECT batches for Co/Month — used on open to check if batch exists |
| LylesWIPCheckBatchState | Return current state for Co/Month/Dept |
| LylesWIPUpdateBatchState | Advance state + record who changed it and when |
| LylesWIPSaveJobRow | MERGE into WipJobData — called on double-click Done |
| LylesWIPGetJobOverrides | SELECT all override rows for Co/Month — used on data load |
| LylesWIPClearJobData | DELETE WipJobData for Co/Month (batch cancel or clear) |
| LylesWIPSaveYearEndSnapshot | INSERT/UPDATE WipYearEndSnapshot at December cycle close |

### New VBA Module: LylesWIPData.bas

Modeled on the existing VistaData.bas pattern already in the workbook:
- `OpenWIPConnection()` / `CloseWIPConnection()` — connects to PPServerName / LylesWIP
- `CreateBatch(co, month, dept, userName)` — calls LylesWIPCreateBatch
- `GetBatchState(co, month, dept)` — calls LylesWIPCheckBatchState
- `UpdateBatchState(co, month, dept, newState, userName)` — calls LylesWIPUpdateBatchState
- `SaveJobRow(co, job, month, overrideFields...)` — calls LylesWIPSaveJobRow
- `GetJobOverrides(co, month, dept)` — returns recordset for merge on data load

---

## Override Data — The Validation Blocker

Nicole and Cindy confirmed on March 25: **the numbers they validate against are not raw Vista
numbers.** Every month, Nicole's team submits Excel override files (revenue projections, cost
projections, completion dates) that are loaded into the database before the WIP schedule is
produced. Without these overrides, the tool's output will never match what Nicole validates
against.

**40 historical override files have been located on the F: drive** (`F:\Workpapers\LCG - Combined\2025 WIP Automation\{Company}\`)
and confirmed: December 2024 through December 2025, covering all four companies (AIC, APC, NESM, WML).

File structure (confirmed — identical across all companies):
- **Revenue sheet:** Month, Contract (job number), Description, GAAP Override Revenue Projection, Ops Override Revenue Projection, Bonus Profit, Anticipated Completion Date
- **Cost sheet:** Month, Job, Job Desc, GAAP Override Cost Projection, Ops Override Cost Projection

These must be loaded into LylesWIP.WipJobData before validation can begin. This is task #12 below.

**The DJS (Division Job Summary) — what Nicole and Brian use to prepare these files:**
The DJS is a shared Excel file (`Job Summary - Northern Southern Construction Totals`) that Brian
Platten updates manually with Vista cost data each period. It includes contract amounts that may
include pending change orders not yet executed in Vista. The accounting-approved final numbers
from the DJS are extracted into the override Excel files. This is why a job like 51.1108 showed
$94M in the tool (the DJS/override value) while Vista only showed $91M (pending COs not yet posted).
The tool is correct; Vista lags.

---

## What Has Been Completed

| # | Item | Date |
|---|------|------|
| 1 | Full code audit of original tool — 34 VBA modules, full architecture documented | Feb 2026 |
| 2 | All 366 email communications between consultant and accounting team reviewed | Feb 2026 |
| 3 | Architecture confirmed: LylesWIP database on PNP server | Mar 2026 |
| 4 | Vista read path fully rewired — all 4 sheets load live from Vista Production (10.112.11.8) | Mar 5, 2026 |
| 5 | Company and department dropdowns wired to Vista | Mar 5, 2026 |
| 6 | GL closed-month check wired to Vista | Mar 5, 2026 |
| 7 | Two formula bugs fixed: 10% GAAP earned revenue threshold (Col W), Prior Projected Profit (Col R) | Mar 2026 |
| 8 | Write path guarded — workbook no longer crashes attempting to call consultant's private database | Mar 5, 2026 |
| 9 | Live demo with Nicole Leasure and Cindy Jordan — ran Company 15 (WML), December 2025 | Mar 25, 2026 |
| 10 | Five additional data bugs confirmed and documented from demo | Mar 25, 2026 |
| 11 | 40 months of historical override files located on F: drive and structure confirmed with Nicole | Mar 31, 2026 |
| 12 | DJS file structure identified — Job Summary files maintained by Brian Platten on F: drive | Mar 31, 2026 |
| 13 | Full sprint plan defined through delivery | Mar 2026 |

---

## Outstanding Tasks

### Phase A — Vista Query Bug Fixes

| # | Bug | Detail | Owner |
|---|-----|--------|-------|
| A1 | ~~DONE~~ — Batch-month date parameter confirmed by Nicole Leasure | JC Cost & Revenue in Viewpoint uses **Ending Month = batch month** (e.g. 12/25 for December WIP); Beginning Month is left blank. This brings in JTD totals through the batch month end. All 7 Vista CTEs in the workbook query must apply this same cutoff. | Josh |
| A2 | ~~DONE~~ — Date filter: batch-month cutoff applied to all CTEs | `@CurrentDate` removed from 3 locations; all now use `@CutOffDate = EOMONTH(@Month)` | Josh |
| A3 | ~~DONE~~ — Future jobs filter | `c.StartMonth <= @CutOffDate` — removed 36 future 2026 jobs from Dec 2025 run | Josh |
| A4 | ~~DONE~~ — Billing accuracy | New `BilledThruMonth` CTE from `bARTH`; fixed $7.9M overstatement on job 51.1129 | Josh |
| A5 | ~~DONE~~ — Jobs closed in/after batch month | `OR (JobStatus NOT IN (1,3) AND MonthClosed >= @Month)` — adds Dec-closed jobs | Josh |
| A6 | ~~DONE~~ — Closed jobs with billing-only activity | `bARTH` billing EXISTS clause — adds 8 Dept 54 retainage jobs | Josh |
| A7 | ~~DONE~~ — Zero-JTD reversal jobs | Already working — cost EXISTS catches reversals; no change needed | Josh |
| A8 | ~~DONE~~ — Close status reconciliation (SQL side) | `VistaClosedAtWipDate` flag added to SELECT. Full mismatch highlighting deferred to Sprint 4 (needs LylesWIP write path) | Josh |

### Phase B — Database Build (P&P Server)

| # | Task | Detail | Owner |
|---|------|--------|-------|
| B1 | ~~DONE~~ — Create LylesWIP database on P&P server | `10.103.30.11` (Cloud-Apps1). `wip.excel.sql` granted db_owner. DDL: `sql/LylesWIP_CreateDB.sql` | Josh |
| B2 | ~~DONE~~ — Create WipBatches table | State machine: Open → ReadyForOps → OpsApproved → AcctApproved | Josh |
| B3 | ~~DONE~~ — Create WipJobData table | NULL override = use Vista value. Unique on (JCCo, Job, WipMonth) | Josh |
| B4 | ~~DONE~~ — Create WipYearEndSnapshot table | Unique on (JCCo, Job, SnapshotYear) | Josh |
| B5 | ~~DONE~~ — Create all stored procedures | 8 procs: CreateBatch, GetBatches, CheckBatchState, UpdateBatchState, SaveJobRow, GetJobOverrides, ClearJobData, SaveYearEndSnapshot | Josh |
| B6 | ~~DONE~~ — Grant execute permissions to WIP SQL login | `wip.excel.sql` has db_owner on LylesWIP | Josh |
| B7 | ~~DONE~~ — Load 40 historical override files into LylesWIP | 4,974 rows loaded via `sql/load_overrides.py`. All 4 companies, Dec 2024–Dec 2025 | Josh |
| B8 | ~~DONE~~ — Verify loaded override data against Nicole's source files | Job 51.1108 Dec-2025: OpsRev=$94,196,098 matches Nicole's file exactly | Josh |

### Phase C — Workbook Write Path (VBA)

| # | Task | Detail | Owner |
|---|------|--------|-------|
| C1 | Build LylesWIPData.bas | New VBA module: OpenWIPConnection, CloseWIPConnection, CreateBatch, GetBatchState, UpdateBatchState, SaveJobRow, GetJobOverrides. Modeled on VistaData.bas | Josh |
| C2 | Stage 1 — batch snapshot on Load | After Vista data loads, call CreateBatch → INSERT all job rows into WipJobData as the starting snapshot. Stage 1 Accounting review reads from this snapshot | Josh |
| C3 | Override merge on data display | When data loads (Stage 1 reopen or Stage 2/3 open), call GetJobOverrides → merge LylesWIP override values over raw Vista data per job. Override takes priority if present | Josh |
| C4 | Stage 2 — wire UpdateRow | Double-click col H (Op's Done) calls SaveJobRow → writes all override field values for that row to WipJobData in LylesWIP | Josh |
| C5 | Stage 3 — wire GAAP UpdateRow | Double-click col I (GAAP Done) calls SaveJobRow with GAAP override fields | Josh |
| C6 | Wire CreateBatch and UseExistingBatch | On batch open: check LylesWIPGetBatches for existing batch. If found → prompt reopen. If not → show division selector → CreateBatch | Josh |
| C7 | Wire Start sheet state transitions | "Ready for Ops: Yes" → UpdateBatchState to ReadyForOps. "Ops Final Approval: Yes" → OpsApproved. "Accounting Final Approval: Yes" → AcctApproved | Josh |
| C8 | Wire CompleteCheck | Before state advance, verify all jobs in batch have IsOpsDone=1 (or IsGAAPDone=1 for Stage 3) | Josh |
| C9 | Cell locking on state advance | Jobs-Ops yellow cells lock once batch reaches OpsApproved. Jobs-GAAP yellow cells lock once AcctApproved | Josh |
| C10 | Wire role-based permissions | Replace hardcoded WIPAccounting with call to pnp.WIPSECGetRole. Show/hide UI elements per role | Josh |
| C11 | Phase 2: Vista write-back | At AcctApproved, write final GAAP override values to Vista bJCOR / bJCOP tables. Deferred — not blocking delivery | Josh |

### Phase D — Validation & Delivery

| # | Task | Detail | Owner |
|---|------|--------|-------|
| D1 | ~~Confirm Col R carry-forward rule with Nicole~~ **RESOLVED — Apr 1, 2026** | **Determination**: Reviewed recording of March 17, 2026 Column Mapping Exercise meeting. Nicole stated directly (~53:47): *"Prior Projected Profit — but what it is in reality it's the prior month schedule column P"*, confirmed by attendees as a manual copy performed before starting each month's cycle. Root cause of Nicole's Col R discrepancy identified: consultant's tool sourced this value from Vista `bJCOR` (annual March plug), which diverges from Nicole's manual copy whenever projections change mid-year. **Implementation rule**: At AcctApproved, snapshot Col P per job into the persistence table. On next month's batch load, restore those values as Col R. WipYearEndSnapshot is not needed for this purpose. | Josh |
| D2 | End-to-end validation — December 2025, WML | With overrides loaded: run Dec 2025 WML, confirm numbers match Nicole's manual WIP Schedule | Josh + Nicole |
| D3 | Validation — AIC, APC, NESM | Repeat validation for remaining three companies | Josh + Nicole |
| D4 | SQL driver install — Brian Platten | Brian needs ODBC SQL driver installed to connect to P&P from his machine | Josh |
| D5 | SQL driver install — Harbir Atwal | Same setup required | Josh |
| D6 | User setup instructions | Document: trusted folder path (C:\Users\[user]\trusted), SQL driver install, workbook first-open steps | Josh |
| D7 | Security cleanup | Re-hide Settings sheet (xlSheetVeryHidden), clear Settings C33 test username, verify no hardcoded test credentials | Josh |
| D8 | Final sign-off | Nicole and Cindy confirm numbers tie and complete three-stage workflow end-to-end | Nicole + Cindy |

---

## Milestones

| Milestone | Target | Status |
|-----------|--------|--------|
| Vista read path complete — all 4 sheets load live from production | Mar 5, 2026 | DONE |
| Live demo with Nicole + Cindy — bugs documented | Mar 25, 2026 | DONE |
| Override files located on F: drive; structure confirmed with Nicole | Mar 31, 2026 | DONE |
| Date parameter confirmed by Nicole — Ending Month = batch month; Beginning Month = blank | Mar 31, 2026 | Complete |
| All Vista query bugs fixed (Phase A) | Apr 11, 2026 | |
| LylesWIP database live on PNP — tables + stored procs | Apr 14, 2026 | |
| 40 historical override files loaded into LylesWIP | Apr 17, 2026 | |
| Write path working — Stage 1 snapshot + Stage 2 double-click Done saves | Apr 24, 2026 | |
| Role-based permissions live | Apr 28, 2026 | |
| Three-stage workflow complete end-to-end | May 1, 2026 | |
| Nicole validation session — Dec 2025 numbers tie | May 5, 2026 | |
| Final fixes from validation | May 7, 2026 | |
| **Delivery — Settings hidden, users set up, signed off** | **May 8, 2026** | |
| Phase 2: Vista bJCOR / bJCOP write-back | TBD post-delivery | |

---

## Meeting Cadence

- **Standing meetings:** Twice weekly (days TBD)
- **All other IT projects on hold** until Auto-WIP is delivered

---

## Key Contacts

| Name | Role | Notes |
|------|------|-------|
| Kevin Shigematsu | CEO — project sponsor | |
| Cindy Jordan | CFO — final approver | |
| Nicole Leasure | VP Corporate Controller — primary validator | |
| Josh Garrison | Director of Technology Innovation — developer | |
| Dane Wildey | CIO | |
| Brian Platten | Controller — maintains DJS; needs SQL driver installed | Action: A1 Crystal Report SQL |
| Harbir Atwal | Controller — needs SQL driver installed | |
| Michael Roberts | Original consultant | Optional knowledge transfer; not blocking delivery |
