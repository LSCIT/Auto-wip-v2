# Auto-WIP Schedule — Project Status Report
**Date:** March 31, 2026
**Prepared by:** Josh Garrison, IT Development
**Distribution:** Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey
**Priority:** HIGHEST — Owner / Board / Senior Management

---

## Executive Summary

The Automated WIP Schedule has been under development with an outside consultant (RCS Plan)
for several years. The decision has been made to bring active development in-house to ensure
timely delivery. The tool is a macro-enabled Excel workbook used monthly by Accounting and
Operations to produce the Work-in-Progress schedule across four companies (AIC, APC, NESM, WML).
The CEO has committed to this platform to the board and ownership.

An independent technical review of the consultant's last delivered version (Rev 5.47p) rates it
at **70–75% production ready — not the 98% that was previously communicated.** The Ops tab and
workflow infrastructure are largely solid. The GAAP tab — the output that drives financial
reporting — has foundational calculation errors that cascade through seven columns. That is not
a cosmetic issue. That is the primary product.

Our in-house work (Rev 5.52p) has fixed the two root-cause formula errors behind those seven
column discrepancies. The Vista data read path is fully operational. However, the save and
approval layer — the mechanism that allows user overrides to be recorded, the three-stage
workflow to run, and approved values to eventually be written back to Vista — does not yet
exist. It is being built.

**Honest current production readiness: approximately 35–40% for the complete intended system.**
The foundational rebuild is solid. The remaining work is concentrated, well-defined, and
actively in progress.

**Estimated delivery: approximately six weeks from today (targeting mid-May 2026).**

---

## Consultant's Last Version — Independent Assessment

An independent review of Rev 5.47p (the consultant's most recent version) against the LCG
Automated WIP Process document and Nicole's reported discrepancies produced the following
findings.

### What Was Working in 5.47p

| Area | Status |
|------|--------|
| Three-phase approval workflow (Accounting → Ops → Accounting Final) | Functional |
| Database connectivity | Functional |
| Batch creation, clearing, and approval triggers | Functional |
| Ops tab profit recognition (30% threshold, loss recognition) | Functional |
| Ops tab derived calculations (CY Revenue, CY Cost, Backlog, etc.) | Functional |
| Double-click Done / Close / Notes UI interactions | Functional |
| Company / Division F4 lookup | Functional |

### What Was Not Working in 5.47p — Nicole's Discrepancies

**Issue 1 — Bonus Tab: Job 54.9033 shows Closed but is Open in Vista**

The workbook has two separate "closed" concepts that are never reconciled: a manual Close
flag set by Ops within the workbook, and the Contract Status imported from Vista. A job can
be flagged closed internally while Vista still shows Open. The system trusts whichever was
set without validating against the other.

**Issue 2 — Bonus Tab: Jobs 56.1022 and 56.1057 missing despite needing to show reversals**

Jobs with $0 JTD totals are filtered out of display. This is correct for most jobs but wrong
when a reversal was posted in the current month — the month-to-date change is real and must
be visible even if JTD nets to zero.

**Issue 3 — GAAP Column R (Prior Projected Profit): showing $0 instead of prior WIP plug**

Column R was using a formula tied to the 10% completion threshold instead of carrying forward
the prior period's approved WIP value. Example: Job 51.1151 showed $0 but should have shown
$596,848 from March's WIP.

**Issue 4 — GAAP Column W (JTD Earned Revenue): core calculation error cascading to seven columns**

For jobs under 10% completion, GAAP requires JTD Earned Revenue to equal JTD Cost. The
formula was not implementing this correctly. Because every downstream GAAP column derives from
Column W, this single error propagated into Columns AD, AF, AH, AJ, AL, AQ, and AR.

Nicole identified errors in all seven. They are not seven separate bugs — they are one
formula error in Column W expressing itself seven ways.

### Independent Assessment Verdict

> *"The system is not ready for unattended production use on the GAAP side. A realistic
> completion estimate: 70–75% overall. The Ops side is ~90%. The GAAP side is ~55–60%."*

---

## What Our In-House Rebuild Has Addressed

### Nicole's Discrepancies — Current Status

| Nicole's Issue | 5.47p | 5.52p (Our Version) |
|----------------|-------|---------------------|
| Job 54.9033 closed in workbook but Open in Vista | Bug present | Not yet fixed — requires close/status reconciliation logic |
| Jobs 56.1022 / 56.1057 missing (zero JTD with reversals) | Bug present | Not yet fixed — requires MTD activity detection in Vista query |
| Column R (Prior Projected Profit) showing $0 | Bug present | Formula patched — full solution requires prior month data from database |
| Column W (JTD Earned Revenue) wrong for <10% jobs | Bug present | **Fixed** — formula now correctly uses JTD Cost when job is under 10% complete |
| Columns AD, AF, AH, AJ, AL, AQ, AR — cascade errors | All wrong | Expected to resolve as result of Column W fix — not yet verified against production data with overrides loaded |

### Additional Issues Found During the March 25 Demo

Five additional data accuracy bugs were confirmed live during the demo with Nicole and Cindy.
All five originate from the Vista read query and must be resolved before validation can succeed.

| Bug | Root Cause | Status |
|-----|-----------|--------|
| Jobs set up after batch month appear on historical runs | Query has no setup-date filter | Not fixed |
| Cost and billing pull to today's date, not batch month end | Date cutoff not applied consistently across all CTEs | Not fixed |
| Billing amounts not accurate | Related to date filter; needs investigation | Not fixed |
| Open jobs with zero current-year activity dropping off | Inclusion rule not implemented | Not fixed |
| Closed jobs with warranty or late-posted costs not reappearing | Same inclusion rule missing | Not fixed |

The correct date filter has been confirmed: the Vista JC Cost & Revenue Crystal Report that
Nicole, Brian, and the division controllers use to prepare the Division Job Summary (DJS)
applies an **Ending Month** parameter equal to the batch month. All Vista queries in the
workbook must match this cutoff.

**Critical context on override data:** Nicole and Cindy confirmed that the numbers they
validate against are not raw Vista numbers. Every month, Nicole's team submits Excel override
files (revenue and cost projections) that override raw Vista calculations. These 40 historical
files have been located on the F: drive and inventoried. Without loading these overrides, the
tool will never match Nicole's validation numbers — raw Vista data alone is insufficient.

### Infrastructure Work Completed

| Item | Status |
|------|--------|
| Vista read path completely rebuilt — all four sheets load live from Vista Production (10.112.11.8) | Complete |
| Company and department dropdowns rebuilt against Vista | Complete |
| GL period close check rebuilt against Vista | Complete |
| Workbook stabilized — zero dependency on consultant's private database | Complete |
| Formula fixes: Column W (earned revenue threshold) and Column R (prior projected profit) | Complete |
| Full code audit of consultant's 34-module codebase | Complete |
| All 40 historical override files (Dec 2024 – Dec 2025) located on F: drive and inventoried | Complete |
| Override file structure confirmed with Nicole — Revenue sheet (GAAP + Ops + Bonus) and Cost sheet | Complete |

---

## What Is Not Yet Built

### The Write Path — Entirely Absent

The original consultant's tool had a partial write path connecting to his private WipDb
database. Our rebuild deliberately removed that dependency and replaced it with guarded stubs
(the workbook will not crash, but nothing saves) while we build the correct replacement.
The write path is the most significant remaining body of work.

**When a user double-clicks Done today, nothing is saved. The three-stage workflow has no
persistence. This is by design during the rebuild — it will be built.**

| Write Path Component | Status |
|---------------------|--------|
| LylesWIP database on P&P server — working layer for all users | Not built |
| Stage 1: snapshot Vista batch data into LylesWIP on batch open | Not built |
| Stage 2: load 40 historical override files into LylesWIP | Not loaded |
| Stage 2: save override values per job (double-click Done → LylesWIP) | Not built |
| Stage 2: load existing overrides from LylesWIP when batch reopens | Not built |
| Three-stage workflow state machine (Open → ReadyForOps → OpsApproved → AcctApproved) | Not built |
| Phase 2: final write-back of GAAP projections to Vista bJCOR / bJCOP | Deferred |

### Architecture: Why LylesWIP on P&P Server, and How It Works

Remote offices and job-site trailers cannot reach Vista's SQL Server directly, but all
locations can reach the P&P server. The monthly cycle works as follows:

```
STAGE 1 — Accounting (corporate, Vista access required)
    Reads live Vista data → snapshots batch into LylesWIP on P&P server
    Marks batch "Ready for Ops"

STAGE 2 — Operations (any location, P&P access only)
    Connects to LylesWIP on P&P server — no Vista connection required
    Reviews and edits override columns
    Double-click Done → saves per-job row to LylesWIP
    Marks batch "Ops Approved"

STAGE 3 — Accounting (corporate)
    Reads from LylesWIP — sees all Ops edits
    Reviews GAAP columns, makes accounting adjustments
    Marks "Accounting Final Approval" → batch locked
    [Phase 2] Final values written back to Vista bJCOR / bJCOP
```

This design means Ops users — superintendents, PMs, and controllers at remote sites and job
trailers — never need a Vista connection. They operate entirely against LylesWIP on P&P,
which is already accessible to them for other P&P workflows.

### Permissions

Role-based permissions are currently hardcoded to the Accounting role for all users. The
P&P server already has a stored procedure (pnp.WIPSECGetRole) that maps Windows logins to
WIP roles. This will be wired before distribution.

---

## Honest Production Readiness Assessment

| Layer | Readiness | Notes |
|-------|-----------|-------|
| Vista read path — data loads | ~65% | Loads correctly; five date/filter bugs remain |
| GAAP formula accuracy | ~70% | Col W and R fixed; cascade columns not yet verified with real override data |
| Bonus tab job inclusion | ~60% | Main jobs present; close reconciliation and zero-JTD reversal logic missing |
| Write path — saving overrides | 0% | Not built; intentionally guarded during rebuild |
| Three-stage workflow | 15% | UI exists; nothing wired to persistence |
| Historical override data | 0% | 40 files located and inventoried; not yet loaded into LylesWIP |
| Permissions | 10% | Hardcoded; P&P proc exists and ready to wire |
| **Overall system** | **~35–40%** | Foundational rebuild complete; feature layer in progress |

**The gap between 98% (what was communicated) and 35–40% (actual) is almost entirely
explained by the write path never having been built.** Our version removes the broken
dependency and replaces it with a correct implementation — but that implementation is
not yet complete.

---

## Task List

### Immediate — Vista Query Fixes

| # | Task | Owner |
|---|------|-------|
| 1 | ~~DONE~~ — Confirmed by Nicole Leasure: JC Cost & Revenue uses **Ending Month = batch month** (e.g. 12/25 for December WIP); Beginning Month is left blank. All Vista CTEs must cut off at batch month-end using this same pattern. | Josh |
| 2 | Apply Ending Month (batch month) as hard cutoff on all cost and billing sub-queries | Josh |
| 3 | Fix future jobs: exclude jobs with Vista setup date after batch month | Josh |
| 4 | Investigate and fix billing accuracy — likely resolves with date fix | Josh |
| 5 | Fix job inclusion: include open jobs with zero current-year activity | Josh |
| 6 | Fix job inclusion: include closed jobs with new current-year cost (warranty / late posts) | Josh |
| 7 | Fix zero-JTD reversal jobs: detect current-period activity even when JTD = $0 | Josh |
| 8 | Fix close status reconciliation: compare workbook Close flag against Vista Contract Status | Josh |

### Database Build (P&P Server — LylesWIP)

| # | Task | Owner |
|---|------|-------|
| 9 | Create LylesWIP database on P&P server | Josh |
| 10 | Create tables: WipBatches, WipJobSnapshot, WipJobData (overrides per job per cycle) | Josh |
| 11 | Create stored procedures: CreateBatch, SaveJobRow, GetJobOverrides, UpdateBatchState, CheckBatchState | Josh |
| 12 | Load 40 historical override files from F: drive into LylesWIP (Dec 2024–Dec 2025, all 4 companies) | Josh |

### Workbook — Write Path (VBA)

| # | Task | Owner |
|---|------|-------|
| 13 | Build LylesWIPData.bas — new VBA module for P&P / LylesWIP connection | Josh |
| 14 | Stage 1: Wire batch creation — on Load Data, snapshot Vista results into LylesWIP | Josh |
| 15 | Stage 2: Wire override merge on re-open — pull existing LylesWIP overrides and apply over raw Vista data | Josh |
| 16 | Stage 2: Wire UpdateRow — save per-job override values to LylesWIP on double-click Done | Josh |
| 17 | Stage 3: Wire three-stage workflow state machine via Start sheet radio buttons | Josh |
| 18 | Wire role-based permissions to pnp.WIPSECGetRole | Josh |
| 19 | Phase 2 (post-delivery): Wire final Vista write-back — GAAP projections to bJCOR / bJCOP | Josh |

### Validation & Delivery

| # | Task | Owner |
|---|------|-------|
| 20 | Confirm Prior Projected Profit (Col R) carry-forward rule with Nicole | Josh + Nicole |
| 21 | End-to-end validation: December 2025, WML Company 15, Nicole confirms numbers tie | Josh + Nicole |
| 22 | Expand validation to AIC, APC, NESM | Josh + Nicole |
| 23 | SQL driver installation for Brian Platten and Harbir Atwal | Josh |
| 24 | User setup instructions (trusted folder path, SQL driver, initial setup steps) | Josh |
| 25 | Re-hide Settings sheet, clear test credentials before any distribution | Josh |
| 26 | Final sign-off: Nicole and Cindy confirm numbers and complete three-stage workflow | Nicole + Cindy |

---

## Milestone Schedule

| Milestone | Target | Status |
|-----------|--------|--------|
| Vista read path rebuilt — all four sheets load from production | Mar 5, 2026 | Complete |
| Formula fixes: Column W and Column R | Mar 2026 | Complete |
| Live demo with Nicole + Cindy — five additional bugs documented | Mar 25, 2026 | Complete |
| Historical override files located and structure confirmed with Nicole | Mar 31, 2026 | Complete |
| Date parameter confirmed by Nicole — Ending Month = batch month; Beginning Month = blank | Mar 31, 2026 | Complete |
| All Vista query bugs fixed | Apr 11, 2026 | |
| LylesWIP database live on P&P, stored procedures complete | Apr 14, 2026 | |
| 40 historical override files loaded into LylesWIP | Apr 17, 2026 | |
| Write path working — overrides save, batches track | Apr 24, 2026 | |
| Role-based permissions live | Apr 28, 2026 | |
| Three-stage workflow complete end-to-end | May 1, 2026 | |
| Nicole validation session — December 2025 numbers confirmed | May 5, 2026 | |
| Final fixes from validation | May 7, 2026 | |
| **Delivery — signed off, users set up, distributed** | **May 8, 2026** | |

---

## Meeting Cadence

- **Standing meetings:** Twice weekly (days to be confirmed)
- All other IT projects on hold until Auto-WIP is delivered

---

## Key Contacts

| Name | Role |
|------|------|
| Kevin Shigematsu | CEO, Lyles Services Co. |
| Cindy Jordan | CFO, Lyles Services Co. |
| Nicole Leasure | VP Corporate Controller |
| Josh Garrison | Director of Technology Innovation |
| Dane Wildey | CIO |
| Brian Platten | Controller — needs SQL driver installed |
| Harbir Atwal | Controller — needs SQL driver installed |
