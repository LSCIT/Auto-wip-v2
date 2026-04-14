# Auto-WIP Schedule — Project Status Report
**Date:** April 6, 2026
**Prepared by:** Josh Garrison, Director of Technology Innovation
**Distribution:** Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey
**Priority:** HIGHEST — Owner / Board / Senior Management escalation

---

## Executive Summary

April 6 began with a demo session with Nicole Leasure and Cindy Jordan, validating the full 3-stage workflow on Rev 5.68p. The demo confirmed multiple values are correct (contract amounts, change in anticipated profit, override projections, workflow state machine) and identified five data accuracy issues — all traced to the same root cause: date filtering in the Vista query was using transaction posted dates instead of fiscal months, causing cost and billing values to diverge from the Crystal Report "JC Cost and Revenue" that Nicole validates against.

All five issues were investigated, root-caused, and fixed the same day. The corrected query was verified against the Crystal Report with **exact penny-match** on both JTD Cost ($87,315,159.22) and Billed to Date ($96,918,206.90) for the benchmark job 51.1129. These fixes, along with several other improvements, were deployed as Rev 5.70p.

> **Production readiness: approximately 88-92%.** JTD Cost and Billed to Date now match Crystal Report exactly. Prior month profit, job inclusion, and Vista write-back copy tables all addressed. Remaining: circular reference edge case (pre-existing), fine-tuning prior profit denominator, multi-company validation.
> **Target delivery: May 8, 2026 — no change.**

---

## Demo Session — April 6, 2026

**Attendees:** Josh Garrison, Nicole Leasure, Cindy Jordan
**Workbook:** WIPSchedule - Rev 5.68p, Company 15 (WML), Division 51, December 2025
**Duration:** ~36 minutes (recorded with transcript)

### What Validated Successfully

| Item | Detail |
|------|--------|
| Contract amounts | Confirmed correct across multiple jobs |
| Change in anticipated profit | All values matched Nicole's WIP |
| Override values (GAAP projections) | 51.1129 contract amount 126,483,581 confirmed |
| Ready for OPS button | Permission check working — blocked unauthorized roles |
| GAAP Done (double-click save) | Saved to database with user name attached |
| Audit trail comments | Override changes recorded with from/to/who/when |
| Closed jobs filtering | Improved — fewer/correct jobs in closed section |

### Issues Identified

| # | Issue | Root Cause | Status |
|---|-------|-----------|--------|
| 1 | JTD Cost pulling beyond batch month | Query used PostedDate instead of Mth (fiscal month) | **Fixed — exact match** |
| 2 | Billed to Date pulling beyond batch month | Query used bARTH (AR side) instead of JB Progress Bills | **Fixed — exact match** |
| 3 | Circular reference on job 51.1158 | Pre-existing formula issue in Michael's workbook (zero-profit edge case) | **Confirmed not our bug** |
| 4 | Prior month profit showing projected instead of recognized | MergePriorMonthProfits wrote OpsRev-OpsCost instead of earned rev - JTD cost | **Fixed** |
| 5 | Job 54.9416 not appearing (StartMonth = Jan 2026) | Query excluded jobs by StartMonth only | **Fixed — cost-exists inclusion** |
| 6 | Completion Date source | Should come from Nicole's override files | **Already working** |

### Nicole's New Business Rule — Job Inclusion

Nicole clarified that job inclusion should be driven by **cost activity first**, not just contract StartMonth:
1. If any cost has hit the job through the batch month → include it
2. If no cost, fall back to StartMonth ≤ batch month (for backlog)

This resolves job 54.9416, which had preliminary design costs in Nov/Dec 2025 but an official StartMonth of January 2026.

---

## What Changed Since April 3, 2026

| Item | Status Apr 3 | Status Today (Apr 6) |
|------|-------------|----------------------|
| JTD Cost accuracy | Off by ~$36K vs Crystal Report (PostedDate filter) | **Exact match** (Mth filter) |
| Billed to Date accuracy | Off by ~$236K vs Crystal Report (bARTH source) | **Exact match** (JB Progress Bills) |
| Job inclusion (54.9416) | Excluded (StartMonth filter only) | **Included** (cost-exists fallback) |
| Prior month profit | Showed projected ($11.25M) | Shows recognized (~$8.3M) |
| LylesWIP override data | GAAP values missing, zeros stored as NULL | **All 4,973 rows reloaded, 0 mismatches** |
| DB state for demo | Ops Done flags set from testing | **Reset — clean for Nicole/Cindy** |
| Vista write-back tables | Not built | **WipJCOP + WipJCOR created, proc tested** |
| Push to Vista button | Not built | **Built (test mode — guards only, no write)** |
| Workbook versions | 5.68p only | **5.68p (demo), 5.70p (fixes), source separated** |
| Formula comparison | Not verified | **22/22 Ops match, 21/22 GAAP match (1 intentional fix)** |

---

## Data Integrity Fix — April 6, 2026

Before the demo, identified and fixed a data loading issue:

- **`to_decimal()` in load_overrides.py** — Nicole's explicit $0 overrides were stored as NULL instead of 0. Fixed: zeros now stored as 0 with Plugged=1.
- **GAAP override values** — 541 GAAP overrides were missing from December 2025. Root cause: values were wiped during April 3 testing. Fixed: all 40 files reloaded (4,973 rows).
- **Test artifacts** — Reset Div 51 batch state, cleared 28 workflow flags, restored 51.1108 override values to Nicole's originals.
- **Verification:** 577 Dec 2025 rows compared against Nicole's source files — **0 mismatches**.

---

## Technical Detail — Query Fixes

### Fix 1: Fiscal Month Filtering (JTD Cost)

The Crystal Report "JC Cost and Revenue" filters by `Mth` (fiscal month), not `PostedDate` or `ActualDate`. A transaction entered in January 2026 can have a fiscal month of December 2025 (or vice versa). Switching all bJCCD date filters from PostedDate/ActualDate to Mth produces an exact match.

| Source | JTD Cost for 51.1129 |
|--------|---------------------|
| Crystal Report (Ending Month 12/25) | $87,315,159.22 |
| Our query (Mth <= @Month) | $87,315,159.22 |
| Old query (PostedDate <= @CutOffDate) | $87,279,469 (off by ~$36K) |

### Fix 2: JB Progress Bills (Billed to Date)

The Crystal Report sources "Billed Amount" from the JB (Job Billing) module via `vrvJBProgressBills`, not from AR (Accounts Receivable) via `bARTH`. AR tracks invoices/payments/retainage separately; JB tracks cumulative billed per progress bill.

| Source | Billed Amount for 51.1129 |
|--------|--------------------------|
| Crystal Report | $96,918,206.90 |
| vrvJBProgressBills (JB side) | $96,918,206.90 |
| bARTH Invoiced (AR side) | $96,682,005.02 (off by ~$236K) |

### Fix 3: Cost-Exists Job Inclusion

Added a secondary inclusion path: if a job has actual cost activity (bJCCD) through the batch month, include it even if StartMonth is after the batch month. Adds 2 jobs for Dec 2025: 54.9416 (Div 54) and 57.0009 (Div 57).

### Fix 4: Recognized Profit (Prior Month)

Changed `MergePriorMonthProfitsOntoSheet` to compute recognized profit (earned revenue - JTD actual cost) instead of projected profit (override rev - override cost). Added `BuildPriorMonthCostLookup` function that queries Vista for prior month JTD costs.

### Circular Reference (51.1158) — Not Our Bug

Formula comparison between Michael's 5.47p and our 5.70p confirmed: **zero formula differences on Jobs-Ops** (22/22 match), **one intentional fix on Jobs-GAAP** (Col Z — removed erroneous 10% threshold, approved by Nicole in Sprint 1). The circular reference is pre-existing and occurs when OpsRev = OpsCost (zero expected profit).

---

## Current Workbook Versions

| Version | Purpose | Status |
|---------|---------|--------|
| **5.68p** | Nicole/Cindy demo version | Frozen — validated, DB reset for testing |
| **5.70p** | Query fixes + Push to Vista button | Built — Crystal Report match verified |

Source code separated: `vba_source/` (5.68p frozen) and `vba_source_569p/` (5.70p development).

---

## Vista Write-Back Infrastructure (Phase 2 Prep)

Built local copies of Vista override tables in LylesWIP for safe testing:

| Table | Mirrors | Purpose |
|-------|---------|---------|
| WipJCOP | Vista bJCOP | GAAP/OPS cost overrides |
| WipJCOR | Vista bJCOR | GAAP/OPS revenue overrides |

Stored procedure `LylesWIPWriteBackToVista` created with:
- Two-pass MERGE pattern (matches Michael's LCGWIPMergeDetail)
- Quarterly guard (Mar/Jun/Sep/Dec only)
- AcctApproved guard (all departments must be approved)
- Smoke-tested: 51.1108 values match Nicole's source exactly

"Push to Vista" button on Start sheet — currently in test mode (validates all guards, shows confirmation, does not execute write).

---

## Remaining for Delivery

| # | Item | Priority | Status |
|---|------|----------|--------|
| R1 | Nicole/Cindy validation on 5.70p | Critical | Ready — fixes address all demo findings |
| R2 | Multi-company testing (AIC, APC, NESM) | High | Not started |
| R3 | Wire permissions to pnp.WIPSECGetRole | Medium | Not started |
| R4 | SQL driver install for Brian Platten + Harbir Atwal | Medium | Not started |
| R5 | Security cleanup (Settings sheet, test credentials) | High | Not started |
| R6 | Circular reference on 51.1158 (pre-existing) | Low | Pre-existing — won't block delivery |
| R7 | Fine-tune prior profit denominator (~$312K gap) | Medium | Close — may resolve with full data reload |
| R8 | Confirm: decimal display preference (cents vs whole dollars) | Low | Question for Nicole |
| R9 | Confirm: GAAP Billed to Date source (same Crystal Report?) | Low | Question for Nicole |

---

## Milestone Schedule

| Milestone | Target | Status |
|-----------|--------|--------|
| Vista read path — all 4 sheets | Mar 28 | Done |
| Vista query bugs fixed (A1–A8) | Apr 11 | Done Mar 31 |
| LylesWIP database + stored procs | Apr 14 | Done Mar 31 |
| Override load (40 files, 4,974 rows) | Apr 17 | Done Mar 31 |
| Validation pipeline (24/24 PASS) | Apr 4 | Done Apr 3 |
| Write path + 3-stage workflow | Apr 14 | Done Apr 3 |
| Data integrity fix + full reload | Apr 6 | Done Apr 6 |
| **Query fixes — Crystal Report exact match** | **Apr 7** | **Done Apr 6** |
| **Vista write-back copy tables** | **Apr 7** | **Done Apr 6** |
| Nicole/Cindy validation (5.70p) | Apr 7-11 | Ready to schedule |
| Multi-company testing | Apr 14-18 | Not started |
| Permissions wire (P&P proc) | Apr 21-25 | Not started |
| Security cleanup | Apr 28 - May 2 | Not started |
| **Production delivery** | **May 8** | **On track** |

---

## Key Contacts

| Name | Role | Relevance |
|------|------|-----------|
| Kevin Shigematsu | CEO / Project Sponsor | Board commitment. Final escalation. |
| Cindy Jordan | CFO | Final sign-off on WIP schedule accuracy. |
| Nicole Leasure | VP Corporate Controller | Primary user and validator. Source of truth for override data. |
| Dane Wildey | CIO | IT oversight, infrastructure. |
| Brian Platten | Controller | Reviewer. Needs SQL driver install. |
| Harbir Atwal | Controller | Reviewer. Needs SQL driver install. |
| Josh Garrison | Dir. of Technology Innovation | Developer. Report author. |

---

*CONFIDENTIAL — Lyles Services Co. internal use only*
