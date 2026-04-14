# Auto-WIP Schedule — Project Status Report
**Date:** April 7, 2026
**Prepared by:** Josh Garrison, Director of Technology Innovation
**Distribution:** Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey
**Priority:** HIGHEST — Owner / Board / Senior Management escalation

---

## Executive Summary

Received first round of user feedback from Nicole Leasure on Rev 5.70p. Nicole identified three data discrepancies on the Jobs-Ops tab for job 51.1129, all traced to the same root cause: VBA code was overwriting sheet formula cells with computed values, bypassing the original template's architecture where visible columns read from hidden "Z-columns" populated by LylesWIP. Two column label changes were also requested. All five items were investigated, root-caused, and fixed in Rev 5.71p the same day.

> **Production readiness: approximately 90-93%.** All Nicole feedback items from 5.70p addressed. Label changes applied. VBA modules imported and workbook ready for re-test.
> **Target delivery: May 8, 2026 — no change.**

---

## Nicole's Feedback — April 7, 2026

**Source:** Email from Nicole Leasure, reviewing Rev 5.70p
**Job referenced:** 51.1129, Company 15 (WML), Division 51

### Issues Reported and Resolved

| # | Nicole's Report | Column | Root Cause | Fix |
|---|----------------|--------|------------|-----|
| 1 | AC should be $8,997,122 — "seems to be pulling from GAAP schedule" | COLJTDPriorProfit (AC) | `MergePriorMonthProfitsOntoSheet` was overwriting the AC formula with a computed "recognized profit" value. The template formula `=BK6` (hidden Z-column COLZPriorBonusProfit) was correct — it reads the prior month's BonusProfit from LylesWIP. Our function was clobbering it. | Rewrote function to populate the Z-column (BK) with the prior month's BonusProfit from LylesWIP, letting the sheet formula handle the visible column. |
| 2 | AG should be $71,707,694 — "Michael's spreadsheet pulls correct numbers for 06/01/25" | COLAPYRev (AG) | Prior Year Revenue formula on Ops = PYCost + PriorYrBonusProfit, but the Vista query stubs PriorYrBonusProfit to 0 (Vista doesn't store WIP history). The bonus backfill function (`MergePriorYearBonusOntoSheet`) wrote the bonus to the AJ column but never corrected AG. | Expanded `MergePriorYearBonusOntoSheet` to also update AG = PYCost (from AH) + bonus from LylesWIP. |
| 3 | AI should be $6,957,477 | COLAPYCalcProfit (AI) | Template has a sheet formula `=AG6-AH6` (revenue - cost), but VBA was overwriting it with `PYEarnedRev - PYCost` using GAAP-based values. | Removed the VBA overwrite. The formula auto-calculates correctly once AG is fixed. |
| 4 | Rename column AD on Jobs-Ops to "MTD Change in Profit" | COLJTDChgProfit (AD) | Label change request | Changed header from "JTD Change In Profit" to "MTD Change in Profit" |
| 5 | Rename column AA on Jobs-GAAP to "MTD Change in Profit" | COLJTDChgProfit (AA) | Label change request | Changed header from "JTD Change In Profit" to "MTD Change in Profit" |

### Root Cause Pattern

All three data issues share the same architectural root cause: **VBA code was writing directly to visible formula columns, destroying the template formulas that read from hidden Z-columns.**

Michael's original template architecture uses a two-layer pattern:
- **Hidden Z-columns** (BK, BN, etc.) are populated by VBA with data from LylesWIP
- **Visible columns** (AC, Q, AI, etc.) contain formulas that read from the Z-columns

The Vista direct-read query (added in Sprint 1) stubs all LylesWIP-sourced fields to 0. The merge functions (`MergePriorMonthProfitsOntoSheet`, `MergePriorYearBonusOntoSheet`) were created to backfill these from LylesWIP after the Vista data loads. The bug was that these functions wrote to the visible columns instead of (or in addition to) the Z-columns, overwriting the formulas.

---

## What Changed Since April 6, 2026

| Item | Status Apr 6 | Status Today (Apr 7) |
|------|-------------|----------------------|
| Nicole/Cindy feedback | Email sent, waiting | **Received — 5 items, all resolved** |
| COLJTDPriorProfit (AC) | Showed computed recognized profit | **Shows prior month BonusProfit via Z-column** |
| COLAPYRev (AG) | Missing bonus (Vista stub = 0) | **PYCost + bonus from LylesWIP year-end** |
| COLAPYCalcProfit (AI) | VBA-overwritten with GAAP values | **Sheet formula =AG-AH restored** |
| Column labels (AD, AA) | "JTD Change In Profit" | **"MTD Change in Profit"** |
| Workbook version | 5.70p | **5.71p (all fixes applied, modules imported)** |

---

## Technical Detail — Fixes Applied

### Fix 1: MergePriorMonthProfitsOntoSheet (LylesWIPData.bas)

**Before:** Computed `recognizedProfit = earnedRev - priorJTDCost` using a Vista JTD cost query, then wrote it directly to COLJTDPriorProfit (AC), destroying the formula.

**After:** Looks up prior month's BonusProfit (`ov(8)`) and OpsRev - OpsCost from LylesWIP, writes them to the hidden Z-columns:
- COLZPriorBonusProfit (BK) — feeds AC via formula `=BK`
- COLZPriorJTDOPsProfit (BN) — feeds Q via formula `=BN`

The `BuildPriorMonthCostLookup` Vista query is no longer called (dead code — can be removed in cleanup).

### Fix 2: MergePriorYearBonusOntoSheet (LylesWIPData.bas)

**Before:** Wrote prior year-end BonusProfit to COLAPYBonusProfit (AJ) only.

**After:** Also reads PYCost from COLAPYCost (AH, already written by GetWipDetail2) and updates COLAPYRev (AG) = PYCost + bonus. This corrects the baseline that GetWipDetail2 wrote with a 0 bonus stub.

### Fix 3: GetWIPDetailData_Modified.bas

**Before:** Line 598 wrote `PYEarnedRev - PYCost` to COLAPYCalcProfit (AI), overwriting the sheet formula `=AG6-AH6`.

**After:** Line removed. The sheet formula auto-calculates once AG and AH are correct.

### Fixes 4 & 5: Workbook Label Changes

Changed via COM automation (unprotect → edit → reprotect → save):
- Jobs-Ops AD3: "JTD Change In Profit" → "MTD Change in Profit"
- Jobs-GAAP AA3: "JTD Change In Profit" → "MTD Change in Profit"

---

## Current Workbook Versions

| Version | Purpose | Status |
|---------|---------|--------|
| **5.68p** | Nicole/Cindy demo version | Frozen |
| **5.70p** | Query fixes + Push to Vista button | Superseded by 5.71p |
| **5.71p** | Nicole feedback fixes + label changes | **Active — ready for re-test** |

Source code: `vba_source_569p/` (active development for 5.70p+).

---

## Remaining for Delivery

| # | Item | Priority | Status |
|---|------|----------|--------|
| R1 | Nicole/Cindy re-test on 5.71p | Critical | Ready — all 5 feedback items addressed |
| R2 | Multi-company testing (AIC Co16, APC Co12, NESM Co13) | High | Not started |
| R3 | Verify residual rounding discrepancies | High | Spot-check after Nicole re-tests |
| R4 | Wire permissions to pnp.WIPSECGetRole | Medium | Not started |
| R5 | Security cleanup (Settings sheet, test credentials) | High | Not started |
| R6 | SQL driver install for Brian Platten + Harbir Atwal | Medium | Not started |
| R7 | Nicole questions: GAAP Billed to Date source, decimal display | Low | Pending response |
| R8 | Circular reference on 51.1158 (pre-existing) | Low | Won't block delivery |

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
| Query fixes — Crystal Report exact match | Apr 7 | Done Apr 6 |
| **Nicole feedback — 5 items resolved** | **Apr 7-11** | **Done Apr 7** |
| Nicole/Cindy re-validation (5.71p) | Apr 7-11 | Ready |
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
