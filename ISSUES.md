# Auto-WIP Issue Tracker
*Maintained manually — canonical source of truth for all reported issues and their resolution.*

---

## Open Issues

### OPEN-1: Column AI value discrepancy for 51.1129
- **Reported:** 2026-04-07 (Nicole email)
- **Status:** Needs confirmation from Nicole
- **Description:** Nicole expected AI = $6,957,477 (from Michael's WIP run for June 2025). Our value is $7,950,623.36 (= prior year Dec 2024 bonus from LylesWIP). Nicole approved 5.71p which produced the same $7,950,623 value, so she may have accepted it. Her stated AG ($71,707,694) minus AH ($63,757,071) = $7,950,623, not $6,957,477 — the numbers she gave are internally inconsistent with `=AG-AH`.
- **Action:** Confirm with Nicole on next call whether $7,950,623 is correct for December 2025.

### OPEN-2: Jobs-GAAP Billed to Date source
- **Reported:** 2026-04-07 (Josh email to Nicole)
- **Status:** Awaiting Nicole's answer
- **Description:** Is the Billed to Date source on Jobs-GAAP the same "JC Cost and Revenue" Crystal Report, or a different report? Currently using `vrvJBProgressBills` for both sheets.

### OPEN-3: Dollar display format
- **Reported:** 2026-04-07 (Josh email to Nicole)
- **Status:** Awaiting Nicole's answer
- **Description:** Should dollar columns display with cents ($1,234.56) or whole dollars ($1,235)?

### OPEN-4: Percent complete format on Jobs-GAAP
- **Reported:** 2026-04-06 (demo session)
- **Status:** Low priority — cosmetic
- **Description:** Nicole's WIP shows 75.8% (one decimal). Our formula calculates to many decimal places. Nicole said fixing the cost cutoff would likely cascade-fix this. Cost cutoff was fixed in 5.70p. Not re-raised since.

---

## Resolved Issues

### Fixed in Rev 5.73p (2026-04-13)

#### 5.73-1: Missing jobs 51.1156 / 51.1157
- **Reported:** 2026-04-10 (Nicole email)
- **Symptom:** Overhead cost-only jobs not appearing on WIP
- **Root cause (dual):**
  1. JobList CTE HardClosed filter excluded jobs closed AFTER the WIP date (MonthClosed=2026-03-01 > Dec 2025). These jobs were open during Dec 2025 but Vista's current status is HardClosed.
  2. Zero-activity WHERE clause excluded jobs with no actual cost AND no contract, even if they had estimates.
- **Fix:** HardClosed filter now includes jobs where `MonthClosed > @Month` (open at WIP date). WHERE clause softened to also check OrigEstCost and ProjFinalCost.
- **Files:** `vba_source/VistaData.bas`, `sql/WIP_Vista_Query.sql`
- **Validated:** 28 jobs in Div51 Dec 2025 — exact match to LylesWIP override count.

#### 5.73-2: Column AC (JTD Prior Profit) regressed
- **Reported:** 2026-04-10 (Nicole email — "fixed in a prior version but re-broke in 5.72p")
- **Symptom:** Column AC showing wrong values
- **Root cause:** `MergePriorMonthProfitsOntoSheet` in 5.72p wrote `opsRev - opsCost` (projected profit) to the visible column COLJTDPriorProfit instead of the actual `bonusProfit` from LylesWIP. The 569p fix (Z-column-only writes) was lost during the 5.72p rewrite.
- **Fix:** Write `bonusProfit` (from `ov(8)`) to both Z-column (`COLZPriorBonusProfit`) AND visible column (`COLJTDPriorProfit`). Also write `opsRev - opsCost` to both `COLZPriorJTDOPsProfit` and `COLPriorProjProfit`. Both Z-columns and visible columns are needed because data rows have values, not formulas.
- **Files:** `vba_source/LylesWIPData.bas` (MergePriorMonthProfitsOntoSheet)
- **Validated:** 51.1129 AC = $8,997,122.18

#### 5.73-3: Column AG (Previous Year Revenue) regressed
- **Reported:** 2026-04-10 (Nicole email)
- **Symptom:** AG missing the bonus component (showing PYCost only instead of PYCost + bonus)
- **Root cause:** `MergePriorYearBonusOntoSheet` in 5.72p only wrote the bonus to COLAPYBonusProfit (AJ) but did NOT recalculate AG = AH + bonus. The 569p correction lines were lost.
- **Fix:** After writing bonus to AJ, also write `AG = PYCost + bonus` and `AI = bonus`.
- **Files:** `vba_source/LylesWIPData.bas` (MergePriorYearBonusOntoSheet)
- **Validated:** 51.1129 AG = $71,707,694.03

#### 5.73-4: Column AU (JTD Billings) regressed
- **Reported:** 2026-04-10 (Nicole email)
- **Symptom:** JTD Billings showing live bJCCM running total instead of date-filtered value
- **Root cause:** `VistaData.bas` in 5.72p lost the `BilledThruMonth` CTE from 569p. The VBA query used `jl.BilledAmt` (live bJCCM.BilledAmt which includes billings posted after the batch month) instead of `vrvJBProgressBills` filtered by `BillMonth <= @Month`.
- **Fix:** Restored BilledThruMonth CTE. SELECT now uses `ISNULL(bt.BilledAmt, 0)` for both BilledAmt and OrgBilledAmt. Added LEFT JOIN.
- **Files:** `vba_source/VistaData.bas`, `sql/WIP_Vista_Query.sql`
- **Validated:** 51.1129 AU = $96,918,206.90 (Crystal Report exact match)

#### 5.73-5: Column AI (Prior Year Calc Profit) stale values
- **Reported:** 2026-04-13 (found during email audit)
- **Symptom:** AI showing -$21.2M for 51.1129 instead of ~$7.95M
- **Root cause:** `GetWIPDetailData_Modified.bas` line 674 writes `PYEarnedRev - PYCost` to AI for ALL sheets. The 569p version had a guard (`' do not overwrite it'`) that skipped AI on the Ops sheet, relying on the template formula `=AG-AH`. But data rows have values, not formulas, so when 5.72p wrote a Vista-computed value and then MergePriorYearBonusOntoSheet corrected AG without updating AI, AI retained the stale value.
- **Fix:** `MergePriorYearBonusOntoSheet` now explicitly writes `AI = bonus` after correcting AG.
- **Files:** `vba_source/LylesWIPData.bas` (MergePriorYearBonusOntoSheet)
- **Validated:** 51.1129 AI = $7,950,623.36 (= AG - AH)

#### 5.73-6: SaveJobRow BonusProfit corruption on GAAP sheet saves (latent bug)
- **Reported:** 2026-04-13 (found during investigation)
- **Symptom:** Not yet triggered in production, but BonusProfit would be erased to NULL when a user clicks "GAAP Done" on the Jobs-GAAP sheet.
- **Root cause:** `COLZOPsBonusNew` doesn't exist on Sheet12 (GAAP). SaveJobRow's `On Error Resume Next` silently caught the error, leaving `bonusProfit = Null`. The stored proc MERGE unconditionally wrote `BonusProfit = @BonusProfit` (NULL), destroying the existing value.
- **Fix (VBA):** Only read BonusProfit on Sheet11 (Ops). On Sheet12, leave as Null.
- **Fix (SQL):** Changed stored proc to `BonusProfit = ISNULL(@BonusProfit, BonusProfit)` — NULL means "keep existing."
- **Files:** `vba_source/LylesWIPData.bas` (SaveJobRow), `sql/LylesWIP_CreateDB.sql` (LylesWIPSaveJobRow proc)
- **Deployed:** Stored proc updated on 10.103.30.11

### Fixed in Rev 5.72p (2026-04-08)

#### 5.72-1: Circular reference on 51.1142
- **Reported:** 2026-04-08 (Nicole email)
- **Root cause:** Vista ContractStatus=3 (HardClosed) passed raw to VBA which only handled 1/2. The 2-to-3 transition fired subtotal logic incorrectly, creating SUM ranges that included themselves.
- **Fix:** ContractStatus mapped to 1 (Open) or 2 (Closed) in SQL CASE expression.
- **Files:** `vba_source/VistaData.bas`

#### 5.72-2: Closed jobs in rows 37-44 shouldn't be showing
- **Reported:** 2026-04-08 (Nicole email — "51.1102-51.1149 shouldn't be showing")
- **Root cause:** EXISTS clauses caught $0 ghost transactions (SL records with ActualCost=0) pulling in long-closed jobs from prior years.
- **Fix:** HardClosed jobs only included if MonthClosed falls within the WIP year (Jan 1 - Dec 1). Also added zero-activity/zero-contract exclusion.
- **Files:** `vba_source/VistaData.bas`, `sql/WIP_Vista_Query.sql`

#### 5.72-3: Subtotals and totals showing $0
- **Reported:** 2026-04-08 (Nicole email — "column W")
- **Root cause:** Override columns (I=Revenue, M=Cost) left at $0 when no explicit override existed. All downstream formulas depend on I/M.
- **Fix:** When no override, default I to Vista projected revenue and M to MAX(Vista projected cost, actual cost). Matches Michael's LCGWIPCreateBatch defaulting logic.
- **Files:** `vba_source/GetWIPDetailData_Modified.bas`

#### 5.72-4: Job 51.1158 Column W = $0 (should be $56,961)
- **Reported:** 2026-04-08 (Nicole email)
- **Root cause:** Same as 5.72-3 (override columns empty), plus template circular formula in Column Z (W-Z-Y-W chain) for zero-profit jobs.
- **Fix:** Override defaulting + explicit Column Z write to break the circular chain.
- **Files:** `vba_source/GetWIPDetailData_Modified.bas`

#### 5.72-5: Job 51.1139 missing from WIP
- **Reported:** 2026-04-08 (Nicole email — "open in 2025 but had no cost or billings")
- **Nicole's rule:** If a job was open on the prior year's (12/31/24) WIP, it has to appear on the following year's WIP regardless of cost/billing status.
- **Root cause:** Job was included by HardClosed-in-WIP-year filter (MonthClosed = Dec 2025) but sorted into Open section instead of Closed.
- **Fix:** ContractStatus mapping shows Vista-closed jobs in the Closed section.
- **Files:** `vba_source/VistaData.bas`

#### 5.72-6: Cost aggregation used PostedDate instead of Mth
- **Reported:** 2026-04-06 (demo) / fixed in 5.72p
- **Root cause:** VBA inline SQL used PostedDate/ActualDate for cost filtering. Crystal Report and Michael's procs use fiscal Mth.
- **Fix:** Changed all date filters in JobCosts CTE to use `d.Mth`.
- **Validated:** 51.1129 JTD Cost = $87,315,159.22 (Crystal Report exact match)

### Fixed in Rev 5.71p (2026-04-07)

#### 5.71-1: Column AC pulling GAAP value instead of Ops
- **Reported:** 2026-04-07 (Nicole email — "AC should be $8,997,122")
- **Root cause:** MergePriorMonthProfitsOntoSheet wrote to visible column directly, overwriting the formula with a computed value from the wrong source.
- **Fix:** Rewrote to populate hidden Z-column COLZPriorBonusProfit with BonusProfit.
- **Note:** This fix was lost in 5.72p and re-fixed in 5.73p (see 5.73-2).

#### 5.71-2: Column AG missing bonus component
- **Reported:** 2026-04-07 (Nicole email — "AG should be $71,707,694")
- **Root cause:** Vista stubs PriorYrBonusProfit to 0. MergePriorYearBonusOntoSheet only updated AJ (bonus column), not AG (total).
- **Fix:** Expanded function to update AG = PYCost + bonus.
- **Note:** This fix was lost in 5.72p and re-fixed in 5.73p (see 5.73-3).

#### 5.71-3: Column AI incorrect
- **Reported:** 2026-04-07 (Nicole email — "AI should be $6,957,477")
- **Fix:** Removed direct VBA overwrite; relied on formula =AG-AH.
- **Note:** This approach broke in 5.72p because data rows don't have formulas. Re-fixed in 5.73p with explicit write (see 5.73-5).

#### 5.71-4: Column AD / AA rename
- **Reported:** 2026-04-07 (Nicole email)
- **Fix:** Renamed to "MTD Change in Profit" on both Jobs-Ops (Col AD) and Jobs-GAAP (Col AA).

### Fixed in Rev 5.70p (2026-04-06)

#### 5.70-1: JTD Cost pulling transactions beyond batch month
- **Reported:** 2026-04-06 (demo session)
- **Root cause:** Date cutoff for bJCCD cost query was not properly bounded by batch month end.
- **Fix:** All cost aggregation filters use fiscal `Mth <= @Month`.
- **Validated:** 51.1129 JTD Cost = $87,315,159.22

#### 5.70-2: Billed to Date from AR instead of JB
- **Reported:** 2026-04-06 (demo session)
- **Root cause:** Using bARTH.Invoiced (AR side) instead of vrvJBProgressBills. AR and JB track billing differently — $236K discrepancy on 51.1129.
- **Fix:** BilledThruMonth CTE using vrvJBProgressBills.AmountBilled_ThisBill filtered by BillMonth <= @Month.
- **Validated:** 51.1129 = $96,918,206.90 (Crystal Report exact match)
- **Note:** This CTE was lost in 5.72p and restored in 5.73p (see 5.73-4).

#### 5.70-3: Job 54.9416 excluded despite having December costs
- **Reported:** 2026-04-06 (demo session)
- **Root cause:** Job had StartMonth = January 2026 but preliminary costs hit in November/December 2025.
- **Nicole's rule:** Two-tier inclusion: (1) if any cost has hit the job, include it; (2) if no cost but StartMonth <= batch month, include it for backlog.
- **Fix:** Added EXISTS check for actual cost records in JobList CTE as alternate inclusion criterion.
- **Files:** `vba_source/VistaData.bas`

#### 5.70-4: Prior month profit showing projected instead of recognized
- **Reported:** 2026-04-06 (demo session)
- **Root cause:** Column AC was pulling anticipated/projected profit instead of the prior month's recognized profit (Column Z value from LylesWIP).
- **Fix:** MergePriorMonthProfitsOntoSheet reads BonusProfit from prior month's LylesWIP override data.

#### 5.70-5: Job 51.1158 circular reference
- **Reported:** 2026-04-06 (demo session)
- **Root cause:** Pre-existing formula issue in the original workbook (Rev 5.47p) that doesn't handle zero-profit jobs cleanly. Not introduced by Auto WIP changes.
- **Fix:** Resolved in 5.72p via override defaulting (see 5.72-4).

---

## Business Rules (from Nicole / Cindy)

These rules were established through the email chain and demo sessions. Reference for future development.

1. **Job inclusion (two-tier):** Include if (a) any actual cost exists in bJCCD through batch month, OR (b) StartMonth <= batch month end date. *(Source: 04/06 demo, Nicole)*
2. **HardClosed jobs:** Include if MonthClosed is within the WIP year OR after the WIP month (job was open at WIP date). Exclude if closed before the WIP year. *(Source: 04/08 Nicole email, refined 04/13)*
3. **Prior year WIP carry-forward:** If a job was open on 12/31 of the prior year, it must appear on the following year's WIP regardless of cost/billing status. *(Source: 04/08 Nicole email)*
4. **Override priority:** LylesWIP override if present; otherwise Vista-calculated value. Never silently discard an override. *(Source: CLAUDE.md)*
5. **GAAP is quarterly only:** March, June, September, December. Non-GAAP months = blank GAAP values. *(Source: CLAUDE.md)*
6. **Completion dates from WIP history:** Come from Nicole's override files, not from Vista JC Contracts. *(Source: 04/06 demo, Nicole)*
7. **Save trigger:** Double-click Col H (Ops Done) or Col I (GAAP Done) = save. NOT auto-save on cell change. *(Source: CLAUDE.md)*
8. **Closed jobs = 100% complete:** Closed jobs should always show 100% complete on the WIP. *(Source: 04/08 Nicole email)*
9. **Cost source matching:** All cost/billing values must match the "JC Cost and Revenue" Crystal Report. Use fiscal Mth for date filtering, not PostedDate. *(Source: 04/06 demo, validated 04/08)*
10. **Override defaulting:** When no explicit override exists, default Revenue to Vista projected revenue and Cost to MAX(Vista projected cost, actual cost). Matches Michael's LCGWIPCreateBatch. *(Source: 04/08 fixes)*

---

## Architectural Lessons Learned

1. **Data rows have values, not formulas.** Template row 6 has formulas but GetWipDetail2 writes values to data rows (9+). Any merge function that needs to update a visible column MUST write the value directly — cannot rely on formulas propagating from Z-columns. *(Learned: 5.73p fixes)*

2. **Z-columns AND visible columns.** Write to Z-columns for SaveJobRow/tooltip backing, AND to visible columns for display. Neither alone is sufficient. *(Learned: 5.73p regression on Col AC/Q)*

3. **Version regression risk.** When making significant VBA changes, diff all modified functions against the prior working version. The 5.72p rewrite introduced many critical fixes but lost three 569p improvements (BilledThruMonth, Z-column writes, AG recalculation). The 569p archive (`vba_source_569p/`) exists for this purpose. *(Learned: 5.72p → 5.73p)*

4. **SQL TRIM in JOINs kills performance.** Never `LTRIM(RTRIM(job))` in JOIN/GROUP BY on bJCCD. Raw field only in joins; trim in SELECT. Violating this: 9-minute queries. Correct: 58 seconds. 8.8M rows. *(Learned: early development)*
