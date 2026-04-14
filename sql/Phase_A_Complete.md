# Phase A — Vista Query Bug Fixes
**Completed: 2026-03-31**
**Database:** Viewpoint (10.112.11.8)
**Query file:** `sql/WIP_Vista_Query.sql`
**Validated against:** WML Company 15, December 2025

---

## Summary of Changes

All fixes are applied directly to `sql/WIP_Vista_Query.sql`. The query now correctly
matches the behavior of the JC Cost & Revenue Crystal Report (`brptJCCostRev`) that
Nicole, Brian, and the division controllers use to produce the Division Job Summary (DJS).

**Net job count change for WML Dec 2025 (all departments):**

| State | Jobs |
|-------|------|
| Original query (before any fixes) | 1,460 |
| After A3 (future jobs removed) | 1,424 |
| After A5 + A6 (correct jobs added back) | **1,434** |

---

## A1 — Date Parameter Direction Confirmed
**Status:** Done
**Owner:** Josh
**Source:** Nicole Leasure email, 2026-03-31

Nicole confirmed: JC Cost & Revenue Crystal Report uses **Ending Month = batch month**
(e.g. `12/31/2025` for December 2025 WIP). Beginning Month is left blank.

Crystal Report ODBC call decoded:
```
{CALL "brptJCCostRev";1(
    15,                            -- @JCCo            = Company
    ' ',                           -- @BeginContract    = blank (all jobs)
    'zzzzzzzzz',                   -- @EndContract      = all jobs
    {ts '2050-01-01 00:00:00'},    -- @ThroughMth       = Ending Month (Nicole enters 12/31/2025)
    {ts '1950-01-01 00:00:00'},    -- @BegMth           = Beginning Month (blank = JTD)
    NULL,                          -- @BegDept
    NULL,                          -- @EndDept
    'O',                           -- @Status           = Open jobs only
    NULL,                          -- @BegMthClosed
    {ts '2025-12-31 00:00:00'}     -- @EndMthClosed     = Ending Closed Month
)}
```

Key takeaway: `@ThroughMth` = `EOMONTH(@WipMonth)`. Beginning Month is intentionally
blank — the report returns JTD totals from job inception through the batch month end.
`@Status = 'O'` was noted as a potential contributor to A6 (see below).

**Query variable added:**
```sql
DECLARE @CutOffDate date = EOMONTH(@Month);  -- matches @ThroughMth
DECLARE @BillingMth smalldatetime = DATEADD(month, DATEDIFF(month, 0, @Month), 0);
-- @CurrentDate REMOVED entirely — all filters use @CutOffDate
```

---

## A2 — Batch-Month Cutoff Applied to OE/CO Transactions
**Status:** Done
**Owner:** Josh

**Problem:** Three places in the query used `@CurrentDate` (= `GETDATE()`) instead of
`@CutOffDate`. This caused Original Estimate (OE) and Change Order (CO) transactions
posted after the batch month to be included in historical runs.

**Locations fixed (3):**

| CTE | Field | Was | Fixed |
|-----|-------|-----|-------|
| JobCosts | `OrigEstCost` (OE transactions) | `<= @CurrentDate` | `<= @CutOffDate` |
| JobCosts | `COEstCost` (CO transactions) | `<= @CurrentDate` | `<= @CutOffDate` |
| COContracts | `COContractAmt` (CO contract amounts) | `<= @CurrentDate` | `<= @CutOffDate` |

**Fix applied in query:**
```sql
-- OrigEstCost
WHEN d.JCTransType = 'OE' AND d.PostedDate <= @CutOffDate

-- COEstCost
WHEN d.JCTransType = 'CO' AND d.PostedDate <= @CutOffDate

-- COContracts CTE
WHEN id.JCTransType = 'CO' AND id.PostedDate <= @CutOffDate
```

---

## A3 — Future Jobs Filter
**Status:** Done
**Owner:** Josh

**Problem:** Jobs set up in Vista after the batch month appeared on historical WIP runs.
Running a December 2025 WIP in March 2026 would show 2026 jobs.

**Investigation findings:**
- `bJCJM.StartDate` is NULL for **all 1,472 WML jobs** — completely useless for filtering
- `bJCCM.StartMonth` (smalldatetime, first of contract start month) is always populated
- **36 jobs** have `StartMonth > 2025-12-31` for Co=15

**Fix applied in JobList CTE WHERE clause:**
```sql
AND c.StartMonth <= @CutOffDate    -- exclude jobs set up after the batch month
```

**Validation:** Exactly 36 jobs removed from Dec 2025 WML run. All confirmed as genuine
2026 startups (e.g. `50.0552.` through `50.0558.`, `51.1167.`, `51.1168.` etc.,
all with `StartMonth = 2026-01-01` and $0 ContractAmt).

---

## A4 — Billing Date-Filtered via bARTH
**Status:** Done
**Owner:** Josh

**Problem:** `bJCCM.BilledAmt` is a live running total — it includes all billings ever
posted to the contract, including those posted after the batch month. Using it on a
historical WIP run overstates billing for active jobs.

**Investigation findings:**
- `bJCBH` does not exist in the Viewpoint database
- `bARTH` (AR Transaction Header) is the correct billing source
  - `ARTransType = 'I'` = Invoice (confirmed: only type with billing amounts)
  - Other types: `'R'` = Payment/Receipt, `'A'` = Adjustment ($0), `'V'` = Void, `'W'` = Write-off
  - Filter field: `Mth` (smalldatetime, first of accounting month) — matches Crystal Report
    period behavior vs `TransDate` (calendar date)

**Validation on key jobs:**

| Job | Live BilledAmt | Filtered (≤ Dec 2025) | Difference |
|-----|---------------|----------------------|------------|
| 51.1108 | $94,054,108 | $94,045,648 | $8,460 |
| 51.1129 | $100,000,647 | $92,074,288 | **$7,926,359** |
| 51.1151 | $5,756,926 | $5,009,265 | $747,661 |

Job 51.1129 had a **$7.9M overstatement** using the live total.

**New CTE added (BilledThruMonth):**
```sql
BilledThruMonth AS (
    SELECT
        ar.JCCo,
        ar.Contract,
        SUM(ar.Invoiced) AS BilledAmt
    FROM bARTH ar WITH (NOLOCK)
    WHERE ar.JCCo = @Co
      AND ar.ARTransType = 'I'
      AND ar.Mth <= @BillingMth      -- @BillingMth = first of batch month
    GROUP BY ar.JCCo, ar.Contract
),
```

**Final SELECT updated:**
```sql
-- Was: jl.BilledAmt
ISNULL(bt.BilledAmt, 0) AS BilledAmt
ISNULL(bt.BilledAmt, 0) AS OrgBilledAmt
```

**JOIN added:**
```sql
LEFT JOIN BilledThruMonth bt ON jl.JCCo = bt.JCCo AND jl.Contract = bt.Contract
```

---

## A5 — Jobs Closed In or After the Batch Month
**Status:** Done
**Owner:** Josh

**Problem:** `bJCJM.JobStatus` reflects the job's CURRENT status, not its status as of
the batch month. Running a December 2025 WIP in March 2026 means jobs closed in January
or February 2026 show `JobStatus = 2` (hard closed) today but were open in December 2025.
These jobs would only be included by the cost EXISTS clause — if they had no December cost
activity they would drop off entirely.

Also affects jobs closed during the batch month itself (MonthClosed = batch month).

**Investigation findings:**
- `bJCCM.MonthClosed` = smalldatetime, first of the closing month — always populated for
  closed jobs
- **46 jobs** for WML Co=15 have `MonthClosed = 2025-12-01` (closed during December 2025)
  and would be missed without this fix if they had no December cost activity
- The 10 jobs closed in March 2025 with no 2025 activity are correctly excluded — they
  were genuinely closed months before the batch month

**Fix added to JobList WHERE OR conditions:**
```sql
-- A5: Jobs closed IN or AFTER the batch month were still open when
-- the WIP was produced. MonthClosed = first of closing month, so
-- >= @Month catches jobs closed during the batch month itself.
OR (j.JobStatus NOT IN (1, 3) AND c.MonthClosed >= @Month)
```

**Validation:** Net +10 jobs added to Dec 2025 WML (most of the 46 December-closed jobs
were already included via cost EXISTS; A5 adds only those with no 2025 cost activity).
Job `51.1139` (closed Dec 31, 2025, no 2025 cost) is a confirmed example of A5 in action.

---

## A6 — Closed Jobs with Billing-Only Activity
**Status:** Done
**Owner:** Josh

**Problem:** The cost EXISTS clause in JobList only checks `bJCCD` (cost transactions).
Closed jobs that received billing activity in the current year but no corresponding cost
(e.g. retainage release invoices, warranty final billings posted after job close) are
invisible to the cost EXISTS and would be excluded.

Crystal Report `@Status = 'O'` was noted as a potential contributor — the Crystal Report
also only shows Open jobs, but the WIP schedule has different inclusion rules.

**Investigation findings:**
- **8 jobs** for Co=15 are closed with `MonthClosed = March 2025` but had 2025 billing
  in `bARTH` (retainage invoices, small amounts $1,760–$2,806 each, all Dept 54 Bakersfield)
- These are NOT caught by A5 (`MonthClosed = March 2025 < @Month`)
- These are NOT caught by cost EXISTS (no 2025 cost transactions)
- `bARTH.ARTransType = 'I'` and `Mth BETWEEN @StartDate AND @BillingMth` correctly
  identifies billing-only activity in the current year

**Fix added to JobList WHERE OR conditions:**
```sql
-- A6: Closed jobs with billing activity in the current year but no cost.
OR EXISTS (
    SELECT 1 FROM bARTH ar WITH (NOLOCK)
    WHERE ar.JCCo = j.JCCo AND ar.Contract = c.Contract
    AND ar.ARTransType = 'I'
    AND ar.Mth BETWEEN @StartDate AND @BillingMth
)
```

---

## A7 — Zero-JTD Reversal Jobs
**Status:** Already working — no changes needed
**Owner:** Josh

**Problem reported:** Jobs 56.1022 and 56.1057 were missing from the WIP despite having
current-period reversal activity. Both jobs have JTD actual cost = $0.00 because reversals
posted in 2025 exactly cancel prior-year cost.

**Investigation findings:**

| Job | MonthClosed | Reversal | JTD Cost | CY Cost |
|-----|-------------|----------|----------|---------|
| 56.1022 | Dec 2024 | Feb 2025: JC -$88,221 | $0.00 | **-$88,067** |
| 56.1057 | Apr 2025 | Apr 2025: JC -$18,925 | $0.00 | **-$16,239** |

Both jobs have 2025 cost transactions (the reversals themselves) which satisfy the
cost EXISTS clause. The query includes them and correctly surfaces the negative CY cost,
which is exactly what Nicole needs to see — the period change is visible even when JTD = $0.

**Root cause of original bug in 5.47p:** Michael's `WipDb` database only had WML data
through September 2025. The missing jobs were absent because they were never loaded into
WipDb, not because of a filter logic error. Our Vista-direct query has no such gap.

**No query changes made.** `CY_Cost` output for both jobs verified:
- 56.1022: CY = -$88,067 (Feb 2025 reversal visible)
- 56.1057: CY = -$16,239 (Apr 2025 reversal visible)

---

## Key Schema Discoveries

| Table | Key Findings |
|-------|-------------|
| `bJCJM` | No setup date column. `JobStatus` = current status (1=Open, 2=Closed, 3=SoftClosed) |
| `bJCCM` | `StartMonth` = contract start month (always populated, use for A3). `StartDate` = NULL for all 1,472 WML jobs (useless). `MonthClosed` = first of closing month (use for A5). `BilledAmt` = live running total (replaced by A4 fix) |
| `bJCCD` | Cost transactions. `JCTransType`: PR/EM use `ActualDate`; all others use `PostedDate`. OE=estimates, CO=change orders, PF=projected final, rest=actual cost |
| `bARTH` | AR Transaction Header. `ARTransType='I'` = invoice. `Mth` = accounting period (use for billing filter, not `TransDate`). `Invoiced` = invoice amount |
| `bJCBH` | Does NOT exist in Viewpoint production |

**WML Departments (Co=15):**

| Dept | Name | Contracts |
|------|------|-----------|
| 50 | Company | 80 |
| 51 | Rocklin | 97 |
| 52 | Fresno | 300 |
| 53 | East Bay | 16 |
| 54 | Bakersfield | 710 |
| 55 | Murrieta | 126 |
| 56 | System Integration | 129 |
| 57 | San Diego | 10 |
| 58 | Los Angeles | 6 |

---

## A8 — Vista Close Status Flag (partial — SQL side only)
**Status:** SQL side done — full mismatch requires LylesWIP (Sprint 2)
**Owner:** Josh

**Problem:** Job 54.9033 was flagged Close='Y' in the workbook but still Open in Vista
(`ContractStatus=1`). Ops and Accounting had a stale Close flag mismatch. Surfacing
this requires comparing two sources:
1. **Workbook Close flag** (col G, set by Ops double-clicking) — stored in LylesWIP (not built yet)
2. **Vista `ContractStatus`** — available now in the query

**What was implemented (Phase A):**

Added `c.MonthClosed` to the JobList CTE SELECT and exposed two new columns in the
final SELECT:

```sql
-- New column 1: raw close date from Vista
jl.MonthClosed,

-- New column 2: A8 flag
CASE
    WHEN jl.ContractStatus = 2 AND (jl.MonthClosed IS NULL OR jl.MonthClosed < @Month)
    THEN 1
    ELSE 0
END AS VistaClosedAtWipDate,
```

**ContractStatus values confirmed (WML Co=15):**
| Value | Meaning | Count (all time) |
|-------|---------|-----------------|
| 1 | Open | 176 |
| 2 | Hard Closed | 190 |
| 3 | Soft Closed | 1,106 |
| 0 | Not yet activated | 2 |

**`ContractStatus` mirrors `JobStatus` exactly** — they are always in sync for WML.

**Flag validation for Dec 2025 WML run (1,434 jobs):**
| Category | Count |
|----------|-------|
| Open in Vista (ContractStatus=1) | 140 |
| VistaClosedAtWipDate=1 (hard-closed before Dec) | 142 |
| Hard-closed DURING Dec (A5 jobs, VistaClosedAtWipDate=0) | 46 |
| Soft-closed (ContractStatus=3, VistaClosedAtWipDate=0) | 1,106 |
| **Total** | **1,434** |

**Note on 3 NULL-MonthClosed jobs:** Jobs `56.1077`, `58.0004`, `58.0005` are 2025 overhead/estimating
internal contracts with `ContractStatus=2` but `MonthClosed=NULL`. Treated as
`VistaClosedAtWipDate=1` (included in the 142 count above). They were never active project
contracts during December 2025.

**What Sprint 2 must add to complete A8:**
When LylesWIP is built, VBA code in the `Close` column double-click handler must compare:
```
If workbookCloseFlag = "Y" And VistaClosedAtWipDate = 0 Then → highlight mismatch (orange)
If workbookCloseFlag = ""  And VistaClosedAtWipDate = 1 Then → flag for review (yellow)
```
The SQL already delivers both `ContractStatus` and `VistaClosedAtWipDate` to enable this.
