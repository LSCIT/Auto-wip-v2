-- =============================================================================
-- WIP Vista Direct Query
-- Replaces: LCGWIPGetDetailPM stored procedure (WipDb)
-- Purpose:  Returns one row per job with all fields needed for WIP Schedule
-- Target:   Viewpoint database on 10.112.11.8 (test), VM111VPPRD1 (prod)
-- Auth:     WIPexcel / WIP@MR@2024
--
-- Parameters (set before execution):
--   @Co         tinyint     - Company number (e.g., 15 for WML)
--   @Month      date        - WIP month (first of month, e.g., '2025-10-01')
--   @DeptList   varchar     - Comma-separated department codes (e.g., '10,20,30')
--   @GroupBy    varchar     - 'Department' or 'PM'
-- =============================================================================

DECLARE @Co         tinyint   = 15;
DECLARE @Month      date      = '2025-10-01';
DECLARE @DeptList   varchar(200) = '10,20,30,40,50,60,70,80,90';
DECLARE @GroupBy    varchar(20) = 'Department';

-- Derived date parameters
DECLARE @CutOffDate date = EOMONTH(@Month);                    -- Last day of WIP month (= @ThroughMth in Crystal Report)
DECLARE @StartDate  date = DATEFROMPARTS(YEAR(@Month), 1, 1);  -- Jan 1 of WIP year
DECLARE @PriorYrEnd date = DATEFROMPARTS(YEAR(@Month) - 1, 12, 31); -- Dec 31 prior year
DECLARE @BillingMth smalldatetime = DATEADD(month, DATEDIFF(month, 0, @Month), 0); -- First of WIP month (for bARTH.Mth filter)
DECLARE @MarchPlug  date = DATEFROMPARTS(YEAR(@Month), 3, 31); -- March WIP baseline
-- NOTE: @CurrentDate removed. All filters must use @CutOffDate to match Crystal Report @ThroughMth behavior.
-- Confirmed by Nicole Leasure 2026-03-31: Crystal Report uses Ending Month = batch month, Beginning Month blank.

-- =============================================================================
-- CTE 1: Parse department list into a table
-- =============================================================================
;WITH DeptFilter AS (
    SELECT LTRIM(RTRIM(value)) AS Department
    FROM STRING_SPLIT(@DeptList, ',')
),

-- =============================================================================
-- CTE 2: Job List — which jobs to include
-- Active jobs + closed jobs with period activity + cross-year reversals
-- =============================================================================
JobList AS (
    SELECT
        j.JCCo,
        j.Job,                          -- Raw for JOIN performance; trim in final SELECT
        j.Description AS JobDescription,
        j.Contract,                     -- Raw for JOIN performance
        j.JobStatus,
        j.ProjectMgr,
        j.ProjMinPct,
        c.Department,
        c.Description AS ContractDescription,
        c.OrigContractAmt,
        c.ContractAmt,
        c.BilledAmt,
        c.ReceivedAmt,
        c.CurrentRetainAmt,
        c.ActualCloseDate AS CompletionDate,
        c.ContractStatus,
        c.MonthClosed,
        d.Description AS DeptDescription,
        ISNULL(pm.Name, '') AS PM
    FROM bJCJM j WITH (NOLOCK)
    JOIN bJCCM c WITH (NOLOCK) ON j.JCCo = c.JCCo AND j.Contract = c.Contract
    JOIN bJCDM d WITH (NOLOCK) ON c.JCCo = d.JCCo AND c.Department = d.Department
    JOIN DeptFilter df ON c.Department = df.Department
    LEFT JOIN bJCMP pm WITH (NOLOCK) ON j.JCCo = pm.JCCo AND j.ProjectMgr = pm.ProjectMgr
    WHERE j.JCCo = @Co
      AND (
        -- A3: Primary inclusion — job set up on or before the batch month (backlog + active)
        c.StartMonth <= @CutOffDate
        -- Nicole rule (2026-04-06 demo): include jobs with cost activity through the batch
        -- month even if StartMonth is after. Example: 54.9416 has costs in Nov/Dec 2025
        -- but StartMonth = Jan 2026 (preliminary design work before official start).
        OR EXISTS (
            SELECT 1 FROM bJCCD cd WITH (NOLOCK)
            WHERE cd.JCCo = j.JCCo AND cd.Job = j.Job
            AND cd.JCTransType NOT IN ('OE','CO','PF')
            AND cd.Mth <= @Month
        )
      )
      AND (
        -- Open jobs (Status 1) always included.
        j.JobStatus = 1

        -- Soft Closed jobs (Status 2) always included — they're still editable in Vista.
        OR j.JobStatus = 2

        -- Hard Closed jobs (Status 3):
        --   1. Closed AFTER the WIP month: job was open at the WIP date, include it.
        --      Example: 51.1156 MonthClosed=2026-03 but Dec 2025 WIP → job was still open.
        --   2. Closed within the WIP year (Jan 1 – Dec 1): include as closed.
        --      Example: 51.1139 MonthClosed=Dec 2025 → show in Closed section.
        -- This excludes only jobs closed BEFORE the WIP year (e.g. 51.1102 Dec 2024).
        OR (j.JobStatus = 3
            AND (ISNULL(c.MonthClosed, '2050-01-01') > @Month
                 OR (ISNULL(c.MonthClosed, '1900-01-01') >= @StartDate
                     AND c.MonthClosed <= DATEFROMPARTS(YEAR(@Month), 12, 1))))
      )
),

-- =============================================================================
-- CTE 3: Job-Level Cost Aggregation
-- Aggregates bJCCD transactions to job level.
-- IMPORTANT: All date filters use Mth (fiscal month), NOT PostedDate/ActualDate.
-- This matches the Crystal Report "JC Cost and Revenue" which filters by
-- Ending Month = batch month. Confirmed 2026-04-06: filtering by Mth produces
-- exact match to Crystal Report (87,315,159.22 for job 51.1129 Dec 2025).
-- Prior approach used PostedDate/ActualDate which diverges from fiscal month.
-- =============================================================================
JobCosts AS (
    SELECT
        d.JCCo,
        d.Job,                          -- Raw for JOIN performance; trim in final SELECT

        -- JTD Actual Cost (all actual transaction types through batch month)
        SUM(CASE
            WHEN d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth <= @Month
                THEN d.ActualCost ELSE 0 END)
        AS ActualCost,

        -- Current Year Actual Cost
        SUM(CASE
            WHEN d.JCTransType NOT IN ('OE','CO','PF')
                AND d.Mth BETWEEN @StartDate AND @Month
                THEN d.ActualCost ELSE 0 END)
        AS CYActualCost,

        -- Period (Month) Actual Cost
        SUM(CASE
            WHEN d.JCTransType NOT IN ('OE','CO','PF')
                AND d.Mth = @Month
                THEN d.ActualCost ELSE 0 END)
        AS PeriodActualCost,

        -- Original Estimate (OE transactions)
        SUM(CASE
            WHEN d.JCTransType = 'OE' AND d.Mth <= @Month
                THEN d.EstCost ELSE 0 END)
        AS OrigEstCost,

        -- Change Order Estimate (CO transactions)
        SUM(CASE
            WHEN d.JCTransType = 'CO' AND d.Mth <= @Month
                THEN d.EstCost ELSE 0 END)
        AS COEstCost,

        -- Projected Final Cost (PF transactions, CostType=99)
        SUM(CASE
            WHEN d.JCTransType = 'PF' AND d.Mth <= @Month
                THEN d.ProjCost ELSE 0 END)
        AS ProjFinalCost,

        -- Prior Period Projected Cost (for comparison — previous month)
        SUM(CASE
            WHEN d.JCTransType = 'PF' AND d.Mth < @Month
                THEN d.ProjCost ELSE 0 END)
        AS PriorProjCost,

        -- March Baseline Projected Cost (for Prior Projected Profit, Col R)
        SUM(CASE
            WHEN d.JCTransType = 'PF' AND d.Mth <= @MarchPlug
                THEN d.ProjCost ELSE 0 END)
        AS MarchProjCost,

        -- Change Order Contract Amount (from bJCCD CO transactions on contract items)
        -- Note: CO contract amounts are tracked at contract item level, not cost detail
        0 AS COContractAmt,

        -- Labor breakdown (CostType 1, 11)
        SUM(CASE
            WHEN d.CostType IN (1,11) AND d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth <= @Month
                THEN d.ActualCost ELSE 0 END)
        AS JTDLaborCost,

        -- Equipment breakdown (CostType 2, 3, 12)
        SUM(CASE
            WHEN d.CostType IN (2,3,12) AND d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth <= @Month
                THEN d.ActualCost ELSE 0 END)
        AS JTDEquipCost,

        -- Subcontractor (CostType 5)
        SUM(CASE
            WHEN d.CostType = 5 AND d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth <= @Month
                THEN d.ActualCost ELSE 0 END)
        AS JTDSubCost,

        -- Material + Other (CostType 4, 6, 7, 13)
        SUM(CASE
            WHEN d.CostType IN (4,6,7,13) AND d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth <= @Month
                THEN d.ActualCost ELSE 0 END)
        AS JTDMatlOtherCost

    FROM bJCCD d WITH (NOLOCK)
    WHERE d.JCCo = @Co
    GROUP BY d.JCCo, d.Job
),

-- =============================================================================
-- CTE 4: Change Order Contract Amounts (from contract item detail table)
-- =============================================================================
COContracts AS (
    SELECT
        id.JCCo,
        j.Job,                          -- Raw for JOIN performance
        SUM(CASE WHEN id.JCTransType = 'CO' AND id.Mth <= @Month
            THEN id.ContractAmt ELSE 0 END) AS COContractAmt
    FROM bJCID id WITH (NOLOCK)
    JOIN bJCJM j WITH (NOLOCK) ON id.JCCo = j.JCCo AND id.Contract = j.Contract
    WHERE id.JCCo = @Co
    GROUP BY id.JCCo, j.Job
),

-- =============================================================================
-- CTE 5: JCCP ProjPlug flags (most recent per job)
-- =============================================================================
ProjPlugs AS (
    SELECT JCCo, Job,               -- Raw for JOIN performance
        MAX(CASE WHEN ProjPlug = 'Y' THEN 1 ELSE 0 END) AS HasProjPlug
    FROM JCCP WITH (NOLOCK)
    WHERE JCCo = @Co
      AND Mth <= DATEADD(month, DATEDIFF(month, 0, @CutOffDate), 0)
    GROUP BY JCCo, Job
),

-- =============================================================================
-- CTE 6: budProjInfo (latest PM projection notes per job)
-- =============================================================================
ProjInfo AS (
    SELECT Co, Job,                  -- Raw for JOIN performance
        DollarsPlug, UnitsPlug, RemCostPerUnit, RCPUPlug, myNotes
    FROM (
        SELECT *, ROW_NUMBER() OVER (
            PARTITION BY Co, Job
            ORDER BY myDate DESC
        ) AS rn
        FROM budProjInfo WITH (NOLOCK)
        WHERE Co = @Co AND myDate <= @CutOffDate
    ) sub
    WHERE rn = 1
),

-- =============================================================================
-- CTE 7: March Baseline Projected Revenue (for Prior Projected Profit)
-- Uses the projected contract amount as of March for the WIP year
-- =============================================================================
MarchBaseline AS (
    SELECT
        jor.JCCo,
        j.Job,                       -- Raw for JOIN performance
        jor.OtherAmount AS MarchProjRevenue
    FROM (
        SELECT JCCo, Contract, OtherAmount,
            ROW_NUMBER() OVER (PARTITION BY JCCo, Contract ORDER BY Month DESC) AS rn
        FROM bJCOR WITH (NOLOCK)
        WHERE JCCo = @Co AND Month <= @MarchPlug
    ) jor
    JOIN bJCJM j WITH (NOLOCK) ON jor.JCCo = j.JCCo AND jor.Contract = j.Contract
    WHERE jor.rn = 1
),

-- =============================================================================
-- CTE 8: Date-filtered billing from JB Progress Bills (A4 fix, updated 2026-04-06)
-- bJCCM.BilledAmt is a live running total — it includes billings posted after
-- the batch month. Using vrvJBProgressBills.AmountBilled_ThisBill which matches
-- the Crystal Report "JC Cost and Revenue" Billed Amount exactly.
-- Prior approach used bARTH.Invoiced (AR side) which was $236K off for 51.1129
-- because AR and JB track billing differently.
-- Validated: 51.1129 = 96,918,206.90 — exact match to Crystal Report.
-- =============================================================================
BilledThruMonth AS (
    SELECT
        pb.JBCo AS JCCo,
        pb.Contract,
        SUM(pb.AmountBilled_ThisBill) AS BilledAmt
    FROM vrvJBProgressBills pb WITH (NOLOCK)
    WHERE pb.JBCo = @Co
      AND pb.BillMonth <= @Month
    GROUP BY pb.JBCo, pb.Contract
),

-- =============================================================================
-- CTE 10: Prior-quarter GAAP Revenue snapshot (Bug 3 — feeds COLZPriorJTDGAAPProfit)
-- Most recent bJCOR row with udPlugged='Y' before the current WIP month
-- =============================================================================
PriorGAAPRev AS (
    SELECT jor.JCCo, j.Job,
           ISNULL(jor.RevCost, 0) AS LastGAAPRev,
           jor.udPlugged AS LastGAAPRevPlugged
    FROM (
        SELECT JCCo, Contract, RevCost, udPlugged,
               ROW_NUMBER() OVER (PARTITION BY JCCo, Contract ORDER BY Month DESC) AS rn
        FROM bJCOR WITH (NOLOCK)
        WHERE JCCo = @Co AND udPlugged = 'Y' AND Month < @Month
    ) jor
    JOIN bJCJM j WITH (NOLOCK) ON j.JCCo = jor.JCCo AND j.Contract = jor.Contract
    WHERE jor.rn = 1
),

-- =============================================================================
-- CTE 11: Prior-quarter GAAP Cost snapshot (Bug 3 — feeds COLZPriorJTDGAAPProfit)
-- Most recent bJCOP row with udPlugged='Y' before the current WIP month
-- =============================================================================
PriorGAAPCost AS (
    SELECT op.JCCo, op.Job,
           ISNULL(op.ProjCost, 0) AS LastGAAPCost,
           op.udPlugged AS LastGAAPCostPlugged
    FROM (
        SELECT JCCo, Job, ProjCost, udPlugged,
               ROW_NUMBER() OVER (PARTITION BY JCCo, Job ORDER BY Month DESC) AS rn
        FROM bJCOP WITH (NOLOCK)
        WHERE JCCo = @Co AND udPlugged = 'Y' AND Month < @Month
    ) op WHERE op.rn = 1
),

-- =============================================================================
-- CTE 12: Prior-year JTD cost as of Dec 31 prior year (Bug 4 — feeds COLZGAAPPYCost)
-- Same date logic as JobCosts CTE but cutoff = @PriorYrEnd
-- =============================================================================
PriorYearJobCosts AS (
    SELECT d.JCCo, d.Job,
        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF')
                      AND d.Mth <= @PriorYrEnd THEN d.ActualCost ELSE 0 END)
        AS PriorYrJTDCost
    FROM bJCCD d WITH (NOLOCK)
    WHERE d.JCCo = @Co
    GROUP BY d.JCCo, d.Job
),

-- =============================================================================
-- CTE 13: Prior-year projected cost as of Dec 31 (Bug 4 — denominator for prior-year %)
-- Most recent bJCOP row with udPlugged='Y' at or before Dec 31 prior year
-- =============================================================================
PriorYearProjCost AS (
    SELECT op.JCCo, op.Job,
           ISNULL(op.ProjCost, 0) AS PriorYrProjCost
    FROM (
        SELECT JCCo, Job, ProjCost,
               ROW_NUMBER() OVER (PARTITION BY JCCo, Job ORDER BY Month DESC) AS rn
        FROM bJCOP WITH (NOLOCK)
        WHERE JCCo = @Co AND udPlugged = 'Y' AND Month <= @PriorYrEnd
    ) op WHERE op.rn = 1
),

-- =============================================================================
-- CTE 14: Prior Month JTD Actual Cost (for prior month recognized profit calc)
-- Used by VBA MergePriorMonthProfitsOntoSheet to compute:
--   PriorEarnedRev = OpsRevOverride(prior) * (PriorMonthJTDCost / OpsCostOverride(prior))
--   PriorRecognizedProfit = PriorEarnedRev - PriorMonthJTDCost
-- Nicole rule (2026-04-06): prior profit = recognized (earned rev - cost), not projected
-- =============================================================================
PriorMonthCosts AS (
    SELECT d.JCCo, d.Job,
        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth < @Month
                 THEN d.ActualCost ELSE 0 END) AS PriorMonthJTDCost
    FROM bJCCD d WITH (NOLOCK)
    WHERE d.JCCo = @Co
    GROUP BY d.JCCo, d.Job
)

-- =============================================================================
-- FINAL SELECT: Join everything together, return one row per job
-- =============================================================================
SELECT
    jl.JCCo,
    LTRIM(RTRIM(jl.Job)) AS Job,        -- Trim only in output
    RTRIM(jl.Contract) AS Contract,      -- Trim only in output
    jl.ContractDescription,
    jl.JobDescription,
    jl.PM,
    jl.Department,
    jl.DeptDescription,
    jl.JobStatus,
    -- ContractStatus mapped to 1 (open) or 2 (closed).
    -- VBA only handles 1 and 2; raw Vista values (1=Open,2=SoftClosed,3=HardClosed)
    -- cause the VBA subtotal logic to misfire on the 2→3 transition.
    -- Nicole confirmed (2026-04-08): closed jobs go in the Closed section regardless
    -- of whether MonthClosed is in the batch month.
    CASE WHEN jl.ContractStatus = 1 THEN 1 ELSE 2 END AS ContractStatus,
    jl.MonthClosed,
    -- A8: Vista closure state as of the WIP batch month.
    -- 1 = Vista hard-closed this job BEFORE the batch month (truly closed at WIP date).
    --     NULL MonthClosed with ContractStatus=2 treated as closed (overhead/overhead-type jobs).
    -- 0 = Open, soft-closed, or hard-closed during/after the batch month (still open at WIP date).
    -- VBA compares this against workbook col G (Close flag from LylesWIP) to surface mismatches.
    -- Original bug: job 54.9033 flagged Close='Y' in workbook but ContractStatus=1 (Open) in Vista.
    CASE
        WHEN jl.ContractStatus = 2 AND (jl.MonthClosed IS NULL OR jl.MonthClosed < @Month)
        THEN 1
        ELSE 0
    END AS VistaClosedAtWipDate,
    jl.CompletionDate,
    jl.ProjectMgr,
    jl.ProjMinPct,

    -- Contract amounts
    jl.OrigContractAmt,
    ISNULL(co.COContractAmt, 0) AS COContractAmt,
    jl.OrigContractAmt + ISNULL(co.COContractAmt, 0) AS CurrentContractAmt,
    jl.ContractAmt AS ProjContract,  -- Projected Revenue (from contract)
    ISNULL(bt.BilledAmt, 0) AS BilledAmt,  -- A4: date-filtered from bARTH (not live bJCCM.BilledAmt)
    jl.ReceivedAmt,
    jl.CurrentRetainAmt,

    -- Actual costs
    ISNULL(jc.ActualCost, 0) AS ActualCost,
    ISNULL(jc.CYActualCost, 0) AS CYActualCost,
    ISNULL(jc.PeriodActualCost, 0) AS PeriodActualCost,

    -- Estimates
    ISNULL(jc.OrigEstCost, 0) AS OrigEstCost,
    ISNULL(jc.COEstCost, 0) AS COEstCost,
    ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0) AS CurrentEstimate,

    -- Projected cost (from PF transactions)
    CASE WHEN ISNULL(jc.ProjFinalCost, 0) = 0
         THEN ISNULL(jc.ActualCost, 0)
         ELSE jc.ProjFinalCost END AS ProjCost,
    ISNULL(jc.PriorProjCost, 0) AS PriorProjCost,
    ISNULL(jc.MarchProjCost, 0) AS MarchProjCost,

    -- Cost breakdowns
    ISNULL(jc.JTDLaborCost, 0) AS JTDLaborCost,
    ISNULL(jc.JTDEquipCost, 0) AS JTDEquipCost,
    ISNULL(jc.JTDSubCost, 0) AS JTDSubCost,
    ISNULL(jc.JTDMatlOtherCost, 0) AS JTDMatlOtherCost,

    -- Projection info
    ISNULL(pp.HasProjPlug, 0) AS HasProjPlug,
    ISNULL(pi.DollarsPlug, 0) AS DollarsPlug,
    ISNULL(pi.UnitsPlug, 0) AS UnitsPlug,
    ISNULL(pi.myNotes, '') AS ProjNotes,
    ISNULL(pi.RemCostPerUnit, 0) AS RemCostPerUnit,
    ISNULL(pi.RCPUPlug, 0) AS RCPUPlug,

    -- March baseline (for Prior Projected Profit)
    ISNULL(mb.MarchProjRevenue, 0) AS MarchProjRevenue,

    -- =========================================================================
    -- Calculated WIP fields (business rules applied in SQL)
    -- =========================================================================

    -- % Complete
    CASE
        WHEN ISNULL(jc.ProjFinalCost, 0) <> 0
            THEN ISNULL(jc.ActualCost, 0) * 1.0 / jc.ProjFinalCost
        WHEN (ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0)) <> 0
            THEN ISNULL(jc.ActualCost, 0) * 1.0 / (ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0))
        ELSE 0
    END AS PctComplete,

    -- Earned Revenue (10% GAAP threshold applied)
    CASE
        WHEN CASE
                WHEN ISNULL(jc.ProjFinalCost, 0) <> 0
                    THEN ISNULL(jc.ActualCost, 0) * 1.0 / jc.ProjFinalCost
                WHEN (ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0)) <> 0
                    THEN ISNULL(jc.ActualCost, 0) * 1.0 / (ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0))
                ELSE 0
             END < 0.10
            THEN ISNULL(jc.ActualCost, 0)  -- Below 10%: earned rev = cost
        ELSE jl.ContractAmt * CASE
                WHEN ISNULL(jc.ProjFinalCost, 0) <> 0
                    THEN ISNULL(jc.ActualCost, 0) * 1.0 / jc.ProjFinalCost
                WHEN (ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0)) <> 0
                    THEN ISNULL(jc.ActualCost, 0) * 1.0 / (ISNULL(jc.OrigEstCost, 0) + ISNULL(jc.COEstCost, 0))
                ELSE 0
             END
    END AS EarnedRevenue,

    -- Projected Profit = ContractAmt - ProjCost
    jl.ContractAmt - CASE
        WHEN ISNULL(jc.ProjFinalCost, 0) = 0 THEN ISNULL(jc.ActualCost, 0)
        ELSE jc.ProjFinalCost
    END AS ProjProfit,

    -- Prior Projected Profit (March WIP plug)
    -- = MarchProjRevenue - MarchProjCost
    ISNULL(mb.MarchProjRevenue, 0) - ISNULL(jc.MarchProjCost, 0) AS PriorProjProfit,

    -- Change in Projected Profit = Current ProjProfit - Prior ProjProfit
    (jl.ContractAmt - CASE
        WHEN ISNULL(jc.ProjFinalCost, 0) = 0 THEN ISNULL(jc.ActualCost, 0)
        ELSE jc.ProjFinalCost
    END)
    - (ISNULL(mb.MarchProjRevenue, 0) - ISNULL(jc.MarchProjCost, 0))
    AS ChgProjProfit,

    -- Original values (same as Vista values since no overrides exist yet)
    -- These populate the COLZORG* hidden columns used by LCGWIPUpdateRow concurrency checks
    ISNULL(jc.ActualCost, 0) AS OrgActualCost,
    ISNULL(jc.CYActualCost, 0) AS OrgCYActualCost,
    ISNULL(bt.BilledAmt, 0) AS OrgBilledAmt,
    0 AS OrgCYBilledAmt,       -- CY billings need bJCBT; not used in visible calc, stays 0
    '' AS [Close],
    '' AS Completed,
    '' AS CompletedGAAP,
    '' AS UserName,
    0 AS BatchSeq,
    CAST(0 AS varbinary(8)) AS RowVersion,

    -- Override fields (budWIPDetail is empty; all zero until Sprint 2 batch load)
    0 AS OpsRev, '' AS OpsRevPlugged, 0 AS OpsCost, '' AS OpsCostPlugged,
    0 AS GAAPRev, '' AS GAAPRevPlugged, 0 AS GAAPCost, '' AS GAAPCostPlugged,
    0 AS BonusProfit, '' AS BonusProfitPlugged, '' AS BonusProfitNotes,
    0 AS PriorYrBonusProfit,

    -- Trend data (6-month history — requires prior WIP batches; zero until Sprint 2)
    0 AS LastProjContract, 0 AS LastProjContract2, 0 AS LastProjContract3,
    0 AS LastProjContract4, 0 AS LastProjContract5, 0 AS LastProjContract6,
    0 AS LastProjCost, 0 AS LastProjCost2, 0 AS LastProjCost3,
    0 AS LastProjCost4, 0 AS LastProjCost5, 0 AS LastProjCost6,
    0 AS LastOpsRev, 0 AS LastOpsRev2, 0 AS LastOpsRev3,
    0 AS LastOpsRev4, 0 AS LastOpsRev5, 0 AS LastOpsRev6,
    '' AS LastOpsRevPlugged, '' AS LastOpsRevPlugged2, '' AS LastOpsRevPlugged3,
    '' AS LastOpsRevPlugged4, '' AS LastOpsRevPlugged5, '' AS LastOpsRevPlugged6,
    0 AS LastOpsCost, 0 AS LastOpsCost2, 0 AS LastOpsCost3,
    0 AS LastOpsCost4, 0 AS LastOpsCost5, 0 AS LastOpsCost6,
    '' AS LastOpsCostPlugged, '' AS LastOpsCostPlugged2, '' AS LastOpsCostPlugged3,
    '' AS LastOpsCostPlugged4, '' AS LastOpsCostPlugged5, '' AS LastOpsCostPlugged6,

    -- Prior-quarter GAAP snapshot (Bug 3 — feeds COLZPriorJTDGAAPProfit)
    ISNULL(pgr.LastGAAPRev, 0) AS LastGAAPRev,
    0 AS LastGAAPRev2, 0 AS LastGAAPRev3, 0 AS LastGAAPRev4, 0 AS LastGAAPRev5, 0 AS LastGAAPRev6,
    ISNULL(pgr.LastGAAPRevPlugged, '') AS LastGAAPRevPlugged,
    '' AS LastGAAPRevPlugged2, '' AS LastGAAPRevPlugged3,
    '' AS LastGAAPRevPlugged4, '' AS LastGAAPRevPlugged5, '' AS LastGAAPRevPlugged6,
    ISNULL(pgc.LastGAAPCost, 0) AS LastGAAPCost,
    0 AS LastGAAPCost2, 0 AS LastGAAPCost3, 0 AS LastGAAPCost4, 0 AS LastGAAPCost5, 0 AS LastGAAPCost6,
    ISNULL(pgc.LastGAAPCostPlugged, '') AS LastGAAPCostPlugged,
    '' AS LastGAAPCostPlugged2, '' AS LastGAAPCostPlugged3,
    '' AS LastGAAPCostPlugged4, '' AS LastGAAPCostPlugged5, '' AS LastGAAPCostPlugged6,
    0 AS LastBonusProfit,
    ISNULL(pmc.PriorMonthJTDCost, 0) AS LastActualCost,

    -- Prior-year revenue/cost (Bug 4 — feeds COLZGAAPPYRev/Cost, COLZOpsPYRev/Cost)
    -- Applies 10% GAAP threshold at Dec 31 prior year
    CASE
        WHEN ISNULL(pyJOP.PriorYrProjCost, 0) > 0
             AND ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost >= 0.10
        THEN ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost * jl.ContractAmt
        ELSE ISNULL(pyJC.PriorYrJTDCost, 0)
    END AS PriorYearGAAPRev,
    ISNULL(pyJC.PriorYrJTDCost, 0) AS PriorYearGAAPCost,
    -- Applies 30% OPS threshold at Dec 31 prior year
    CASE
        WHEN ISNULL(pyJOP.PriorYrProjCost, 0) > 0
             AND ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost >= 0.30
        THEN ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost * jl.ContractAmt
        ELSE ISNULL(pyJC.PriorYrJTDCost, 0)
    END AS PriorYearOpsRev,
    ISNULL(pyJC.PriorYrJTDCost, 0) AS PriorYearOpsCost,

    -- Notes (from budProjInfo or budWIPDetail — empty until Sprint 2)
    '' AS OpsRevNotes, '' AS OpsCostNotes,
    '' AS GAAPRevNotes, '' AS GAAPCostNotes

FROM JobList jl
LEFT JOIN JobCosts jc ON jl.JCCo = jc.JCCo AND jl.Job = jc.Job
LEFT JOIN COContracts co ON jl.JCCo = co.JCCo AND jl.Job = co.Job
LEFT JOIN BilledThruMonth bt ON jl.JCCo = bt.JCCo AND jl.Contract = bt.Contract
LEFT JOIN ProjPlugs pp ON jl.JCCo = pp.JCCo AND jl.Job = pp.Job
LEFT JOIN ProjInfo pi ON jl.JCCo = pi.Co AND jl.Job = pi.Job
LEFT JOIN MarchBaseline mb ON jl.JCCo = mb.JCCo AND jl.Job = mb.Job
LEFT JOIN PriorGAAPRev pgr ON jl.JCCo = pgr.JCCo AND jl.Job = pgr.Job
LEFT JOIN PriorGAAPCost pgc ON jl.JCCo = pgc.JCCo AND jl.Job = pgc.Job
LEFT JOIN PriorYearJobCosts pyJC ON jl.JCCo = pyJC.JCCo AND jl.Job = pyJC.Job
LEFT JOIN PriorYearProjCost pyJOP ON jl.JCCo = pyJOP.JCCo AND jl.Job = pyJOP.Job
LEFT JOIN PriorMonthCosts pmc ON jl.JCCo = pmc.JCCo AND jl.Job = pmc.Job

-- Exclude zero-activity zero-contract jobs (Michael's LCGWIPCreateBatch final WHERE).
-- Softened: include jobs with estimates (OE/CO/PF) even if no actual cost yet.
-- This retains overhead/cost-only jobs (e.g. 51.1156, 51.1157) that have projected
-- cost but zero contract. Only exclude truly empty shells.
WHERE NOT (ISNULL(jc.ActualCost, 0) = 0
       AND ISNULL(jc.OrigEstCost, 0) = 0
       AND ISNULL(jc.ProjFinalCost, 0) = 0
       AND (jl.OrigContractAmt + ISNULL(co.COContractAmt, 0)) = 0)

ORDER BY
    CASE WHEN @GroupBy = 'Department' THEN jl.Department ELSE jl.PM END,
    CASE WHEN jl.ContractStatus = 1 THEN 1 ELSE 2 END,
    jl.Contract;
