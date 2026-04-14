-- =============================================================================
-- WipDb vs Vista Comparison — December 2025, Company 15 (WML)
-- Run against: 10.112.11.8 (Viewpoint), Database: Viewpoint
-- Auth: WIPexcel / current password
-- Purpose: Explain to Nicole/Cindy what changed between Rev 5.47p and Rev 5.53p
-- =============================================================================

DECLARE @Co      tinyint = 15;
DECLARE @Month   date    = '2025-12-01';
DECLARE @EndDate date    = '2025-12-31';   -- cutoff for JTD cost
DECLARE @StartCY date    = '2025-01-01';   -- current year start

-- =============================================================================
-- QUERY 1: Are Nicole's Dec 2025 overrides already loaded into Viewpoint?
-- If Michael loaded her import file, these should have rows.
-- =============================================================================
SELECT 'REVENUE OVERRIDES (bJCOR) — Dec 2025' AS Source,
       COUNT(*)                          AS JobCount,
       SUM(r.RevCost)                    AS GAAP_Rev_Override,
       SUM(r.OtherAmount)                AS Ops_Rev_Override,
       SUM(r.RevCost - r.OtherAmount)    AS GAAP_Minus_Ops_Rev
FROM   bJCOR r WITH (NOLOCK)
WHERE  r.JCCo = @Co
  AND  r.Month = @Month
  AND  r.udPlugged = 'Y';

SELECT 'COST OVERRIDES (bJCOP) — Dec 2025' AS Source,
       COUNT(*)                            AS JobCount,
       SUM(p.ProjCost)                     AS GAAP_Cost_Override,
       SUM(p.OtherAmount)                  AS Ops_Cost_Override,
       SUM(p.ProjCost - p.OtherAmount)     AS GAAP_Minus_Ops_Cost
FROM   bJCOP p WITH (NOLOCK)
WHERE  p.JCCo = @Co
  AND  p.Month = @Month
  AND  p.udPlugged = 'Y';
GO

-- =============================================================================
-- QUERY 2: What the old system (Michael's WipDb path) had approved through
-- the most recently finalized month — from LCGWIPSchedule.
-- This is the benchmark Nicole was comparing against.
-- =============================================================================
SELECT TOP 20
    s.Co,
    s.Mth,
    LTRIM(RTRIM(s.Job))                         AS Job,
    s.JobDesc,
    s.Dept,
    s.OpsEarnedRev                              AS Approved_Ops_EarnedRev,
    s.GAAPEarnedRev                             AS Approved_GAAP_EarnedRev,
    s.OpsProjCost                               AS Approved_Ops_ProjCost,
    s.GAAPProjCost                              AS Approved_GAAP_ProjCost,
    s.OpsEarnedRev - s.OpsProjCost              AS Approved_Ops_Profit,
    s.GAAPEarnedRev - s.GAAPProjCost            AS Approved_GAAP_Profit,
    s.ReadyForOps,
    s.FinalApproval,
    s.AcctApproval
FROM dbo.LCGWIPSchedule s WITH (NOLOCK)
WHERE s.Co = 15
  AND s.Mth = '2025-11-01'   -- most recent finalized month
  AND s.Dept = '54'
ORDER BY LTRIM(RTRIM(s.Job));
GO

-- Totals for last finalized month (all depts, all jobs)
SELECT
    s.Co,
    s.Mth,
    s.Dept,
    COUNT(*)                            AS JobCount,
    SUM(s.OpsEarnedRev)                 AS Total_Ops_EarnedRev,
    SUM(s.GAAPEarnedRev)                AS Total_GAAP_EarnedRev,
    SUM(s.OpsProjCost)                  AS Total_Ops_ProjCost,
    SUM(s.GAAPProjCost)                 AS Total_GAAP_ProjCost,
    SUM(s.OpsEarnedRev - s.OpsProjCost) AS Total_Ops_Profit,
    SUM(s.GAAPEarnedRev - s.GAAPProjCost) AS Total_GAAP_Profit
FROM dbo.LCGWIPSchedule s WITH (NOLOCK)
WHERE s.Co = 15
  AND s.Mth = '2025-11-01'
GROUP BY s.Co, s.Mth, s.Dept
ORDER BY s.Dept;
GO

-- =============================================================================
-- QUERY 3: What our new workbook shows — Vista formula-calculated values
-- (no overrides applied — this is what 5.53p shows right now for Dec 2025)
-- Dept 54, Company 15 for direct comparison with smoke-test results
-- =============================================================================
DECLARE @Co      tinyint = 15;
DECLARE @Month   date    = '2025-12-01';
DECLARE @EndDate date    = '2025-12-31';
DECLARE @StartCY date    = '2025-01-01';
DECLARE @OpsThreshold  decimal(5,2) = 0.30;
DECLARE @GAAPThreshold decimal(5,2) = 0.10;

WITH JobBase AS (
    SELECT
        j.JCCo,
        j.Job,
        c.Contract,
        c.Description,
        c.Department,
        c.ContractAmt,
        c.BilledAmt,
        j.JobStatus
    FROM bJCJM j WITH (NOLOCK)
    JOIN bJCCM c WITH (NOLOCK) ON j.JCCo = c.JCCo AND j.Contract = c.Contract
    WHERE j.JCCo = @Co
      AND c.Department = '54'    -- ← change dept here
),
Costs AS (
    SELECT
        d.JCCo,
        d.Job,
        SUM(CASE WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF')
                  AND d.PostedDate <= @EndDate THEN d.ActualCost
             WHEN d.JCTransType IN ('PR','EM')
                  AND d.ActualDate <= @EndDate THEN d.ActualCost
             ELSE 0 END)                                AS JTDCost,
        SUM(CASE WHEN d.JCTransType = 'PF'
                  AND d.PostedDate <= @EndDate THEN d.ProjCost ELSE 0 END) AS ProjFinalCost
    FROM bJCCD d WITH (NOLOCK)
    WHERE d.JCCo = @Co
    GROUP BY d.JCCo, d.Job
),
RevenueOverride AS (
    -- Current month override revenue (if Nicole/Michael loaded Dec 2025)
    SELECT r.JCCo, r.Contract,
           r.RevCost   AS GAAP_Rev_Override,
           r.OtherAmount AS Ops_Rev_Override
    FROM bJCOR r WITH (NOLOCK)
    WHERE r.JCCo = @Co
      AND r.Month = @Month
      AND r.udPlugged = 'Y'
),
CostOverride AS (
    -- Current month override cost (if Nicole/Michael loaded Dec 2025)
    SELECT p.JCCo, p.Job,
           p.ProjCost    AS GAAP_Cost_Override,
           p.OtherAmount AS Ops_Cost_Override
    FROM bJCOP p WITH (NOLOCK)
    WHERE p.JCCo = @Co
      AND p.Month = @Month
      AND p.udPlugged = 'Y'
)
SELECT
    j.Department                                    AS Dept,
    LTRIM(RTRIM(j.Job))                             AS Job,
    j.Description,
    j.ContractAmt                                   AS ContractAmt,
    ISNULL(c.JTDCost, 0)                            AS JTD_Cost,
    ISNULL(c.ProjFinalCost, 0)                      AS ProjFinalCost_Vista,
    -- % Complete (formula)
    CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 THEN 0
         ELSE ISNULL(c.JTDCost,0) * 1.0 / c.ProjFinalCost END AS PctComplete,
    -- Ops Earned Rev — formula (no override)
    CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
              ISNULL(c.JTDCost,0) * 1.0 / NULLIF(c.ProjFinalCost,0) < @OpsThreshold
         THEN ISNULL(c.JTDCost, 0)
         ELSE j.ContractAmt * (ISNULL(c.JTDCost,0) * 1.0 / c.ProjFinalCost)
    END                                             AS Ops_EarnedRev_Formula,
    -- GAAP Earned Rev — formula (no override)
    CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
              ISNULL(c.JTDCost,0) * 1.0 / NULLIF(c.ProjFinalCost,0) < @GAAPThreshold
         THEN ISNULL(c.JTDCost, 0)
         ELSE j.ContractAmt * (ISNULL(c.JTDCost,0) * 1.0 / c.ProjFinalCost)
    END                                             AS GAAP_EarnedRev_Formula,
    -- Nicole's override values (from bJCOR/bJCOP — what Michael would have loaded)
    ro.GAAP_Rev_Override,
    ro.Ops_Rev_Override,
    co2.GAAP_Cost_Override,
    co2.Ops_Cost_Override,
    -- Delta: Override vs Formula (the difference Nicole will see between 5.47p and 5.53p)
    ro.Ops_Rev_Override - CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
              ISNULL(c.JTDCost,0) * 1.0 / NULLIF(c.ProjFinalCost,0) < @OpsThreshold
         THEN ISNULL(c.JTDCost, 0)
         ELSE j.ContractAmt * (ISNULL(c.JTDCost,0) * 1.0 / c.ProjFinalCost)
    END                                             AS Ops_Rev_Delta_Override_vs_Formula
FROM   JobBase j
LEFT   JOIN Costs c          ON j.JCCo = c.JCCo AND j.Job = c.Job
LEFT   JOIN RevenueOverride ro ON j.JCCo = ro.JCCo AND j.Contract = ro.Contract
LEFT   JOIN CostOverride co2  ON j.JCCo = co2.JCCo AND j.Job = co2.Job
WHERE  j.JobStatus IN (1, 2, 3)    -- Active, Inactive, Closed
ORDER  BY LTRIM(RTRIM(j.Job));
GO

-- =============================================================================
-- QUERY 4: Summary delta — how far off is our formula vs Nicole's overrides?
-- Run this first to get the headline numbers for the meeting.
-- =============================================================================
DECLARE @Co      tinyint = 15;
DECLARE @Month   date    = '2025-12-01';
DECLARE @EndDate date    = '2025-12-31';

WITH Costs AS (
    SELECT d.JCCo, d.Job,
        SUM(CASE WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF')
                  AND d.PostedDate <= @EndDate THEN d.ActualCost
             WHEN d.JCTransType IN ('PR','EM')
                  AND d.ActualDate <= @EndDate THEN d.ActualCost
             ELSE 0 END) AS JTDCost,
        SUM(CASE WHEN d.JCTransType = 'PF'
                  AND d.PostedDate <= @EndDate THEN d.ProjCost ELSE 0 END) AS ProjFinalCost
    FROM bJCCD d WITH (NOLOCK)
    WHERE d.JCCo = @Co
    GROUP BY d.JCCo, d.Job
)
SELECT
    'GAAP' AS Basis,
    COUNT(j.Job)                                          AS Jobs,
    SUM(r.RevCost)                                        AS Override_Rev,
    SUM(CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
                  ISNULL(c.JTDCost,0)*1.0/NULLIF(c.ProjFinalCost,0) < 0.10
             THEN ISNULL(c.JTDCost,0)
             ELSE cm.ContractAmt * (ISNULL(c.JTDCost,0)*1.0/c.ProjFinalCost) END) AS Formula_Rev,
    SUM(r.RevCost) - SUM(
        CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
                  ISNULL(c.JTDCost,0)*1.0/NULLIF(c.ProjFinalCost,0) < 0.10
             THEN ISNULL(c.JTDCost,0)
             ELSE cm.ContractAmt * (ISNULL(c.JTDCost,0)*1.0/c.ProjFinalCost) END) AS Rev_Delta,
    SUM(p.ProjCost)                                       AS Override_Cost,
    SUM(ISNULL(c.ProjFinalCost,0))                        AS Formula_Cost,
    SUM(p.ProjCost) - SUM(ISNULL(c.ProjFinalCost,0))     AS Cost_Delta
FROM bJCJM j WITH (NOLOCK)
JOIN bJCCM cm WITH (NOLOCK) ON j.JCCo = cm.JCCo AND j.Contract = cm.Contract
JOIN bJCOR r  WITH (NOLOCK) ON j.JCCo = r.JCCo  AND j.Contract = r.Contract
                            AND r.Month = @Month AND r.udPlugged = 'Y'
JOIN bJCOP p  WITH (NOLOCK) ON j.JCCo = p.JCCo  AND j.Job = p.Job
                            AND p.Month = @Month AND p.udPlugged = 'Y'
LEFT JOIN Costs c ON j.JCCo = c.JCCo AND j.Job = c.Job
WHERE j.JCCo = @Co

UNION ALL

SELECT
    'OPS' AS Basis,
    COUNT(j.Job),
    SUM(r.OtherAmount),
    SUM(CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
                  ISNULL(c.JTDCost,0)*1.0/NULLIF(c.ProjFinalCost,0) < 0.30
             THEN ISNULL(c.JTDCost,0)
             ELSE cm.ContractAmt * (ISNULL(c.JTDCost,0)*1.0/c.ProjFinalCost) END),
    SUM(r.OtherAmount) - SUM(
        CASE WHEN ISNULL(c.ProjFinalCost,0) = 0 OR
                  ISNULL(c.JTDCost,0)*1.0/NULLIF(c.ProjFinalCost,0) < 0.30
             THEN ISNULL(c.JTDCost,0)
             ELSE cm.ContractAmt * (ISNULL(c.JTDCost,0)*1.0/c.ProjFinalCost) END),
    SUM(p.OtherAmount),
    SUM(ISNULL(c.ProjFinalCost,0)),
    SUM(p.OtherAmount) - SUM(ISNULL(c.ProjFinalCost,0))
FROM bJCJM j WITH (NOLOCK)
JOIN bJCCM cm WITH (NOLOCK) ON j.JCCo = cm.JCCo AND j.Contract = cm.Contract
JOIN bJCOR r  WITH (NOLOCK) ON j.JCCo = r.JCCo  AND j.Contract = r.Contract
                            AND r.Month = @Month AND r.udPlugged = 'Y'
JOIN bJCOP p  WITH (NOLOCK) ON j.JCCo = p.JCCo  AND j.Job = p.Job
                            AND p.Month = @Month AND p.udPlugged = 'Y'
LEFT JOIN Costs c ON j.JCCo = c.JCCo AND j.Job = c.Job
WHERE j.JCCo = @Co;
GO
