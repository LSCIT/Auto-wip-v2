-- =============================================================================
-- WIP Vista Direct — Validation Test Queries
-- Run against: 10.112.11.8 (Viewpoint test server)
-- Auth: WIPexcel / WIP@MR@2024
-- =============================================================================

-- =============================================================================
-- TEST 0: Connectivity Test
-- =============================================================================
SELECT @@SERVERNAME AS ServerName, DB_NAME() AS DatabaseName;
GO

-- =============================================================================
-- TEST 1: Nicole's Discrepancy #1 — Job 54.9033 Status
-- Expected: Should show current live status from Vista, not stale WipDb cache
-- =============================================================================
SELECT
    LTRIM(RTRIM(j.Job)) AS Job,
    j.JobStatus,
    CASE j.JobStatus
        WHEN 1 THEN 'Active'
        WHEN 2 THEN 'Inactive'
        WHEN 3 THEN 'Closed'
        ELSE 'Unknown(' + CAST(j.JobStatus AS varchar) + ')'
    END AS StatusDescription,
    j.Description
FROM bJCJM j WITH (NOLOCK)
WHERE j.JCCo = 15 AND LTRIM(RTRIM(j.Job)) = '54.9033.';
GO

-- =============================================================================
-- TEST 2: Nicole's Discrepancy #2 — Missing Cross-Year Reversal Jobs
-- Jobs 56.1022 and 56.1057 should appear even with JTD=$0
-- These have current-year reversals of prior-year costs
-- =============================================================================

-- Check if these jobs exist in Vista
SELECT LTRIM(RTRIM(j.Job)) AS Job, j.JobStatus, j.Description
FROM bJCJM j WITH (NOLOCK)
WHERE j.JCCo = 15 AND LTRIM(RTRIM(j.Job)) IN ('56.1022.', '56.1057.');

-- Check for current year activity (should have reversal transactions)
SELECT
    LTRIM(RTRIM(d.Job)) AS Job,
    d.JCTransType,
    d.PostedDate,
    d.ActualDate,
    d.ActualCost,
    d.Description
FROM bJCCD d WITH (NOLOCK)
WHERE d.JCCo = 15
  AND LTRIM(RTRIM(d.Job)) IN ('56.1022.', '56.1057.')
  AND (
    (d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate BETWEEN '2025-01-01' AND '2025-10-31')
    OR (d.JCTransType IN ('PR','EM') AND d.ActualDate BETWEEN '2025-01-01' AND '2025-10-31')
  )
ORDER BY d.Job, d.PostedDate;
GO

-- =============================================================================
-- TEST 3: Nicole's Discrepancy #3 — Prior Projected Profit (Col R)
-- Job 51.1151: Expected $596,848 (from March WIP plug)
-- Prior Projected Profit = MarchProjRevenue - MarchProjCost
-- =============================================================================

-- March baseline projected revenue (from bJCOR)
SELECT TOP 1
    jor.JCCo,
    LTRIM(RTRIM(j.Job)) AS Job,
    jor.Month,
    jor.OtherAmount AS MarchProjRevenue
FROM bJCOR jor WITH (NOLOCK)
JOIN bJCJM j WITH (NOLOCK) ON jor.JCCo = j.JCCo AND jor.Contract = j.Contract
WHERE j.JCCo = 15 AND LTRIM(RTRIM(j.Job)) = '51.1151.'
  AND jor.Month <= '2025-03-31'
ORDER BY jor.Month DESC;

-- March baseline projected cost (from bJCCD PF transactions as of March)
SELECT
    d.JCCo,
    LTRIM(RTRIM(d.Job)) AS Job,
    SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= '2025-03-31' THEN d.ProjCost ELSE 0 END) AS MarchProjCost
FROM bJCCD d WITH (NOLOCK)
WHERE d.JCCo = 15 AND LTRIM(RTRIM(d.Job)) = '51.1151.'
GROUP BY d.JCCo, LTRIM(RTRIM(d.Job));
-- Prior Projected Profit = MarchProjRevenue - MarchProjCost
-- Should = $596,848
GO

-- =============================================================================
-- TEST 4: Nicole's Discrepancy #4 — JTD Earned Revenue (Col W)
-- Job 51.1151: Expected $656,956 (= cost, because <10% complete)
-- Current tool shows $695,927 (wrong — used percentage calc ignoring 10% threshold)
-- =============================================================================

-- Get actual cost and projected cost for 51.1151
SELECT
    LTRIM(RTRIM(d.Job)) AS Job,
    SUM(CASE
        WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate <= '2025-10-31'
            THEN d.ActualCost ELSE 0 END)
    + SUM(CASE
        WHEN d.JCTransType IN ('PR','EM') AND d.ActualDate <= '2025-10-31'
            THEN d.ActualCost ELSE 0 END)
    AS JTDActualCost,
    SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= '2025-10-31'
        THEN d.ProjCost ELSE 0 END) AS ProjFinalCost
FROM bJCCD d WITH (NOLOCK)
WHERE d.JCCo = 15 AND LTRIM(RTRIM(d.Job)) = '51.1151.'
GROUP BY LTRIM(RTRIM(d.Job));

-- Then calculate:
-- PctComplete = ActualCost / ProjFinalCost
-- If PctComplete < 0.10:
--   EarnedRevenue = ActualCost  (should be $656,956)
-- Else:
--   EarnedRevenue = ContractAmt * PctComplete

-- Get contract amount
SELECT c.ContractAmt, c.OrigContractAmt
FROM bJCCM c WITH (NOLOCK)
JOIN bJCJM j WITH (NOLOCK) ON c.JCCo = j.JCCo AND c.Contract = j.Contract
WHERE j.JCCo = 15 AND LTRIM(RTRIM(j.Job)) = '51.1151.';
GO

-- =============================================================================
-- TEST 5: Full WIP Query for Job 51.1151 (combined validation)
-- This runs the main Vista query filtered to just this one job
-- =============================================================================
DECLARE @Co tinyint = 15;
DECLARE @CutOffDate date = '2025-10-31';
DECLARE @StartDate date = '2025-01-01';
DECLARE @PriorYrEnd date = '2024-12-31';
DECLARE @CurrentDate date = GETDATE();
DECLARE @MarchPlug date = '2025-03-31';

SELECT
    LTRIM(RTRIM(j.Job)) AS Job,
    c.Description AS ContractDescription,
    c.ContractAmt,
    c.OrigContractAmt,
    c.BilledAmt,
    j.JobStatus,
    c.ContractStatus,

    -- Actual cost
    ISNULL((
        SELECT SUM(CASE
            WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate <= @CutOffDate THEN d.ActualCost ELSE 0 END)
        + SUM(CASE
            WHEN d.JCTransType IN ('PR','EM') AND d.ActualDate <= @CutOffDate THEN d.ActualCost ELSE 0 END)
        FROM bJCCD d WITH (NOLOCK)
        WHERE d.JCCo = j.JCCo AND d.Job = j.Job
    ), 0) AS JTDActualCost,

    -- Projected cost
    ISNULL((
        SELECT SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= @CutOffDate THEN d.ProjCost ELSE 0 END)
        FROM bJCCD d WITH (NOLOCK)
        WHERE d.JCCo = j.JCCo AND d.Job = j.Job
    ), 0) AS ProjFinalCost,

    -- % Complete
    CASE WHEN ISNULL((
        SELECT SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= @CutOffDate THEN d.ProjCost ELSE 0 END)
        FROM bJCCD d WITH (NOLOCK)
        WHERE d.JCCo = j.JCCo AND d.Job = j.Job
    ), 0) <> 0
    THEN ISNULL((
        SELECT SUM(CASE
            WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate <= @CutOffDate THEN d.ActualCost ELSE 0 END)
        + SUM(CASE
            WHEN d.JCTransType IN ('PR','EM') AND d.ActualDate <= @CutOffDate THEN d.ActualCost ELSE 0 END)
        FROM bJCCD d WITH (NOLOCK)
        WHERE d.JCCo = j.JCCo AND d.Job = j.Job
    ), 0) * 1.0 / NULLIF((
        SELECT SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= @CutOffDate THEN d.ProjCost ELSE 0 END)
        FROM bJCCD d WITH (NOLOCK)
        WHERE d.JCCo = j.JCCo AND d.Job = j.Job
    ), 0)
    ELSE 0 END AS PctComplete,

    -- March projected cost (for Prior Projected Profit)
    ISNULL((
        SELECT SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= @MarchPlug THEN d.ProjCost ELSE 0 END)
        FROM bJCCD d WITH (NOLOCK)
        WHERE d.JCCo = j.JCCo AND d.Job = j.Job
    ), 0) AS MarchProjCost,

    -- March projected revenue
    ISNULL((
        SELECT TOP 1 jor.OtherAmount
        FROM bJCOR jor WITH (NOLOCK)
        WHERE jor.JCCo = j.JCCo AND jor.Contract = j.Contract AND jor.Month <= @MarchPlug
        ORDER BY jor.Month DESC
    ), 0) AS MarchProjRevenue

FROM bJCJM j WITH (NOLOCK)
JOIN bJCCM c WITH (NOLOCK) ON j.JCCo = c.JCCo AND j.Contract = c.Contract
WHERE j.JCCo = @Co AND LTRIM(RTRIM(j.Job)) = '51.1151.';
GO

-- =============================================================================
-- TEST 6: Job Count Sanity Check
-- Compare number of jobs returned by Vista query vs. what the tool showed
-- =============================================================================
SELECT
    c.Department,
    COUNT(*) AS JobCount,
    SUM(CASE WHEN j.JobStatus IN (1, 3) THEN 1 ELSE 0 END) AS ActiveJobs,
    SUM(CASE WHEN j.JobStatus = 2 THEN 1 ELSE 0 END) AS ClosedJobs
FROM bJCJM j WITH (NOLOCK)
JOIN bJCCM c WITH (NOLOCK) ON j.JCCo = c.JCCo AND j.Contract = c.Contract
JOIN bJCDM d WITH (NOLOCK) ON c.JCCo = d.JCCo AND c.Department = d.Department
WHERE j.JCCo = 15
  AND j.JobStatus IN (1, 2, 3)
GROUP BY c.Department
ORDER BY c.Department;
GO

-- =============================================================================
-- TEST 7: PNP Projection Match
-- Pick 3-5 jobs and compare our projected cost with what PNP would show
-- Both should read the same PF transactions from bJCCD
-- =============================================================================
SELECT TOP 5
    LTRIM(RTRIM(d.Job)) AS Job,
    SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= '2025-10-31' THEN d.ProjCost ELSE 0 END) AS ProjFinalCost,
    SUM(CASE
        WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate <= '2025-10-31' THEN d.ActualCost ELSE 0 END)
    + SUM(CASE
        WHEN d.JCTransType IN ('PR','EM') AND d.ActualDate <= '2025-10-31' THEN d.ActualCost ELSE 0 END)
    AS JTDActualCost
FROM bJCCD d WITH (NOLOCK)
WHERE d.JCCo = 15
GROUP BY LTRIM(RTRIM(d.Job))
HAVING SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= '2025-10-31' THEN d.ProjCost ELSE 0 END) > 0
ORDER BY LTRIM(RTRIM(d.Job));
GO

-- =============================================================================
-- TEST 8: GAAP Quarterly Check
-- Verify the month parameter logic — non-quarterly months should have blank GAAP
-- This is enforced in VBA, not SQL. The query returns data regardless.
-- Quarterly months: 3 (March), 6 (June), 9 (September), 12 (December)
-- =============================================================================
-- VBA check: MONTH(Mo) IN (3, 6, 9, 12)
-- If not quarterly, GAAP sheet should not be populated
SELECT 'This is a VBA-side check, not SQL. Verify in the workbook.' AS Note;
GO
