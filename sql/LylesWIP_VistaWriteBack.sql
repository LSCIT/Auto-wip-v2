-- =============================================================================
-- Vista Write-Back Copy Tables + Stored Procedures
-- Server:  10.103.30.11 (Cloud-Apps1 / P&P server)
-- Database: LylesWIP
-- Purpose: Local copies of Vista bJCOP/bJCOR for testing the write-back flow
--          without touching Vista production. When ready to go live, the VBA
--          write-back button targets Vista instead of these tables.
-- =============================================================================
-- Column mapping (Vista → LylesWIP copy):
--   bJCOP.ProjCost     = GAAP Cost Override
--   bJCOP.OtherAmount  = OPS Cost Override
--   bJCOR.RevCost      = GAAP Revenue Override
--   bJCOR.OtherAmount  = OPS Revenue Override
--   udPlugged = 'Y'    = user manually entered this value
--
-- Key difference: bJCOP is keyed by Job, bJCOR is keyed by Contract.
-- =============================================================================

USE LylesWIP;
GO

-- =============================================================================
-- WipJCOP — mirrors Vista bJCOP (Job Cost Override by Period)
-- =============================================================================
IF OBJECT_ID('dbo.WipJCOP', 'U') IS NOT NULL
    DROP TABLE dbo.WipJCOP;
GO

CREATE TABLE dbo.WipJCOP (
    JCCo            TINYINT         NOT NULL,
    Job             VARCHAR(10)     NOT NULL,   -- Raw Vista job number with trailing dot
    Month           SMALLDATETIME   NOT NULL,   -- First of WIP month
    ProjCost        DECIMAL(12,2)   NOT NULL DEFAULT 0,   -- GAAP Cost Override
    OtherAmount     DECIMAL(12,2)   NOT NULL DEFAULT 0,   -- OPS Cost Override
    Notes           VARCHAR(MAX)    NULL,
    udPlugged       CHAR(1)         NULL DEFAULT 'N',     -- 'Y' = user override
    -- Tracking columns (not in Vista — for our audit trail)
    WrittenBy       VARCHAR(100)    NULL,       -- Windows username or 'AutoWIP'
    WrittenAt       DATETIME        NULL DEFAULT GETDATE(),
    BatchId         INT             NULL,       -- FK to WipBatches.Id (which batch triggered this write)
    CONSTRAINT PK_WipJCOP PRIMARY KEY (JCCo, Job, Month)
);
GO

PRINT 'Created WipJCOP table.';
GO

-- =============================================================================
-- WipJCOR — mirrors Vista bJCOR (Job Cost Override by Revenue / Contract-level)
-- =============================================================================
IF OBJECT_ID('dbo.WipJCOR', 'U') IS NOT NULL
    DROP TABLE dbo.WipJCOR;
GO

CREATE TABLE dbo.WipJCOR (
    JCCo            TINYINT         NOT NULL,
    Contract        VARCHAR(10)     NOT NULL,   -- Contract number (NOT Job)
    Month           SMALLDATETIME   NOT NULL,   -- First of WIP month
    RevCost         DECIMAL(12,2)   NOT NULL DEFAULT 0,   -- GAAP Revenue Override
    OtherAmount     DECIMAL(12,2)   NOT NULL DEFAULT 0,   -- OPS Revenue Override
    Notes           VARCHAR(MAX)    NULL,
    udPlugged       CHAR(1)         NULL DEFAULT 'N',     -- 'Y' = user override
    -- Tracking columns (not in Vista — for our audit trail)
    WrittenBy       VARCHAR(100)    NULL,
    WrittenAt       DATETIME        NULL DEFAULT GETDATE(),
    BatchId         INT             NULL,       -- FK to WipBatches.Id
    CONSTRAINT PK_WipJCOR PRIMARY KEY (JCCo, Contract, Month)
);
GO

PRINT 'Created WipJCOR table.';
GO

-- =============================================================================
-- STORED PROC: LylesWIPWriteBackToVista
-- Two-pass MERGE pattern (matches Michael's LCGWIPMergeDetail approach).
-- Reads approved data from WipJobData, writes to WipJCOP/WipJCOR.
-- When ready for production: change target tables to Vista bJCOP/bJCOR.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPWriteBackToVista', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPWriteBackToVista;
GO

CREATE PROCEDURE dbo.LylesWIPWriteBackToVista
    @Co         TINYINT,
    @Month      DATE,
    @DeptList   VARCHAR(200),   -- Comma-separated dept codes
    @UserName   VARCHAR(100),
    @rcode      INT OUTPUT,
    @msg        VARCHAR(500) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET @rcode = 0;
    SET @msg = '';

    -- =========================================================================
    -- Guard: GAAP quarterly only (Mar, Jun, Sep, Dec)
    -- =========================================================================
    IF MONTH(@Month) % 3 <> 0
    BEGIN
        SET @rcode = 1;
        SET @msg = 'Write-back is only allowed on GAAP quarter months (Mar/Jun/Sep/Dec).';
        RETURN;
    END

    -- =========================================================================
    -- Guard: All departments in scope must be AcctApproved
    -- =========================================================================
    DECLARE @DeptTable TABLE (Dept VARCHAR(10));
    INSERT INTO @DeptTable (Dept)
    SELECT LTRIM(RTRIM(value)) FROM STRING_SPLIT(@DeptList, ',');

    IF EXISTS (
        SELECT 1 FROM dbo.WipBatches b
        JOIN @DeptTable d ON b.Department = d.Dept
        WHERE b.JCCo = @Co AND b.WipMonth = @Month
          AND b.BatchState <> 'AcctApproved'
    )
    BEGIN
        SET @rcode = 2;
        SET @msg = 'All departments must be AcctApproved before writing back to Vista.';
        RETURN;
    END

    -- =========================================================================
    -- Build source data from WipJobData (only plugged overrides)
    -- Job number in WipJobData = both Job (for bJCOP) and Contract (for bJCOR)
    -- since in LCG's structure, Job and Contract are the same value.
    -- =========================================================================

    BEGIN TRY
        BEGIN TRANSACTION;

        -- =====================================================================
        -- WipJCOP (Cost overrides) — Pass 1: UPDATE existing rows
        -- =====================================================================
        MERGE INTO dbo.WipJCOP AS T
        USING (
            SELECT w.JCCo, w.Job, @Month AS Month,
                   ISNULL(w.GAAPCostOverride, 0) AS ProjCost,
                   ISNULL(w.OpsCostOverride, 0)  AS OtherAmount,
                   w.GAAPCostNotes AS Notes,
                   CASE WHEN w.GAAPCostPlugged = 1 THEN 'Y' ELSE 'N' END AS udPlugged
            FROM dbo.WipJobData w
            JOIN @DeptTable d ON LEFT(w.Job, CHARINDEX('.', w.Job) - 1) = d.Dept
            WHERE w.JCCo = @Co AND w.WipMonth = @Month
        ) AS S
        ON T.JCCo = S.JCCo AND T.Job = S.Job AND T.Month = S.Month
        WHEN MATCHED THEN
            UPDATE SET
                ProjCost    = S.ProjCost,
                OtherAmount = S.OtherAmount,
                Notes       = S.Notes,
                udPlugged   = S.udPlugged,
                WrittenBy   = @UserName,
                WrittenAt   = GETDATE();

        -- WipJCOP — Pass 2: INSERT new rows
        MERGE INTO dbo.WipJCOP AS T
        USING (
            SELECT w.JCCo, w.Job, @Month AS Month,
                   ISNULL(w.GAAPCostOverride, 0) AS ProjCost,
                   ISNULL(w.OpsCostOverride, 0)  AS OtherAmount,
                   w.GAAPCostNotes AS Notes,
                   CASE WHEN w.GAAPCostPlugged = 1 THEN 'Y' ELSE 'N' END AS udPlugged
            FROM dbo.WipJobData w
            JOIN @DeptTable d ON LEFT(w.Job, CHARINDEX('.', w.Job) - 1) = d.Dept
            WHERE w.JCCo = @Co AND w.WipMonth = @Month
        ) AS S
        ON T.JCCo = S.JCCo AND T.Job = S.Job AND T.Month = S.Month
        WHEN NOT MATCHED THEN
            INSERT (JCCo, Job, Month, ProjCost, OtherAmount, Notes, udPlugged, WrittenBy, WrittenAt)
            VALUES (S.JCCo, S.Job, S.Month, S.ProjCost, S.OtherAmount, S.Notes, S.udPlugged, @UserName, GETDATE());

        -- =====================================================================
        -- WipJCOR (Revenue overrides) — Pass 1: UPDATE existing rows
        -- Note: bJCOR is keyed by Contract, not Job. In LCG's data, Contract = Job.
        -- =====================================================================
        MERGE INTO dbo.WipJCOR AS T
        USING (
            SELECT w.JCCo, w.Job AS Contract, @Month AS Month,
                   ISNULL(w.GAAPRevOverride, 0) AS RevCost,
                   ISNULL(w.OpsRevOverride, 0)  AS OtherAmount,
                   w.GAAPRevNotes AS Notes,
                   CASE WHEN w.GAAPRevPlugged = 1 THEN 'Y' ELSE 'N' END AS udPlugged
            FROM dbo.WipJobData w
            JOIN @DeptTable d ON LEFT(w.Job, CHARINDEX('.', w.Job) - 1) = d.Dept
            WHERE w.JCCo = @Co AND w.WipMonth = @Month
        ) AS S
        ON T.JCCo = S.JCCo AND T.Contract = S.Contract AND T.Month = S.Month
        WHEN MATCHED THEN
            UPDATE SET
                RevCost     = S.RevCost,
                OtherAmount = S.OtherAmount,
                Notes       = S.Notes,
                udPlugged   = S.udPlugged,
                WrittenBy   = @UserName,
                WrittenAt   = GETDATE();

        -- WipJCOR — Pass 2: INSERT new rows
        MERGE INTO dbo.WipJCOR AS T
        USING (
            SELECT w.JCCo, w.Job AS Contract, @Month AS Month,
                   ISNULL(w.GAAPRevOverride, 0) AS RevCost,
                   ISNULL(w.OpsRevOverride, 0)  AS OtherAmount,
                   w.GAAPRevNotes AS Notes,
                   CASE WHEN w.GAAPRevPlugged = 1 THEN 'Y' ELSE 'N' END AS udPlugged
            FROM dbo.WipJobData w
            JOIN @DeptTable d ON LEFT(w.Job, CHARINDEX('.', w.Job) - 1) = d.Dept
            WHERE w.JCCo = @Co AND w.WipMonth = @Month
        ) AS S
        ON T.JCCo = S.JCCo AND T.Contract = S.Contract AND T.Month = S.Month
        WHEN NOT MATCHED THEN
            INSERT (JCCo, Contract, Month, RevCost, OtherAmount, Notes, udPlugged, WrittenBy, WrittenAt)
            VALUES (S.JCCo, S.Contract, S.Month, S.RevCost, S.OtherAmount, S.Notes, S.udPlugged, @UserName, GETDATE());

        COMMIT TRANSACTION;

        -- Report counts
        DECLARE @copCount INT, @corCount INT;
        SELECT @copCount = COUNT(*) FROM dbo.WipJCOP WHERE JCCo = @Co AND Month = @Month;
        SELECT @corCount = COUNT(*) FROM dbo.WipJCOR WHERE JCCo = @Co AND Month = @Month;

        SET @msg = 'Write-back complete. WipJCOP: ' + CAST(@copCount AS VARCHAR) +
                   ' rows, WipJCOR: ' + CAST(@corCount AS VARCHAR) + ' rows.';

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        SET @rcode = -1;
        SET @msg = 'Write-back failed: ' + ERROR_MESSAGE();
    END CATCH
END
GO

PRINT 'Created LylesWIPWriteBackToVista stored procedure.';
GO
