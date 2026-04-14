-- =============================================================================
-- LylesWIP Database Creation Script
-- Server:  10.103.30.11 (Cloud-Apps1 / P&P server)
-- Run as:  sysadmin (sa or equivalent)
-- Purpose: Creates the LylesWIP database and all schema objects for Auto-WIP
-- Sprint:  3
-- =============================================================================
-- INSTRUCTIONS:
--   1. Run this entire script as sysadmin (sa or Windows admin account)
--   2. wip.excel.sql login will be granted db_owner on LylesWIP
--   3. After this runs, execute load_overrides.py to seed historical data
-- =============================================================================

USE master;
GO

-- =============================================================================
-- Create database
-- =============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.databases WHERE name = 'LylesWIP')
BEGIN
    CREATE DATABASE LylesWIP;
    PRINT 'LylesWIP database created.';
END
ELSE
    PRINT 'LylesWIP database already exists — skipping CREATE.';
GO

USE LylesWIP;
GO

-- =============================================================================
-- Grant db_owner to application login
-- =============================================================================
IF NOT EXISTS (
    SELECT 1 FROM sys.database_principals WHERE name = 'wip.excel.sql'
)
BEGIN
    CREATE USER [wip.excel.sql] FOR LOGIN [wip.excel.sql];
    PRINT 'Database user wip.excel.sql created.';
END

IF IS_ROLEMEMBER('db_owner', 'wip.excel.sql') = 0
BEGIN
    ALTER ROLE db_owner ADD MEMBER [wip.excel.sql];
    PRINT 'wip.excel.sql granted db_owner.';
END
GO

-- =============================================================================
-- TABLE: WipBatches
-- One row per company / month / department. Tracks the batch lifecycle.
-- =============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'WipBatches')
BEGIN
    CREATE TABLE dbo.WipBatches (
        Id              INT IDENTITY(1,1) PRIMARY KEY,
        JCCo            TINYINT       NOT NULL,
        WipMonth        DATE          NOT NULL,   -- First of month (e.g. 2025-12-01)
        Department      VARCHAR(10)   NOT NULL,
        -- State machine: Open → ReadyForOps → OpsApproved → AcctApproved
        BatchState      VARCHAR(20)   NOT NULL DEFAULT 'Open',
        CreatedBy       VARCHAR(100)  NOT NULL,
        CreatedAt       DATETIME      NOT NULL DEFAULT GETDATE(),
        StateChangedBy  VARCHAR(100)  NULL,
        StateChangedAt  DATETIME      NULL,
        CONSTRAINT UQ_WipBatches UNIQUE (JCCo, WipMonth, Department)
    );
    PRINT 'Table WipBatches created.';
END
GO

-- =============================================================================
-- TABLE: WipJobData
-- One row per job per WIP month. Stores all override values.
-- NULL override = use Vista-calculated value (override takes priority if present).
-- =============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'WipJobData')
BEGIN
    CREATE TABLE dbo.WipJobData (
        Id                  INT IDENTITY(1,1) PRIMARY KEY,
        JCCo                TINYINT        NOT NULL,
        Job                 VARCHAR(50)    NOT NULL,   -- Raw Vista job number (e.g. '51.1108.')
        WipMonth            DATE           NOT NULL,   -- First of month
        -- Revenue overrides (NULL = use Vista-calculated value)
        OpsRevOverride      DECIMAL(15,2)  NULL,
        OpsRevPlugged       BIT            NOT NULL DEFAULT 0,
        GAAPRevOverride     DECIMAL(15,2)  NULL,
        GAAPRevPlugged      BIT            NOT NULL DEFAULT 0,
        -- Cost overrides (NULL = use Vista-calculated value)
        OpsCostOverride     DECIMAL(15,2)  NULL,
        OpsCostPlugged      BIT            NOT NULL DEFAULT 0,
        GAAPCostOverride    DECIMAL(15,2)  NULL,
        GAAPCostPlugged     BIT            NOT NULL DEFAULT 0,
        -- Bonus profit
        BonusProfit         DECIMAL(15,2)  NULL,
        -- Notes
        OpsRevNotes         VARCHAR(500)   NULL,
        GAAPRevNotes        VARCHAR(500)   NULL,
        OpsCostNotes        VARCHAR(500)   NULL,
        GAAPCostNotes       VARCHAR(500)   NULL,
        -- Completion date (user-entered)
        CompletionDate      DATE           NULL,
        -- Per-job workflow flags
        IsClosed            BIT            NOT NULL DEFAULT 0,
        IsOpsDone           BIT            NOT NULL DEFAULT 0,
        IsGAAPDone          BIT            NOT NULL DEFAULT 0,
        -- Audit
        UserName            VARCHAR(100)   NOT NULL DEFAULT '',
        UpdatedAt           DATETIME       NOT NULL DEFAULT GETDATE(),
        Source              VARCHAR(20)    NOT NULL DEFAULT 'UserEdit',  -- 'ExcelImport' or 'UserEdit'
        CONSTRAINT UQ_WipJobData UNIQUE (JCCo, Job, WipMonth)
    );
    PRINT 'Table WipJobData created.';
END
GO

-- =============================================================================
-- TABLE: WipYearEndSnapshot
-- Populated at December AcctApproved. Source for Prior Year columns next year.
-- =============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'WipYearEndSnapshot')
BEGIN
    CREATE TABLE dbo.WipYearEndSnapshot (
        Id                  INT IDENTITY(1,1) PRIMARY KEY,
        JCCo                TINYINT        NOT NULL,
        Job                 VARCHAR(50)    NOT NULL,
        SnapshotYear        SMALLINT       NOT NULL,  -- e.g. 2025
        PriorYearGAAPRev    DECIMAL(15,2)  NULL,
        PriorYearGAAPCost   DECIMAL(15,2)  NULL,
        PriorYearOpsRev     DECIMAL(15,2)  NULL,
        PriorYearOpsCost    DECIMAL(15,2)  NULL,
        BonusProfit         DECIMAL(15,2)  NULL,
        CreatedAt           DATETIME       NOT NULL DEFAULT GETDATE(),
        CONSTRAINT UQ_WipYearEndSnapshot UNIQUE (JCCo, Job, SnapshotYear)
    );
    PRINT 'Table WipYearEndSnapshot created.';
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPCreateBatch
-- Creates a new batch for Co/Month/Dept, or returns existing one.
-- Called from VBA Module6.CreateBatch after Ops selects division.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPCreateBatch', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPCreateBatch;
GO
CREATE PROCEDURE dbo.LylesWIPCreateBatch
    @JCCo       TINYINT,
    @WipMonth   DATE,
    @Department VARCHAR(10),
    @CreatedBy  VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (
        SELECT 1 FROM dbo.WipBatches
        WHERE JCCo = @JCCo AND WipMonth = @WipMonth AND Department = @Department
    )
    BEGIN
        INSERT INTO dbo.WipBatches (JCCo, WipMonth, Department, BatchState, CreatedBy)
        VALUES (@JCCo, @WipMonth, @Department, 'Open', @CreatedBy);
    END

    SELECT Id, JCCo, WipMonth, Department, BatchState, CreatedBy, CreatedAt
    FROM dbo.WipBatches
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth AND Department = @Department;
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPGetBatches
-- Returns all batches for a Co/Month (all departments).
-- Called on workbook open to check if batches exist before prompting division select.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPGetBatches', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPGetBatches;
GO
CREATE PROCEDURE dbo.LylesWIPGetBatches
    @JCCo     TINYINT,
    @WipMonth DATE
AS
BEGIN
    SET NOCOUNT ON;
    SELECT Id, JCCo, WipMonth, Department, BatchState, CreatedBy, CreatedAt,
           StateChangedBy, StateChangedAt
    FROM dbo.WipBatches
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth
    ORDER BY Department;
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPCheckBatchState
-- Returns the current state for a specific Co/Month/Dept batch.
-- Called from Module6.UseExistingBatch and from state-gate checks.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPCheckBatchState', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPCheckBatchState;
GO
CREATE PROCEDURE dbo.LylesWIPCheckBatchState
    @JCCo       TINYINT,
    @WipMonth   DATE,
    @Department VARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT Id, BatchState, CreatedBy, CreatedAt, StateChangedBy, StateChangedAt
    FROM dbo.WipBatches
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth AND Department = @Department;
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPUpdateBatchState
-- Advances the batch state. Validates allowed transitions.
-- Transitions: Open → ReadyForOps → OpsApproved → AcctApproved
-- Called from: RFOYes_Click, OFAYes_Click, AFAYes_Click in FormButtons.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPUpdateBatchState', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPUpdateBatchState;
GO
CREATE PROCEDURE dbo.LylesWIPUpdateBatchState
    @JCCo           TINYINT,
    @WipMonth       DATE,
    @Department     VARCHAR(10),
    @NewState       VARCHAR(20),
    @ChangedBy      VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @CurrentState VARCHAR(20);
    DECLARE @MonthStr    VARCHAR(20) = CONVERT(VARCHAR(20), @WipMonth, 23);

    SELECT @CurrentState = BatchState
    FROM dbo.WipBatches
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth AND Department = @Department;

    IF @CurrentState IS NULL
    BEGIN
        RAISERROR('Batch not found for Co=%d, Month=%s, Dept=%s', 16, 1,
                  @JCCo, @MonthStr, @Department);
        RETURN;
    END

    -- Validate allowed state transitions
    IF NOT (
        (@CurrentState = 'Open'         AND @NewState = 'ReadyForOps')  OR
        (@CurrentState = 'ReadyForOps'  AND @NewState = 'OpsApproved')  OR
        (@CurrentState = 'OpsApproved'  AND @NewState = 'AcctApproved') OR
        (@NewState = 'Open')  -- Allow reset to Open from any state (admin use)
    )
    BEGIN
        RAISERROR('Invalid state transition from %s to %s', 16, 1,
                  @CurrentState, @NewState);
        RETURN;
    END

    UPDATE dbo.WipBatches
    SET BatchState     = @NewState,
        StateChangedBy = @ChangedBy,
        StateChangedAt = GETDATE()
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth AND Department = @Department;

    SELECT Id, BatchState, StateChangedBy, StateChangedAt
    FROM dbo.WipBatches
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth AND Department = @Department;
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPSaveJobRow
-- MERGE one job row into WipJobData. Called on double-click Done (col H or I).
-- NULL values for overrides mean "no override — use Vista value."
-- Plugged = 1 means the value was manually entered/overridden (yellow highlight).
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPSaveJobRow', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPSaveJobRow;
GO
CREATE PROCEDURE dbo.LylesWIPSaveJobRow
    @JCCo           TINYINT,
    @Job            VARCHAR(50),
    @WipMonth       DATE,
    @OpsRevOverride     DECIMAL(15,2) = NULL,
    @OpsRevPlugged      BIT           = 0,
    @GAAPRevOverride    DECIMAL(15,2) = NULL,
    @GAAPRevPlugged     BIT           = 0,
    @OpsCostOverride    DECIMAL(15,2) = NULL,
    @OpsCostPlugged     BIT           = 0,
    @GAAPCostOverride   DECIMAL(15,2) = NULL,
    @GAAPCostPlugged    BIT           = 0,
    @BonusProfit        DECIMAL(15,2) = NULL,
    @OpsRevNotes        VARCHAR(500)  = NULL,
    @GAAPRevNotes       VARCHAR(500)  = NULL,
    @OpsCostNotes       VARCHAR(500)  = NULL,
    @GAAPCostNotes      VARCHAR(500)  = NULL,
    @CompletionDate     DATE          = NULL,
    @IsClosed           BIT           = 0,
    @IsOpsDone          BIT           = 0,
    @IsGAAPDone         BIT           = 0,
    @UserName           VARCHAR(100)  = ''
AS
BEGIN
    SET NOCOUNT ON;

    MERGE dbo.WipJobData AS target
    USING (SELECT @JCCo AS JCCo, @Job AS Job, @WipMonth AS WipMonth) AS source
        ON target.JCCo = source.JCCo
       AND target.Job = source.Job
       AND target.WipMonth = source.WipMonth
    WHEN MATCHED THEN
        UPDATE SET
            OpsRevOverride   = @OpsRevOverride,
            OpsRevPlugged    = @OpsRevPlugged,
            GAAPRevOverride  = @GAAPRevOverride,
            GAAPRevPlugged   = @GAAPRevPlugged,
            OpsCostOverride  = @OpsCostOverride,
            OpsCostPlugged   = @OpsCostPlugged,
            GAAPCostOverride = @GAAPCostOverride,
            GAAPCostPlugged  = @GAAPCostPlugged,
            BonusProfit      = ISNULL(@BonusProfit, BonusProfit),
            OpsRevNotes      = @OpsRevNotes,
            GAAPRevNotes     = @GAAPRevNotes,
            OpsCostNotes     = @OpsCostNotes,
            GAAPCostNotes    = @GAAPCostNotes,
            CompletionDate   = @CompletionDate,
            IsClosed         = @IsClosed,
            IsOpsDone        = @IsOpsDone,
            IsGAAPDone       = @IsGAAPDone,
            UserName         = @UserName,
            UpdatedAt        = GETDATE(),
            Source           = 'UserEdit'
    WHEN NOT MATCHED THEN
        INSERT (JCCo, Job, WipMonth,
                OpsRevOverride, OpsRevPlugged, GAAPRevOverride, GAAPRevPlugged,
                OpsCostOverride, OpsCostPlugged, GAAPCostOverride, GAAPCostPlugged,
                BonusProfit, OpsRevNotes, GAAPRevNotes, OpsCostNotes, GAAPCostNotes,
                CompletionDate, IsClosed, IsOpsDone, IsGAAPDone, UserName, Source)
        VALUES (@JCCo, @Job, @WipMonth,
                @OpsRevOverride, @OpsRevPlugged, @GAAPRevOverride, @GAAPRevPlugged,
                @OpsCostOverride, @OpsCostPlugged, @GAAPCostOverride, @GAAPCostPlugged,
                @BonusProfit, @OpsRevNotes, @GAAPRevNotes, @OpsCostNotes, @GAAPCostNotes,
                @CompletionDate, @IsClosed, @IsOpsDone, @IsGAAPDone, @UserName, 'UserEdit');
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPGetJobOverrides
-- Returns all WipJobData rows for a company/month.
-- VBA merges these over the live Vista values after data load.
-- No department filter — VBA handles that via the job list from the Vista query.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPGetJobOverrides', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPGetJobOverrides;
GO
CREATE PROCEDURE dbo.LylesWIPGetJobOverrides
    @JCCo     TINYINT,
    @WipMonth DATE
AS
BEGIN
    SET NOCOUNT ON;
    SELECT
        JCCo, Job, WipMonth,
        OpsRevOverride, OpsRevPlugged, GAAPRevOverride, GAAPRevPlugged,
        OpsCostOverride, OpsCostPlugged, GAAPCostOverride, GAAPCostPlugged,
        BonusProfit,
        OpsRevNotes, GAAPRevNotes, OpsCostNotes, GAAPCostNotes,
        CompletionDate, IsClosed, IsOpsDone, IsGAAPDone,
        UserName, UpdatedAt
    FROM dbo.WipJobData
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth;
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPClearJobData
-- Deletes all WipJobData for a Co/Month (batch cancel or clear).
-- Does NOT delete WipBatches — use LylesWIPUpdateBatchState to reset state.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPClearJobData', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPClearJobData;
GO
CREATE PROCEDURE dbo.LylesWIPClearJobData
    @JCCo     TINYINT,
    @WipMonth DATE
AS
BEGIN
    SET NOCOUNT ON;
    DELETE FROM dbo.WipJobData
    WHERE JCCo = @JCCo AND WipMonth = @WipMonth;

    SELECT @@ROWCOUNT AS RowsDeleted;
END
GO

-- =============================================================================
-- STORED PROC: LylesWIPSaveYearEndSnapshot
-- Called at December AcctApproved. Captures final approved values as the
-- prior-year baseline for next year's WIP runs.
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPSaveYearEndSnapshot', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPSaveYearEndSnapshot;
GO
CREATE PROCEDURE dbo.LylesWIPSaveYearEndSnapshot
    @JCCo               TINYINT,
    @Job                VARCHAR(50),
    @SnapshotYear       SMALLINT,
    @PriorYearGAAPRev   DECIMAL(15,2) = NULL,
    @PriorYearGAAPCost  DECIMAL(15,2) = NULL,
    @PriorYearOpsRev    DECIMAL(15,2) = NULL,
    @PriorYearOpsCost   DECIMAL(15,2) = NULL,
    @BonusProfit        DECIMAL(15,2) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    MERGE dbo.WipYearEndSnapshot AS target
    USING (SELECT @JCCo AS JCCo, @Job AS Job, @SnapshotYear AS SnapshotYear) AS source
        ON target.JCCo = source.JCCo
       AND target.Job = source.Job
       AND target.SnapshotYear = source.SnapshotYear
    WHEN MATCHED THEN
        UPDATE SET
            PriorYearGAAPRev  = @PriorYearGAAPRev,
            PriorYearGAAPCost = @PriorYearGAAPCost,
            PriorYearOpsRev   = @PriorYearOpsRev,
            PriorYearOpsCost  = @PriorYearOpsCost,
            BonusProfit       = @BonusProfit,
            CreatedAt         = GETDATE()
    WHEN NOT MATCHED THEN
        INSERT (JCCo, Job, SnapshotYear,
                PriorYearGAAPRev, PriorYearGAAPCost,
                PriorYearOpsRev, PriorYearOpsCost, BonusProfit)
        VALUES (@JCCo, @Job, @SnapshotYear,
                @PriorYearGAAPRev, @PriorYearGAAPCost,
                @PriorYearOpsRev, @PriorYearOpsCost, @BonusProfit);
END
GO

-- =============================================================================
-- Verify
-- =============================================================================
SELECT 'Tables:' AS [Object], name AS ObjectName FROM sys.tables WHERE schema_id = SCHEMA_ID('dbo')
UNION ALL
SELECT 'Procs:', name FROM sys.procedures WHERE schema_id = SCHEMA_ID('dbo')
ORDER BY [Object], ObjectName;
GO

PRINT 'LylesWIP schema creation complete.';
GO
