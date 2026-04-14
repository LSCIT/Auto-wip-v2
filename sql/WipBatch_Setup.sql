-- ============================================================
-- WIP Batch Tracking Setup
-- Run against: 10.112.11.8 / Viewpoint database
-- Created: 2026-03-23
-- Purpose: Replaces WipDb batch management with a clean,
--          owned table directly in Viewpoint.
--
-- COLLATION NOTE: This database uses Latin1_General_BIN (binary).
--   - All object names are case-sensitive and must be referenced
--     with exact casing (e.g., "udWIPBatch" not "udwipbatch")
--   - String data comparisons are byte-for-byte exact
--   - BatchState values ('Open', 'ReadyForOps', etc.) must always
--     be passed with this exact casing from VBA
--
-- SAFE TO RE-RUN: All statements are guarded with IF NOT EXISTS
--   or CREATE OR ALTER. Will not overwrite existing data.
-- ============================================================


-- ============================================================
-- TABLE: udWIPBatch
-- One row per company / WIP month / department cycle.
-- State machine: Open -> ReadyForOps -> OpsApproved -> AcctApproved
-- Named with 'ud' prefix per Viewpoint custom table convention.
-- ============================================================
IF OBJECT_ID('dbo.udWIPBatch', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.udWIPBatch (
        BatchId         INT IDENTITY(1,1) PRIMARY KEY,
        JCCo            TINYINT         NOT NULL,
        WipMonth        DATE            NOT NULL,
        Department      CHAR(2)         NOT NULL,
        BatchState      VARCHAR(20)     NOT NULL CONSTRAINT DF_udWIPBatch_State DEFAULT 'Open',
        CreatedBy       VARCHAR(100)    NOT NULL,
        CreatedAt       DATETIME        NOT NULL CONSTRAINT DF_udWIPBatch_Created DEFAULT GETDATE(),
        StateChangedBy  VARCHAR(100)    NULL,
        StateChangedAt  DATETIME        NULL,
        CONSTRAINT UQ_udWIPBatch UNIQUE (JCCo, WipMonth, Department)
    );
    PRINT 'Created table dbo.udWIPBatch';
END
ELSE
BEGIN
    PRINT 'Table dbo.udWIPBatch already exists — skipped. Run SELECT * FROM dbo.udWIPBatch to inspect.';
END
GO


-- ============================================================
-- PROC: LylesWIPBatchGet
-- Returns all batches for a company + month.
-- Called by UseExistingBatch to check if the month is already open.
-- ============================================================
IF OBJECT_ID('dbo.LylesWIPBatchGet', 'P') IS NULL
    EXEC('CREATE PROCEDURE dbo.LylesWIPBatchGet AS BEGIN END');
GO

ALTER PROCEDURE dbo.LylesWIPBatchGet
    @Co     TINYINT,
    @Month  DATE
AS
BEGIN
    SET NOCOUNT ON;
    SELECT BatchId, Department, BatchState, CreatedBy, CreatedAt
    FROM   dbo.udWIPBatch
    WHERE  JCCo = @Co AND WipMonth = @Month
    ORDER BY Department;
END
GO
PRINT 'Created/updated procedure dbo.LylesWIPBatchGet';


-- ============================================================
-- PROC: LylesWIPBatchCreate
-- Creates a new batch for a single dept, or returns the existing one.
-- @IsNew = 1 if a new row was inserted, 0 if it already existed.
-- NOTE: Dept is CHAR(2) — pass exactly 2 characters.
--       Trailing space padding from CHAR type is handled automatically.
-- ============================================================
IF OBJECT_ID('dbo.LylesWIPBatchCreate', 'P') IS NULL
    EXEC('CREATE PROCEDURE dbo.LylesWIPBatchCreate AS BEGIN END');
GO

ALTER PROCEDURE dbo.LylesWIPBatchCreate
    @Co         TINYINT,
    @Month      DATE,
    @Dept       CHAR(2),
    @UserName   VARCHAR(100),
    @BatchId    INT OUTPUT,
    @IsNew      BIT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    IF EXISTS (
        SELECT 1 FROM dbo.udWIPBatch
        WHERE JCCo = @Co AND WipMonth = @Month AND Department = @Dept
    )
    BEGIN
        SELECT @BatchId = BatchId
        FROM   dbo.udWIPBatch
        WHERE  JCCo = @Co AND WipMonth = @Month AND Department = @Dept;
        SET @IsNew = 0;
    END
    ELSE
    BEGIN
        INSERT INTO dbo.udWIPBatch (JCCo, WipMonth, Department, BatchState, CreatedBy)
        VALUES (@Co, @Month, @Dept, 'Open', @UserName);
        SET @BatchId = SCOPE_IDENTITY();
        SET @IsNew = 1;
    END
END
GO
PRINT 'Created/updated procedure dbo.LylesWIPBatchCreate';


-- ============================================================
-- PROC: LylesWIPBatchSetState
-- Sets the batch state for a given co/month/dept.
-- Valid states (case-sensitive — Latin1_General_BIN):
--   'Open'          set by system when batch is created
--   'ReadyForOps'   set by RFOYes_Click (Accounting → Ops handoff)
--   'OpsApproved'   set by OFAYes_Click (Ops final approval)
--   'AcctApproved'  set by AFAYes_Click (Accounting final approval)
-- ============================================================
IF OBJECT_ID('dbo.LylesWIPBatchSetState', 'P') IS NULL
    EXEC('CREATE PROCEDURE dbo.LylesWIPBatchSetState AS BEGIN END');
GO

ALTER PROCEDURE dbo.LylesWIPBatchSetState
    @Co         TINYINT,
    @Month      DATE,
    @Dept       CHAR(2),
    @NewState   VARCHAR(20),
    @UserName   VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.udWIPBatch
    SET BatchState     = @NewState,
        StateChangedBy = @UserName,
        StateChangedAt = GETDATE()
    WHERE JCCo = @Co AND WipMonth = @Month AND Department = @Dept;

    IF @@ROWCOUNT = 0
        RAISERROR('No batch found for Co=%d, Month=%s, Dept=%s', 16, 1,
                  @Co, @Month, @Dept);
END
GO
PRINT 'Created/updated procedure dbo.LylesWIPBatchSetState';


-- ============================================================
-- VERIFY: Quick sanity check after install
-- ============================================================
SELECT
    OBJECT_ID('dbo.udWIPBatch',           'U') AS udWIPBatch_TableId,
    OBJECT_ID('dbo.LylesWIPBatchGet',     'P') AS LylesWIPBatchGet_ProcId,
    OBJECT_ID('dbo.LylesWIPBatchCreate',  'P') AS LylesWIPBatchCreate_ProcId,
    OBJECT_ID('dbo.LylesWIPBatchSetState','P') AS LylesWIPBatchSetState_ProcId;
-- All four values should be non-NULL after a successful run.
GO
