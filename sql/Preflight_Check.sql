-- ============================================================
-- Viewpoint Pre-Flight Check
-- Run against: 10.112.11.8 / Viewpoint database
-- Purpose: Before running WipBatch_Setup.sql, confirm what
--          already exists so we don't clobber anything.
-- ============================================================

-- ============================================================
-- 1. Database collation (must know this — is it CS or CI?)
-- ============================================================
SELECT
    name                    AS DatabaseName,
    collation_name          AS Collation,
    CASE WHEN collation_name LIKE '%_CS_%' THEN 'CASE SENSITIVE'
         WHEN collation_name LIKE '%_CI_%' THEN 'case insensitive'
         ELSE 'unknown'
    END                     AS CaseSensitivity
FROM sys.databases
WHERE name = DB_NAME();
GO

-- ============================================================
-- 2. Server-level collation
-- ============================================================
SELECT
    SERVERPROPERTY('ServerName')   AS ServerName,
    SERVERPROPERTY('Collation')    AS ServerCollation;
GO

-- ============================================================
-- 3. ALL user-defined tables (ud* prefix) in this database
--    These are the custom tables Michael or others have added.
-- ============================================================
SELECT
    t.name              AS TableName,
    s.name              AS SchemaName,
    t.create_date       AS CreatedDate,
    t.modify_date       AS ModifiedDate,
    (SELECT COUNT(*) FROM sys.columns c WHERE c.object_id = t.object_id) AS ColumnCount,
    p.rows              AS RowCount
FROM sys.tables t
JOIN sys.schemas s      ON s.schema_id = t.schema_id
JOIN sys.partitions p   ON p.object_id = t.object_id AND p.index_id IN (0,1)
WHERE t.name LIKE 'ud%'
ORDER BY t.name;
GO

-- ============================================================
-- 4. SPECIFICALLY check for our target objects
--    (using exact case since collation is CS)
-- ============================================================

-- Does udWIPBatch exist?
SELECT
    CASE WHEN OBJECT_ID('dbo.udWIPBatch', 'U') IS NOT NULL
         THEN '*** EXISTS — DO NOT re-run CREATE TABLE ***'
         ELSE 'OK — does not exist yet'
    END AS udWIPBatch_Status;

-- Does udWIPJV exist? (Michael's existing table — should be there)
SELECT
    CASE WHEN OBJECT_ID('dbo.udWIPJV', 'U') IS NOT NULL
         THEN 'EXISTS (expected — Michael created this)'
         ELSE 'NOT FOUND (unexpected)'
    END AS udWIPJV_Status;
GO

-- ============================================================
-- 5. ALL stored procedures with LCG, Lyles, or ud prefix
--    Shows everything Michael created + what we'd be adding
-- ============================================================
SELECT
    p.name              AS ProcName,
    s.name              AS SchemaName,
    p.create_date       AS CreatedDate,
    p.modify_date       AS ModifiedDate
FROM sys.procedures p
JOIN sys.schemas s ON s.schema_id = p.schema_id
WHERE p.name LIKE 'LCG%'
   OR p.name LIKE 'Lyles%'
   OR p.name LIKE 'ud%'
   OR p.name LIKE 'lcg%'      -- catch case variants
   OR p.name LIKE 'lyles%'
ORDER BY p.name;
GO

-- ============================================================
-- 6. SPECIFICALLY check for our three target procs
-- ============================================================
SELECT
    CASE WHEN OBJECT_ID('dbo.LylesWIPBatchGet',      'P') IS NOT NULL THEN '*** EXISTS ***' ELSE 'OK — does not exist' END AS LylesWIPBatchGet_Status,
    CASE WHEN OBJECT_ID('dbo.LylesWIPBatchCreate',   'P') IS NOT NULL THEN '*** EXISTS ***' ELSE 'OK — does not exist' END AS LylesWIPBatchCreate_Status,
    CASE WHEN OBJECT_ID('dbo.LylesWIPBatchSetState', 'P') IS NOT NULL THEN '*** EXISTS ***' ELSE 'OK — does not exist' END AS LylesWIPBatchSetState_Status;
GO

-- ============================================================
-- 7. Check udWIPJV structure (Michael's table — confirms our
--    naming convention is right and shows column patterns)
-- ============================================================
SELECT
    c.name          AS ColumnName,
    t.name          AS DataType,
    c.max_length    AS MaxLength,
    c.is_nullable   AS IsNullable
FROM sys.columns c
JOIN sys.types t    ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.udWIPJV', 'U')
ORDER BY c.column_id;
GO

-- ============================================================
-- 8. Check if Department in bJCDM is CHAR or VARCHAR, and length
--    (so our CHAR(2) in udWIPBatch matches the source type)
-- ============================================================
SELECT
    c.name          AS ColumnName,
    t.name          AS DataType,
    c.max_length    AS MaxLength,
    c.is_nullable   AS IsNullable
FROM sys.columns c
JOIN sys.types t    ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.bJCDM', 'U')
  AND c.name = 'Department';
GO

-- ============================================================
-- 9. Show existing LCGWIPBatchCheck1 proc definition
--    (if it exists — shows how Michael stored batch state,
--     confirms we're not stepping on his approach)
-- ============================================================
IF OBJECT_ID('dbo.LCGWIPBatchCheck1', 'P') IS NOT NULL
BEGIN
    EXEC sp_helptext 'LCGWIPBatchCheck1';
END
ELSE
BEGIN
    SELECT 'LCGWIPBatchCheck1 does not exist in this database' AS Note;
END
GO

-- ============================================================
-- SUMMARY: What to do with results
-- ============================================================
-- If udWIPBatch shows "EXISTS":
--   → Someone already ran our script (or Michael had a similar table).
--   → Run "SELECT * FROM dbo.udWIPBatch" to see what's in it.
--   → Do NOT re-run WipBatch_Setup.sql.
--
-- If LylesWIPBatch* procs show "EXISTS":
--   → Use ALTER PROCEDURE instead of CREATE PROCEDURE in setup script.
--   → Or DROP first (only safe if table is also being recreated).
--
-- If Department in bJCDM is NOT CHAR(2):
--   → Update the @Dept parameter type in WipBatch_Setup.sql to match.
--
-- If LCGWIPBatchCheck1 exists and has batch state logic:
--   → Review its WipBatch table to see if we should migrate data.
