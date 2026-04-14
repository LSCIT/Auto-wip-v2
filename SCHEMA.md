# Database Schema Reference
*Last updated: 2026-04-02*

---

## Vista (Viewpoint) — Read Only
**Server:** `10.112.11.8` (prod) / `VM111VPPRD1` (by hostname)
**Database:** `Viewpoint`
**Auth:** `WIPexcel` / see Settings C8/C9 (VPPassword updated March 5 — get from workbook Settings sheet)
**Access pattern:** MSOLEDBSQL ADODB connection from VBA. Read-only — no writes to Vista in Phase 1.

### Tables Used by WIP Query

| Table | Purpose | Key Columns |
|-------|---------|-------------|
| `bJCJM` | Job Master | `JCCo`, `Job` (VARCHAR, trailing dot for standard jobs, no dot for sub-jobs), `Description`, `Contract`, `JobStatus` (1=Open,2=SoftClosed,3=HardClosed), `ProjectMgr`, `ProjMinPct` |
| `bJCCM` | Contract Master | `JCCo`, `Contract`, `Department` (matches division in WIP context), `OrigContractAmt`, `ContractAmt`, `BilledAmt`, `ReceivedAmt`, `CurrentRetainAmt`, `MonthClosed`, `ContractStatus`, `StartMonth` |
| `bJCDM` | Department Master | `JCCo`, `Department`, `Description` |
| `bJCCD` | Job Cost by Cost Detail | `JCCo`, `Job` (raw — no trim in JOINs!), `Mth`, `CostType`, `ActualCost`, `ProjCost`, `JCTransType`, `PostedDate`, `ActualDate` |
| `bJCMP` | Project Manager | `JCCo`, `ProjectMgr`, `Name` |
| `bGLCO` | GL Company | `GLCo`, `LastMthSubClsd` — used for GL closed-month check |
| `bARTH` | AR Transaction History | `JCCo`, `Contract`, `ARTransType` ('I'=Invoice), `Mth` — used for billing-activity job inclusion |
| `bHQCO` | HQ Company | `HQCo`, `Name` — company name/list for dropdowns |

### Job Number Format in bJCJM
- Standard jobs: `{div}.{job}.` — e.g., `56.1004.` (trailing dot always present)
- Sub-jobs: `{div}.{job}.{phase}` — e.g., `56.1010.01` (NO trailing dot; only 2 exist: 56.1009.01, 56.1010.01)
- All jobs confirmed via SQL: every bJCJM row has exactly 2 dots (counting the standard trailing dot)
- `Job` column is a fixed-length field; pyodbc returns it with trailing spaces — always RTRIM in Python/SQL

### WIP Query Parameters
```sql
@Co         TINYINT   -- Company (15=WML, 16=AIC, 12=APC, 13=NESM)
@Month      DATE      -- First of month: '2025-12-01'
@DeptList   VARCHAR   -- Comma-separated dept codes: '54' or '51,52,53'
@GroupBy    VARCHAR   -- 'Department' (always for WIP Schedule)

-- Derived:
@CutOffDate = EOMONTH(@Month)       -- Last day of month (e.g. 2025-12-31)
@StartDate  = Jan 1 of WIP year     -- For current-year activity filters
@PriorYrEnd = Dec 31 prior year     -- For cross-year reversal detection
@BillingMth = First of WIP month    -- For bARTH billing filter
```

### Job Inclusion Logic (JobList CTE)
A job appears in the WIP if ANY of:
1. `JobStatus = 1` (Open) AND `c.StartMonth <= @CutOffDate`
2. `JobStatus = 2` (Soft Closed) AND `MonthClosed >= @StartDate` (closed in current year)
3. `JobStatus = 3` (Hard Closed) AND `MonthClosed >= @Month` (closed this month or later)
4. Has actual cost activity in `bJCCD` in current year (catches closed jobs with year-to-date reversals)
5. Has billing activity in `bARTH` in current year (catches closed jobs with retainage/warranty invoices)

**A7 rule:** Zero-JTD cross-year reversal jobs are included naturally by the EXISTS on bJCCD.

### Key Business Rule: No LTRIM/RTRIM in JOINs
```sql
-- CORRECT (58 seconds on 8.8M rows):
JOIN bJCCD cd ON cd.JCCo = j.JCCo AND cd.Job = j.Job

-- WRONG (9+ minutes — causes full table scan):
JOIN bJCCD cd ON cd.JCCo = j.JCCo AND LTRIM(RTRIM(cd.Job)) = LTRIM(RTRIM(j.Job))
```
Trim only in the final SELECT output column, never in JOIN predicates or GROUP BY.

### Tables NOT Used (legacy Michael Roberts tables)
- `budWIPDetail` — empty in production; Michael never finished the batch population proc
- `budMoJobSumAppr` — superseded; LylesWIP replaces this
- `LCGWIPSchedule` — historical final approved ops data through Nov 2025 (2,685 rows); read-only reference
- `udWIPJV` — does NOT exist on production Vista (10.112.11.8); only on JERC test server

### Future Write-Back Tables (Phase 2 — post-delivery)
- `bJCOR` — Viewpoint job cost override table; target for writing final approved OPS values back to Vista
- `bJCOP` — Viewpoint job cost override by period; target for GAAP values
- Spec TBD: confirm column mapping with Nicole before Phase 2 write-back build

---

## LylesWIP — Read + Write
**Server:** `10.103.30.11` (Cloud-Apps1 / P&P server)
**Database:** `LylesWIP`
**Auth:** `wip.excel.sql` / `WES@2024` (db_owner on LylesWIP)
**Schema:** Created by `sql/LylesWIP_CreateDB.sql`

### Tables

#### dbo.WipBatches
One row per company / month / department. Tracks batch lifecycle.
```
Id              INT IDENTITY PK
JCCo            TINYINT          -- Company code (15, 16, 12, 13)
WipMonth        DATE             -- First of month ('2025-12-01')
Department      VARCHAR(10)      -- Division code ('54', '51', etc.)
BatchState      VARCHAR(20)      -- 'Open' | 'ReadyForOps' | 'OpsApproved' | 'AcctApproved'
CreatedBy       VARCHAR(100)
CreatedAt       DATETIME
StateChangedBy  VARCHAR(100)     NULL
StateChangedAt  DATETIME         NULL
UNIQUE (JCCo, WipMonth, Department)
```

#### dbo.WipJobData
One row per job per WIP month. All override values for a job.
NULL override = "no override, use Vista-calculated value".
Non-NULL with Plugged=1 = user manually entered this value.
```
Id                  INT IDENTITY PK
JCCo                TINYINT
Job                 VARCHAR(50)      -- Raw Vista job number with trailing dot ('51.1108.')
WipMonth            DATE
OpsRevOverride      DECIMAL(15,2)    NULL  -- NULL = no override
OpsRevPlugged       BIT              0/1
GAAPRevOverride     DECIMAL(15,2)    NULL
GAAPRevPlugged      BIT
OpsCostOverride     DECIMAL(15,2)    NULL
OpsCostPlugged      BIT
GAAPCostOverride    DECIMAL(15,2)    NULL
GAAPCostPlugged     BIT
BonusProfit         DECIMAL(15,2)    NULL
OpsRevNotes         VARCHAR(500)     NULL
GAAPRevNotes        VARCHAR(500)     NULL
OpsCostNotes        VARCHAR(500)     NULL
GAAPCostNotes       VARCHAR(500)     NULL
CompletionDate      DATE             NULL
IsClosed            BIT              0/1
IsOpsDone           BIT              0/1   -- Set on Col H double-click
IsGAAPDone          BIT              0/1   -- Set on Col I double-click
UserName            VARCHAR(100)
UpdatedAt           DATETIME
Source              VARCHAR(20)      -- 'ExcelImport' or 'UserEdit'
UNIQUE (JCCo, Job, WipMonth)
```

#### dbo.WipYearEndSnapshot
Populated at December AcctApproved. Source for prior-year columns next year.
```
Id                  INT IDENTITY PK
JCCo                TINYINT
Job                 VARCHAR(50)
SnapshotYear        SMALLINT         -- e.g. 2025
PriorYearGAAPRev    DECIMAL(15,2)    NULL
PriorYearGAAPCost   DECIMAL(15,2)    NULL
PriorYearOpsRev     DECIMAL(15,2)    NULL
PriorYearOpsCost    DECIMAL(15,2)    NULL
BonusProfit         DECIMAL(15,2)    NULL
CreatedAt           DATETIME
UNIQUE (JCCo, Job, SnapshotYear)
```

### Current Data (as of Apr 2, 2026)
- `WipJobData`: 4,974 rows (4 companies, Dec 2024–Dec 2025, from 40 historical files)
- `WipBatches`: empty (no live batches yet — write path not built)
- `WipYearEndSnapshot`: empty

### Stored Procedures

| Procedure | Called From | Purpose |
|-----------|-------------|---------|
| `LylesWIPCreateBatch` | VBA CreateBatch() | INSERT or return existing batch for Co/Month/Dept |
| `LylesWIPGetBatches` | VBA UseExistingBatch() | All batches for Co/Month |
| `LylesWIPCheckBatchState` | VBA GetBatchState() | Current state for Co/Month/Dept |
| `LylesWIPUpdateBatchState` | VBA UpdateBatchState() | Advance state (enforces valid transitions) |
| `LylesWIPSaveJobRow` | VBA SaveJobRow() | MERGE one job row into WipJobData |
| `LylesWIPGetJobOverrides` | VBA BuildOverrideLookup() | All WipJobData rows for Co/Month |
| `LylesWIPClearJobData` | VBA (admin) | DELETE all WipJobData for Co/Month |
| `LylesWIPSaveYearEndSnapshot` | VBA (December AcctApproved) | MERGE year-end snapshot |

### State Transition Rules (in LylesWIPUpdateBatchState)
```
Open         → ReadyForOps    ✓
ReadyForOps  → OpsApproved    ✓
OpsApproved  → AcctApproved   ✓
Any          → Open           ✓ (admin reset)
All other transitions → RAISERROR
```

---

## P&P Server — Permissions DB
**Server:** `10.103.30.11`
**Database:** `PnpMain`
**Stored proc:** `pnp.WIPSECGetRole @UserName VARCHAR` → returns role string

| Role | Access |
|------|--------|
| `WIPAccounting` | All stages: load Vista data, Ready for Ops, GAAP edits, Final Approval |
| `WIPLevel2` | Stage 1 load + Ready for Ops; Stage 2 Ops edits + Ops Approval; no GAAP |
| `WipInitialApproval` | Stage 2: Ops edits and Ops Done only |
| `WipFinalApproval` | Stage 2: Ops Final Approval button only |
| `WipViewOnly` | Read-only all stages |

Currently hardcoded to `WIPAccounting` in Permissions_Modified.bas.
Sprint 5 task: wire to `pnp.WIPSECGetRole` with `Environ("USERNAME")`.

---

## Python Tools — Database Connections

### sql/load_overrides.py
```python
SERVER   = '10.103.30.11'
DATABASE = 'LylesWIP'
USERNAME = 'wip.excel.sql'
PASSWORD = 'WES@2024'
# pyodbc: ODBC Driver 18 for SQL Server, TrustServerCertificate=yes, Encrypt=no
```

### validate_wip.py / validate_*.py
Reads only from local .xltm files via openpyxl (no macros run). No database connection.

### Direct Vista queries (ad-hoc Python sessions)
```python
SERVER   = '10.112.11.8'
DATABASE = 'Viewpoint'
USERNAME = 'WIPexcel'
PASSWORD = 'WIP@MR@2024'   # May have been updated — verify from workbook Settings C9
# pyodbc: ODBC Driver 18 for SQL Server, TrustServerCertificate=yes, Encrypt=no
```
