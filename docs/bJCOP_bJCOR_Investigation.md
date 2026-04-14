# bJCOP & bJCOR — Vista Write-Back Investigation
*Investigated: April 3, 2026 by Josh Garrison (via Claude Code)*
*Server: 10.112.11.8 (Viewpoint production)*

---

## Summary

Both tables are safe to write to directly. They have no FK constraints and no cascading side effects.
However, they DO have active triggers that **validate referential integrity** and **audit all changes**
to `bHQMA`. Michael Roberts' existing `LCGWIPMergeDetail` proc already writes to both tables
successfully using a two-pass MERGE pattern.

---

## Table Schemas

### bJCOP — Job Cost Override by Period
**Primary Key:** `(JCCo, Job, Month)`
**Row Count:** 5,777

| Column | Type | Nullable | Default | WIP Schedule Mapping |
|--------|------|----------|---------|----------------------|
| JCCo | bCompany (tinyint) | No | — | Company code (15=WML, 16=AIC, 12=APC, 13=NESM) |
| Job | bJob (varchar 10) | No | — | Job number (raw, with trailing dot for standard jobs) |
| Month | bMonth (smalldatetime) | No | — | First of WIP month |
| ProjCost | bDollar (decimal 12,2) | No | 0 | **GAAP Cost Override** |
| OtherAmount | bDollar (decimal 12,2) | No | 0 | **OPS Cost Override** |
| Notes | varchar(max) | Yes | NULL | GAAP Cost Notes |
| udPlugged | bYN (char 1) | Yes | 'N' | 'Y' = user manually entered this value |

### bJCOR — Job Cost Override by Revenue (Contract-level)
**Primary Key:** `(JCCo, Contract, Month)`
**Row Count:** 2,860

| Column | Type | Nullable | Default | WIP Schedule Mapping |
|--------|------|----------|---------|----------------------|
| JCCo | bCompany (tinyint) | No | — | Company code |
| Contract | bContract (varchar 10) | No | — | Contract number (**NOT Job** — important distinction) |
| Month | bMonth (smalldatetime) | No | — | First of WIP month |
| RevCost | bDollar (decimal 12,2) | No | 0 | **GAAP Revenue Override** |
| OtherAmount | bDollar (decimal 12,2) | No | 0 | **OPS Revenue Override** |
| Notes | varchar(max) | Yes | NULL | GAAP Revenue Notes |
| udPlugged | bYN (char 1) | Yes | 'N' | 'Y' = user manually entered this value |

**Key difference:** bJCOP is keyed by **Job**, bJCOR is keyed by **Contract**. Multiple jobs can share one contract.

---

## Triggers (6 Total — All Active, Not Disabled)

### bJCOP Triggers

#### btJCOPi (INSERT)
- **Validates** JCCo exists in `bJCCO` — rollback if not
- **Validates** Job exists in `bJCJM` — rollback if not
- **Audits** insert to `bHQMA` (RecType='A') when `bJCCO.AuditProjectionOverrides = 'Y'`

#### btJCOPu (UPDATE)
- **Blocks** changes to PK columns: JCCo, Job, Month — rollback with error
- **Audits** changes to `ProjCost` to `bHQMA` (RecType='C', old value, new value)
- **Audits** changes to `OtherAmount` to `bHQMA` (RecType='C', old value, new value)
- Audit only fires when `bJCCO.AuditProjectionOverrides = 'Y'`

#### btJCOPd (DELETE)
- **Audits** delete to `bHQMA` (RecType='D')

### bJCOR Triggers

#### btJCORi (INSERT)
- **Validates** JCCo exists in `bJCCO` — rollback if not
- **Validates** Contract exists in `bJCCM` — rollback if not
- **Audits** insert to `bHQMA`

#### btJCORu (UPDATE)
- **Blocks** changes to PK columns: JCCo, Contract, Month
- **Audits** changes to `RevCost` and `OtherAmount`

#### btJCORd (DELETE)
- **Audits** delete to `bHQMA`

### Trigger Implications for Write-Back
1. **INSERT will fail** if Job/Contract doesn't exist in the master table. This is good — it catches bad data.
2. **UPDATE can only change** ProjCost, OtherAmount, Notes, udPlugged. Cannot change PK columns.
3. **All changes are audited** to bHQMA with SQL login name, timestamp, old/new values.
4. **No cascading effects** — triggers don't update other tables (except bHQMA audit log).
5. **Trigger error on zero-row MERGE INSERT** — Michael split his MERGE into two passes to work around this.

---

## Audit Configuration

```sql
-- Company 15 (WML) has auditing enabled:
SELECT JCCo, AuditProjectionOverrides FROM bJCCO WHERE JCCo = 15
-- Result: JCCo=15, AuditProjectionOverrides='Y'
```

Most recent audit entries (from bHQMA):
- **2026-03-31** — `nleasure@lylessc.com_23418` updated bJCOR RevCost (0 → 126,483,581) for job 51.1129, Dec 2025
- **2026-03-31** — `nleasure@lylessc.com_23418` updated bJCOP ProjCost (0 → 115,233,581) for job 51.1129, Dec 2025
- **2026-03-26** — `bpoochigian@lylessc.com_23418` inserted rows in bJCOP for Co 12 jobs

Nicole and her team are actively writing to these tables through the Viewpoint UI.

---

## Foreign Key Constraints

**None.** Zero FK constraints on either table. Referential integrity is enforced entirely by the
INSERT triggers (standard Viewpoint architecture — triggers serve as the constraint layer).

---

## All Objects Referencing These Tables

Found via `sys.syscomments` search:

| Object | Type | Read/Write | Purpose |
|--------|------|------------|---------|
| `JCOP` | View | Read | Security-filtered view — wraps bJCOP with `vfDataTypeSecurity` check |
| `JCOR` | View | Read | Security-filtered view — wraps bJCOR with `vfDataTypeSecurity` check |
| `JCOverridesCost` | View | Read | Reporting/query view |
| `LCGWIPMergeDetail` | Stored Proc | **Write** | The write path — MERGE from LCGWIPDetail staging into bJCOP/bJCOR + LCGWIPSchedule |
| `LCGRCSMergeJCOP` | Stored Proc | **Write** | Older Michael write proc (via global temp table). Same MERGE pattern. |
| `LCGRCSMergeJCOR` | Stored Proc | **Write** | Older Michael write proc (via global temp table). Same MERGE pattern. |
| `LCGMoJobSum2` | Stored Proc | Read | Monthly job summary report |
| `LCGMoJobSum3` | Stored Proc | Read | Monthly job summary report |
| `LCGMoJobSum4` | Stored Proc | Read | Monthly job summary report |
| `bspJCInitOverrides` | Stored Proc | **Write** | Viewpoint built-in — seeds override rows when new month initialized (INSERT IF NOT EXISTS only) |
| `btJCOPi/u/d` | Triggers | — | Validation + audit (see above) |
| `btJCORi/u/d` | Triggers | — | Validation + audit (see above) |

---

## LCGWIPMergeDetail — How the Existing Write Path Works

This is Michael Roberts' proc that writes approved GAAP values from the WIP staging table
(`LCGWIPDetail`) into bJCOP/bJCOR. Created 2024-06-14.

### Parameters
```sql
@Co    INT           -- Company
@Dept  VARCHAR(200)  -- Comma-separated dept list
@Month DATE          -- WIP month
@rcode INT OUTPUT
@msg   VARCHAR(500) OUTPUT
```

### Logic Flow

1. **Parse dept list** into `#DeptList` temp table (XML split on commas)

2. **GAAP-only guard:** `IF MONTH(@Month) % 3 = 0` — bJCOP/bJCOR writes ONLY run on quarterly months (March, June, September, December). Non-GAAP months skip entirely.

3. **bJCOP — Two-pass MERGE** (cost overrides):
   - Pass 1 (UPDATE existing): `MERGE INTO bJCOP ... WHEN MATCHED THEN UPDATE SET ProjCost, OtherAmount, Notes, udPlugged`
   - Pass 2 (INSERT new): `MERGE INTO bJCOP ... WHEN NOT MATCHED THEN INSERT (JCCo, Job, Month, ProjCost, OtherAmount, Notes, udPlugged)`
   - Source: `LCGWIPDetail.GAAPCost → ProjCost`, `GAAPOtherCostAmount → OtherAmount`, `GAAPCostNotes → Notes`, `GAAPCostPlugged → udPlugged`
   - **Join key**: `M.JCCo = T.JCCo AND M.Job = T.Contract AND M.Month = T.Month`
   - Note: uses `T.Contract` (from LCGWIPDetail) as the Job value — naming is confusing but correct

4. **bJCOR — Two-pass MERGE** (revenue overrides):
   - Same two-pass pattern
   - Source: `LCGWIPDetail.GAAPRev → RevCost`, `GAAPOtherRevAmount → OtherAmount`, `GAAPRevNotes → Notes`, `GAAPRevPlugged → udPlugged`
   - **Join key**: `M.JCCo = T.JCCo AND M.Contract = T.Contract AND M.Month = T.Month`

5. **LCGWIPSchedule — MERGE** (ops override archive — runs always, not just quarterly):
   - MERGE by `(Co, Month, Job=Contract)`
   - Stores: CompletionDate, Dept, Estimator, PM, ProjectExec, OpsRev (ContractAdj), OpsCost (CostAdj), notes, completion flags, plug flags, BonusProfit

6. **Cleanup:** `DELETE FROM LCGWIPDetail` for the processed Co/Month/Depts

### Why Two-Pass MERGE?
Comment in proc: `NOTE: Merge broken into two parts for JCOR and JCOP due to trigger error if not inserting any records`

The btJCOPi/btJCORi INSERT triggers validate row counts. A single MERGE statement that matches
all rows (UPDATE only, zero inserts) still fires the INSERT trigger with `@@rowcount = 0`, which
the trigger handles correctly (`if @numrows = 0 return`). However, Michael found this caused errors
in practice, so he split UPDATE and INSERT into separate MERGE statements.

---

## LCGRCSMergeJCOP / LCGRCSMergeJCOR — Older Write Procs

These are Michael's earlier versions. Instead of reading from `LCGWIPDetail`, they accept a global
temp table name (`@MyTmpTable`) and build dynamic SQL to MERGE from it. Same two-pass MERGE pattern.
These are the procs called by the old `UploadData.bas` → `CopyWIPDetail` path (via
`LCGRCSUpdatebudWIPDetail`, which is commented out).

---

## bspJCInitOverrides — Viewpoint Built-in

Viewpoint's own proc for initializing override rows when a new fiscal period is opened.
- INSERTs into bJCOP and bJCOR via the `JCOP`/`JCOR` security views
- Only inserts if row does NOT already exist (`NOT EXISTS` check)
- No conflict with WIP writes — it only creates seed rows with default values

---

## Implications for Our Write-Back

### What's Safe
- Direct INSERT/UPDATE to bJCOP and bJCOR is proven safe — Michael's procs do it, Nicole's team does it through Viewpoint UI
- Triggers handle validation (bad Job/Contract = automatic rollback with error message)
- Audit trail is automatic and useful (tracks who changed what, when)
- No cascading side effects beyond audit logging

### What to Watch For
1. **bJCOP.Job must exist in bJCJM** — trigger validates on INSERT
2. **bJCOR.Contract must exist in bJCCM** — trigger validates on INSERT
3. **udPlugged must be set correctly** — 'Y' for user overrides, 'N' for system values. All read queries filter on `udPlugged = 'Y'`
4. **GAAP quarterly only** — only write bJCOP/bJCOR on Mar/Jun/Sep/Dec months
5. **Two-pass MERGE pattern** — don't combine UPDATE and INSERT in one MERGE (trigger issue)
6. **Contract vs Job confusion** — `LCGWIPDetail.Contract` column is used as bJCOP.Job AND bJCOR.Contract. The naming in Michael's staging table is misleading.

### Recommended Write Approach
Follow Michael's `LCGWIPMergeDetail` pattern exactly:
1. Stage data in LCGWIPDetail (or our LylesWIP.WipJobData equivalent)
2. Two-pass MERGE for bJCOP (UPDATE existing, then INSERT new)
3. Two-pass MERGE for bJCOR (UPDATE existing, then INSERT new)
4. MERGE to LCGWIPSchedule for ops archive
5. Clean up staging data

### SQL Login for Writes
The `WIPexcel` account currently has at minimum SELECT + EXECUTE permissions on these objects.
Confirm it also has INSERT/UPDATE on `bJCOP` and `bJCOR` before building the write path.
Michael's GRANT comments suggest `public` role has EXECUTE on the LCG procs, so the procs
themselves may run under a different security context (check `EXECUTE AS` or proxy permissions).
