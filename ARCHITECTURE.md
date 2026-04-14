# Auto-WIP Architecture
*Last updated: 2026-04-08*

---

## Data Flow

```
Nicole's Excel files (F: drive / Data-from-Nicole/)
    │
    ▼  load_overrides.py (one-time historical seed + re-run to add rows)
    │
    ▼
LylesWIP DB  ←──────────────────────────────────────────────────────┐
(10.103.30.11                                                        │
 Cloud-Apps1)                                                        │
    │                                                                │
    │  LylesWIPGetJobOverrides (stored proc)                         │
    │                                                                │
    ▼                                                                │
Excel Workbook (VBA)  ←── GetWIPDataFromVista ──── Vista (10.112.11.8)
    │                      (VistaData.bas / 7-CTE query)
    │
    │  After Vista load:
    │  MergeOverridesOntoSheet merges LylesWIP overrides over Vista values
    │  Z columns (COLZOPsRev, COLZGAAPRev, etc.) = what LylesWIP wrote
    │
    │  On double-click Done (col H or I):
    └──► LylesWIPSaveJobRow  ──► LylesWIP WipJobData   [NOT YET BUILT]
```

**Why P&P and not Vista for writes:** Remote offices and job-site trailers cannot reach Vista SQL Server
(10.112.11.8). All ops users write to LylesWIP on P&P (10.103.30.11), which is accessible everywhere.
Accounting users (Nicole/Cindy) need Vista only for the initial data load.

---

## Three-Stage Workflow

```
STAGE 1 — Accounting Initial Review  (Nicole, corporate — needs Vista access)
  1. Start sheet: select Company + Month + Division(s)
  2. CreateBatch in LylesWIP → batch created in WipBatches (State = Open)
  3. Vista data loads → GetWIPDataFromVista → MergeOverridesOntoSheet
  4. Nicole reviews Jobs-Ops tab (read-only at this stage)
  5. Start sheet → "Ready for Ops: Yes" → LylesWIPUpdateBatchState → ReadyForOps
  6. SaveAs .xlsm + ClearFormOnOpen=False → distribute to Ops PMs/Division Managers

STAGE 2 — Operations Review  (PMs/Division Managers, any location — no Vista needed)
  1. Open distributed .xlsm → connects to LylesWIP only (no Vista connection)
  2. Loads existing batch data + overrides from LylesWIP
  3. Jobs-Ops tab → edit yellow override columns
  4. Double-click col H (Op's Done) per row → LylesWIPSaveJobRow (IsOpsDone=1)
  5. Double-click col G (Close) if job ready to close
  6. Double-click Notes column to enter/collapse notes
  7. All rows Done → "Ops Final Approval: Yes" → LylesWIPUpdateBatchState → OpsApproved

STAGE 3 — Accounting Final Approval  (Nicole/Cindy, corporate)
  1. Open workbook → loads from LylesWIP (all Ops edits visible)
  2. Jobs-GAAP tab → edit GAAP yellow columns (OPS tab is locked at this stage)
  3. Double-click col I (GAAP Done) per row → LylesWIPSaveJobRow (IsGAAPDone=1)
  4. "Accounting Final Approval: Yes" → LylesWIPUpdateBatchState → AcctApproved
  5. Batch locked. If December: LylesWIPSaveYearEndSnapshot populates prior-year baseline.
```

## Batch State Machine

```
Open  ──►  ReadyForOps  ──►  OpsApproved  ──►  AcctApproved
           (RFOYes_Click)   (OFAYes_Click)    (AFAYes_Click)
                                                     │
                                              [If December]
                                                     ▼
                                           WipYearEndSnapshot
```

State transitions enforced by `LylesWIPUpdateBatchState` stored proc.
Reset to Open allowed from any state (admin use only).

---

## Companies and Divisions

| Company | JCCo | Divisions in scope | Notes |
|---------|------|--------------------|-------|
| W. M. Lyles Co. (WML) | 15 | 50, 51, 52, 53, 54, 55, 56, 57, 58 | Div50 = Company overhead |
| Advanced Integration & Controls (AIC) | 16 | 70, 71, 72, 73, 74, 75, 76, 77, 78 | Div70 = Company overhead |
| American Paving Co. (APC) | 12 | 20, 21 | Div20 = General, Div21 = Company |
| New England Sheet Metal (NESM) | 13 | 31, 32, 33, 35 | No Div34 |

**CRITICAL: Division ≠ job number prefix in all cases.**
The WIP query filters by `bJCCM.Department` (contract-level), NOT by job number prefix.
APC jobs `20.252x.`–`20.257x.` have contracts in Dept=21, NOT Dept=20.
Always trust bJCCM.Department, not the leading digits of the job number.

**Sub-jobs:** Only 2 exist across all 4 companies: `56.1009.01` and `56.1010.01` (both WML Div56).
Sub-jobs in Vista bJCJM have NO trailing dot (unlike standard jobs: `56.1004.`).
LylesWIP stores them with trailing dot (`56.1010.01.`). VBA must normalize before lookup:
`If Right(jobNum, 1) <> "." Then jobNum = jobNum & "."`
This fix is in LylesWIPData.bas (deployed Apr 2).

---

## Workbook Structure

| Sheet CodeName | Tab Label | Purpose |
|----------------|-----------|---------|
| Sheet2 | Settings | DB connections, env config — **xlSheetVeryHidden** before delivery |
| Sheet17 | Start | Entry: Company, Month, Division. State machine radio buttons. |
| Sheet11 | Jobs-Ops | Operations view. Col H = Ops Done. Yellow = editable. |
| Sheet12 | Jobs-GAAP | GAAP/Accounting view. Col I = GAAP Done. Yellow = editable. |
| Sheet13 | Jobs-Ops vs GAAP | Comparison view. Not loadable from batch snapshots. |
| Sheet15 | JV's-Ops | Joint Venture Ops. Deferred — confirm JV workflow with Nicole before building. |
| Sheet16 | JV's-GAAP | Joint Venture GAAP. Deferred. |

### Settings Sheet Named Ranges (Sheet2)

| Range | Cell | Value |
|-------|------|-------|
| VistaServerName | C3 | `10.112.11.8` (prod) / `VM111VPPRD1` (by name) |
| VistaDBName | C4 | `Viewpoint` |
| PPServerName | C5 | `10.103.30.11` |
| PPDBName | C6 | `LylesWIP` |
| WIPDBName | C7 | `LylesWIP` |
| VPUsername | C8 | Vista SQL auth user (`WIPexcel`) |
| VPPassword | C9 | Vista SQL auth password |
| PPUsername | C10 | P&P auth user (`wip.excel.sql`) |
| PPPassword | C11 | P&P auth password |
| ClearFormOnOpen | (settings area) | Bool — True in master, False in distributed .xlsm copies |

---

## Override Merge Logic

When `GetWipDetail2(sh)` loads a sheet:
1. Vista query runs → raw financial values populate all standard columns
2. **Override columns default to Vista values** (Rev 5.72p+):
   - `COLOvrRevProj` (Col I): if no override, defaults to ProjContract (or BilledAmt if contract=0)
   - `COLOvrCostProj` (Col M): if no override, defaults to MAX(ProjCost, ActualCost)
   - `COLJTDBonusProfit` (Col Z): if no override, calculated as Michael's formula:
     loss jobs → projected loss; >30% complete → (PctComplete × Revenue) − Cost; else 0
   - This matches Michael's `LCGWIPCreateBatch` which pre-merged defaults into WIPDetail.
3. `BuildOverrideLookup(co, month)` queries `LylesWIPGetJobOverrides` → Scripting.Dictionary keyed by Job (case-insensitive, vbTextCompare)
4. `MergeOverridesOntoSheet(sh)` loops all data rows, looks up each job, writes override values to Z columns:
   - `COLZOPsRev`, `COLZOPsCost`, `COLZGAAPRev`, `COLZGAAPCost`, `COLZOPsBonus`
   - `COLCompDate` (completion date)
5. Z columns = ground truth for what LylesWIP persisted. Visible yellow cells read from Z columns.

**NULL in WipJobData** = no override, use Vista-calculated value.
**Non-NULL with Plugged=1** = user override, written to Z column.

---

## Michael's Original Architecture (WipDb — reference only)

Michael's system used a pre-populated `WIPDetail` table on WipDb (`10.103.30.11`):
1. `LCGWIPCreateBatch` → queries Vista via linked server, creates rows in WIPDetail with defaults
2. `LCGWIPUpdateCostBill` → refreshes cost/billing from Vista into existing WIPDetail rows
3. `LCGWIPGetDetailPM` → reads from WIPDetail (simple SELECT + permissions filter)
4. `LCGWIPUpdateRowOps/GAAP` → saves individual user edits
5. `LCGWIPUploadWIPDetail1` → pushes to Vista `LCGWIPDetail` + calls `LCGWIPMergeDetail`

Key differences from our approach:
- Michael used `bJCCP` (cost by phase) and `bJCIP` (billing by item); we use `bJCCD` and `vrvJBProgressBills`
- Michael's `ContractStatus` mapped to 1/2 only (open/closed at WIP date); we now do the same (Rev 5.72p)
- Michael's `#JobList` filter: Open + SoftClosed always; HardClosed only if MonthClosed in WIP year
- Reference procs saved in `sql/reference/`

---

## Distribution Workflow (to build — Sprint 5)

Problem: `.xltm` distributed to PMs opens blank (template behavior in Windows Explorer).

Fix: "Save & Distribute to Ops" button on Start sheet:
1. `Sheet2.Range("ClearFormOnOpen").Value = False` in the copy
2. `ThisWorkbook.SaveAs filename, xlOpenXMLWorkbookMacroEnabled` (`.xlsm`, NOT SaveCopyAs)
3. Restore `ClearFormOnOpen = True` in Nicole's master
4. Distribution folder: configured as named range in Settings sheet (TBD path)

See `memory/distribution_workflow.md` for full spec.
