# VBA Module Inventory
*Last updated: 2026-04-02*
*Active workbook: `vm/WIPSchedule - Rev 5.68p.xltm`*

Files ending in `_Modified.bas` = current deployed version. Originals kept for diff reference only.

---

## Deployed Modules (in workbook as of Apr 2)

### VistaData.bas ✅ Deployed
**Purpose:** Opens ADODB connection to Vista (10.112.11.8/Viewpoint) and executes the 7-CTE WIP read query.
**Key functions:**
- `OpenVistaConnection() → Boolean` — connects using Settings C3/C4/C8/C9
- `CloseVistaConnection()` — closes module-level `mVistaConn`
- `GetWIPDataFromVista(co, month, deptList) → ADODB.Recordset` — executes `WIP_Vista_Query.sql` logic inline
- `GetCompanyList() → ADODB.Recordset` — queries `bHQCO` for company dropdown
- `GetDeptList(co) → ADODB.Recordset` — queries `bJCDM` for division dropdown
- `GetJVDataFromVista(co, month) → ADODB.Recordset` — queries JV tables; includes guard for missing `udWIPJV` on prod
**Known issues:** None. `udWIPJV` guard deployed — prod server lacks this table.
**Performance rule:** Raw `Job` field in all JOINs/GROUP BY. Trim only in final SELECT.

### GetWIPDetailData_Modified.bas ✅ Deployed
**Purpose:** Orchestrates data load for a single sheet (Jobs-Ops or Jobs-GAAP). Called as `GetWipDetail2(sh)`.
**Flow:** calls `GetWIPDataFromVista` → writes rows to sheet → calls `BuildOverrideLookup` → calls `MergeOverridesOntoSheet`
**Key function:** `GetWipDetail2(sh As Worksheet)` — main entry point from Start sheet load button
**Known issues:** None.

### LylesWIPData.bas ✅ Deployed (⚠️ needs re-import on VM — trailing-dot fix added Apr 2)
**Purpose:** Connects to LylesWIP on P&P server. Handles override merge on load. Write-back functions stub-ready.
**Key functions — read side (complete):**
- `OpenWIPConnection() → Boolean` — connects using Settings C5/C6/C10/C11
- `CloseWIPConnection()`
- `BuildOverrideLookup(co, month) → Scripting.Dictionary` — calls `LylesWIPGetJobOverrides`, keys by Job (vbTextCompare)
- `MergeOverridesOntoSheet(sh)` — writes Z columns (COLZOPsRev etc.) from override dict
- `MergePriorMonthProfitsOntoSheet(sh)` — merges prior-month Col P values
- `MergePriorYearBonusOntoSheet(sh)` — merges prior-year bonus values
**Key functions — write side (NOT YET BUILT — add to this module):**
- `SaveJobRow(co, job, month, ...)` → calls `LylesWIPSaveJobRow` stored proc
- `CreateBatch(co, month, dept, userName)` → calls `LylesWIPCreateBatch`
- `GetBatchState(co, month, dept) → String` → calls `LylesWIPCheckBatchState`
- `UpdateBatchState(co, month, dept, newState, userName)` → calls `LylesWIPUpdateBatchState`
**Known issues/quirks:**
- Sub-jobs (`56.1010.01`) in Vista have no trailing dot. Fix deployed: `If Right(jobNum,1)<>"." Then jobNum=jobNum&"."` in all 3 Merge* loops before dictionary lookup.
- `SummaryData` named range is 18 rows in template but GetWipDetail2 can write more. Fix deployed: `xlUp` scan determines actual last data row; loop bound = `Max(Rows.Count, lastDataRow-range.Row+1)`.
- `LylesWIPGetJobOverrides` has NO department filter — returns all jobs for co+month. VBA filters by sheet content naturally (only merges rows already on the sheet from Vista).

### Module1_Modified.bas ✅ Deployed
**Purpose:** Populates Company dropdown on Start sheet.
**Key function:** `GetCompanyData()` — queries `bHQCO` via Vista for company list
**Known issues:** None.

### Module2_Modified.bas ✅ Deployed
**Purpose:** Populates company name display on Start sheet.
**Key function:** `GetCoName(co) → String` — queries `bHQCO` via Vista
**Known issues:** None.

### GetudWIPJV_Modified.bas ✅ Deployed
**Purpose:** Loads JV data for JV's-Ops and JV's-GAAP sheets.
**Key function:** `GetJVData(sh)` — queries Vista JV tables; guards against missing `udWIPJV`.
**Known issues:** `udWIPJV` does not exist on production Vista (10.112.11.8). Guard deployed returns empty set gracefully.

### Module6_Phase1_Complete.bas ✅ Deployed (as Module6)
**Purpose:** Batch creation flow on Start sheet. All write functions guarded with Exit Sub (Sprint 1 guard).
**Key functions:**
- `UseExistingBatch()` — checks for existing batch; currently dept-picker only (write path not yet wired)
- `CreateBatch()` — guarded: shows MsgBox and exits; to be replaced with LylesWIP call in Sprint 4
- `UpdateRow()`, `UpdateRowJV()`, `UpdateApprovals()` — all guarded with Exit Sub
- `GLCheck(co, month)` → queries `bGLCO.LastMthSubClsd` from Vista (real GL closed-month check)
- `CompleteCheck()` — returns True always (stub; real check in Sprint 5)
**Known issues:** All write-path functions guarded. Must be unwrapped/replaced in Sprint 4.

### ClearFormsMod.bas ✅ Deployed
**Purpose:** Form clear/reset utilities.
**Key functions:**
- `ClearForms3()` — clears sheet data but leaves sheets protected (⚠️ do NOT use between BatchValidate iterations — causes 1004 error on next iteration's protect/unprotect cycle)
- `ResetWorkbook()` — full clean reset (use this in BatchValidate between iterations)
**Known issues:** `ClearForms3` leaves sheets in protected state. Always use `ResetWorkbook` in automation.

### ThisWorkbook_Modified.cls ✅ Deployed
**Purpose:** Workbook open/close events.
**Key events:**
- `Workbook_Open()` — checks `ClearFormOnOpen` named range; clears form if True
- `Workbook_BeforeClose()` — calls `CloseVistaConnection()` + `CloseWIPConnection()`
**Known issues:** `ClearFormOnOpen` = True in Nicole's master, False in distributed .xlsm copies (distribution workflow builds this).

### Permissions_Modified.bas ⚠️ Ready to import — NOT YET IN WORKBOOK
**Purpose:** Role-based access control. Current version hardcodes `WIPAccounting` for all users.
**Key function:** `GetSecurity() → String` — currently returns "WIPAccounting" always
**To deploy:** Import this module to replace the original Permissions.bas. Eliminates "Security Settings Not Valid" popup on open.
**Future work (Sprint 5):** Wire to `pnp.WIPSECGetRole` stored proc on PnpMain DB for real role lookup.

### FormButtons.bas ✅ Deployed
**Purpose:** Start sheet button click handlers.
**Key subs:** `RFOYes_Click`, `RFONo_Click`, `OFAYes_Click`, `AFAYes_Click` — state machine transitions (currently partially guarded; write-back to wire in Sprint 5)
**Known issues:** State machine wired partially. Full wiring in Sprint 5.

### SetColumnRanges.bas ✅ Deployed
**Purpose:** Defines the `NumDict` dictionary — maps sheet CodeName to column index for every named column. This is the column mapping table.
**Usage:** `NumDict(sh.CodeName)("COLJobNumber")` returns the column index for job number on that sheet.
**Known issues:** None. The authoritative column map for VBA.

### UploadData.bas ✅ Deployed (guarded)
**Purpose:** Originally sent data to WipDb. All functions guarded with Exit Sub.
**Key functions guarded:** `SendToVista()`, `CopyWIPDetail()` — both Exit Sub immediately.

### Module3.bas ✅ Deployed (guarded)
**Purpose:** Originally called `UpdateCostBill`. Guarded with Exit Sub.

### Module4.bas ✅ Deployed
**Purpose:** Utility functions (date helpers, formatting). No write-path content.

### UserAccountInfo.bas ✅ Deployed
**Purpose:** Returns `Environ("USERNAME")` for audit trail. Used by GetSecurity and SaveJobRow calls.

---

## Batch & Validation Tools (not in workbook — Python/standalone)

### BatchValidate.bas ✅ Deployed (⚠️ needs re-import on VM — now 24 combos)
**Purpose:** Unattended snapshot generation for all company/division combinations.
**Key sub:** `BatchValidateAll()` — loops 24 combos, calls `ResetWorkbook` + `GetWipDetail2` + `SaveCopyAs`
**Output:** `C:\Trusted\validate-d3\{co}-{div}.xltm` (24 files)
**Combos:** WML Div 50–58 (9), AIC Div 70–78 (9), APC Div 20–21 (2), NESM Div 31/32/33/35 (4)
**Known issues:** Uses `ResetWorkbook` NOT `ClearForms3` between iterations. ClearForms3 leaves sheets protected → 1004 error.

### validate_wip.py ✅ Complete
**Purpose:** Compares Nicole's Dec 2025 WIP History Import files against 24 workbook snapshots.
**Input:** `Data-from-Nicole/` (Nicole's files), `vm/validate-d3/` (snapshots)
**Output:** `vm/validation_report.xlsx` — 5 sheets: Summary, Mismatches, All Override Checks, Nicole Jobs Not In WB, Notes
**Run:** `python3 validate_wip.py` from repo root
**Known issues:** None. See VALIDATION.md for current results.

### sql/load_overrides.py ✅ Complete
**Purpose:** Reads all 40 Nicole Excel files (Revenue + Cost sheets) → MERGE into LylesWIP.WipJobData.
**Run:** `python3 sql/load_overrides.py` (or `--dry-run`)
**normalize_job():** Right-pads job segment to 4 digits (`73.105.` → `73.1050.`); adds trailing dot.
**Known issues:** Does NOT handle sub-jobs specifically — `56.1010.01` becomes `56.1010.01.` (trailing dot added). This matches the VBA fix in LylesWIPData.bas.

---

## Sheet Class Modules (event handlers)

| Module | Sheet | Key Events |
|--------|-------|-----------|
| Sheet11.cls | Jobs-Ops | `BeforeDoubleClick` — col H = Ops Done (write-back to wire) |
| Sheet12.cls | Jobs-GAAP | `BeforeDoubleClick` — col I = GAAP Done (write-back to wire) |
| Sheet17.cls | Start | Form control events for company/month/division selection |
| Sheet2.cls | Settings | Protected — no events |
| Sheet13–16.cls | Comparison/JV | Minimal or placeholder events |

---

## Archived / Superseded Files (do not import these)
| File | Why Superseded |
|------|---------------|
| `GetWIPDetailData.bas` | Original — WipDb-connected version |
| `GetudWIPJV.bas` | Original — WipDb-connected version |
| `Module1.bas`, `Module2.bas` | Original — WipDb-connected versions |
| `Permissions.bas` | Original — has `pnp.WIPSECGetRole` call that causes popup |
| `Module6.bas` | Original — unguarded WipDb write calls |
| `VistaData_BatchAdditions.bas` | Superseded planning artifact — functionality folded into VistaData.bas |
| `FormButtons_RFOYes_Updated.bas` | Draft — partially implemented, not deployed |
| `Module6_*_Replacement.bas` | Sprint planning drafts — reference only |
| `ThisWorkbook.cls` | Original — has C1 bug (`uname="cjordan"`) |
