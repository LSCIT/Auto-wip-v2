# Session Notes — March 5, 2026
# Auto-WIP Phase 1: Full Deployment & Testing

## Session Goal
Pick up from prior sessions, get the entire WIP Schedule workbook pulling live data from Vista/Viewpoint, and eliminate all WipDb (middleman database) dependencies.

---

## What We Accomplished

### 1. All Data-Path Modules Deployed to VM
Every module that touches the read/data-load path is now live on the VM workbook (`vm/WIPSchedule_Unlocked1.xltm`):

| Module | What Changed |
|--------|-------------|
| **VistaData** | Central Vista connection + SQL query builder. Added `GetVistaCompanyList()`, `GetVistaCompanyName()`, `GetVistaDepartmentList()`, `GetJVDataFromVista()` |
| **Module1** | `GetCompanyData` and `GetDeptData` now query Vista `bHQCO`/`bJCDM` instead of WipDb stored procs |
| **Module2** | `GetCoName` rewired to Vista `bHQCO` |
| **Permissions** | `GetSecurity` hardcodes `WIPAccounting` role (bypasses unreachable PnpMain) |
| **GetWIPDetailData** | Main data load rewired to use `BuildWIPQuery` against Vista JC tables |
| **GetudWIPJV** | JV data from Vista `udWIPJV` + `bJCCD`/`bJCCM` join |
| **Module6 (GLCheck)** | Queries Vista `bGLCO.LastMthSubClsd` instead of WipDb `LCGWIPGLCheck` |
| **Module6 (UseExistingBatch)** | Bypasses WipDb batch management, shows dept picker directly |

### 2. Phase 1 Write Guards Deployed
All write-path functions (anything that would try to save/update/upload to WipDb) now have guards:

**Module6:**
- `CreateBatch` — `Exit Sub`
- `ClearWIPDetail` — MsgBox "not available" + `Exit Sub`
- `UpdateRow` — Silent `Exit Sub` (fires on every cell edit)
- `UpdateRowJV` — Silent `Exit Sub`
- `UpdateApprovals` — `Exit Sub`
- `CompleteCheck` — Returns `True` always + `Exit Function`
- `MarkAllComplete` — `Exit Function`
- `ApprCheck2` — `Exit Sub`

**Module3:**
- `UpdateCostBill` — `Exit Sub`

**UploadData:**
- `SendToVista` — MsgBox "not available" + `Exit Sub`
- `CopyWIPDetail` — MsgBox "not available" + `Exit Sub`

### 3. ResetWorkbook Button Added
Created `ResetWorkbook` sub in `ClearFormsMod` — one-click full reset for testing:
- Clears all data from Sheets 11-16
- Resets Start page fields (Company, Month, Division)
- Resets approval radio buttons
- Closes Vista connection
- Handles sheet protection/unprotection
- Button added to Start sheet (Sheet17) via VBA Immediate Window

### 4. Full End-to-End Test Passed
- Company: **15 (WML)**
- Month: **12/1/2025** (shows green — Dec is open per Vista bGLCO)
- Division: **54**
- All 4 data sheets loaded successfully: Jobs-Ops, Jobs-GAAP, JV's-Ops, JV's-GAAP

### 5. Bugs Fixed This Session
- **`Dim T As Long` compile error** in GetudWIPJV — VBA implicit Variant typing conflict. Fixed by removing the explicit Dim.
- **"Can't execute code in break mode"** — stale debug state from prior error. Fixed with Reset button in VBA editor.
- **Guard placement inconsistency** — guards must go AFTER the `End If` of the error control block, not before.
- **C1 bug (uname="cjordan")** — confirmed already removed in prior session, no action needed.
- **SQL account password change** — updated on Settings sheet (C9/VPPassword) mid-session.

---

## Current State of the Workbook
- **Zero WipDb dependency** — read path uses Vista, write path is guarded
- **All data loads live from Vista/Viewpoint** on jerc-sql.viewpointdata.cloud,4730
- **Backup copy saved** by Josh
- **Settings sheet** is currently visible (needs to be re-hidden before delivery)

---

## Architecture Discussion: Web App for Phase 2+
Josh asked about long-term alternatives to Excel/VBA. Agreed a web app is the right move:

**Why:**
- Eliminates VBA maintenance (34 modules, 11K+ lines of spaghetti)
- Server-side credentials (not stored in spreadsheet cells)
- Multi-user access without file sharing
- Real-time dashboards for execs
- Deploy once, access anywhere

**Recommended Stack:**
- ASP.NET API backend (matches PNP4/PNP5 environment)
- React or Blazor frontend with data grid
- Export-to-Excel button for Nicole/Cindy
- Direct Vista SQL Server connection

**Strategy:** Finish Phase 1 in Excel to prove numbers match → earn trust → "same numbers, better platform" for Phase 2.

---

## Next Steps (Priority Order)

1. **Validate Nicole's 11 discrepancies** — Compare WIP Schedule output against her Jan 7 email findings using Dec 2025 data. This is what proves we got it right.

2. **Re-hide Settings sheet** — Run `Sheet2.Visible = xlSheetVeryHidden` in Immediate Window before handing off.

3. **Save as "WIP Phase 1"** — Clean copy for Cindy Jordan (CFO) and Nicole Leasure (Corporate Controller) to validate.

4. **Begin web app architecture** — Sketch data model, API endpoints, and UI wireframes for Phase 2.

---

## Key Reference Files
- `vba_source/VistaData.bas` — All Vista connection and query logic
- `vba_source/COLUMN_MAPPING.md` — Rosetta Stone: proc fields to sheet columns to Vista tables
- `DEPLOYMENT_GUIDE.md` — Step-by-step deployment instructions
- `sql/WIP_Vista_Query.sql` — Standalone Vista query (446 lines, 7 CTEs)
- `sql/Validation_Tests.sql` — 8 test queries for Nicole's discrepancies
- `.claude/projects/.../memory/MEMORY.md` — Persistent project memory for Claude sessions
