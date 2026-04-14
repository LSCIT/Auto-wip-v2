# Deploy Guide — WIPSchedule Rev 5.52p

## What Changes vs Original

**Read path → Vista** (all data comes from Vista/Viewpoint directly):
- Company/dept dropdowns, GL closed-month check, WIP data load, JV data load

**Write path → unchanged** (CreateBatch, UpdateRow, UpdateApprovals, etc. still use WipDb as original):
- This is correct — Nicole/Cindy's changes still save back through WipDb

**Formula patches** already in the workbook (Bugs 1+2 for GAAP formulas in Jobs-GAAP tab).

---

## Current State of 5.52p
- ✅ VistaData.bas — deployed
- ✅ Formula patches (Bug 1: Col W threshold, Bug 2: Col R Prior Proj Profit)
- ❌ Module1, Module2, Permissions, GetWIPDetailData, GetudWIPJV, ThisWorkbook — still original
- ❌ Module6 — UseExistingBatch and GLCheck still point to WipDb

---

## Deploy Procedure (Windows, Excel VBA Editor)

Open `vm/WIPSchedule - Rev 5.52p.xltm` → **Alt+F11**

---

### Full Module Replacements

For each:
1. Click the module name in the Project pane
2. **Ctrl+A → Delete**
3. Copy the full contents of the source file and **Paste**
4. If first line is `Attribute VB_Name = "..."` — **delete that line** (VBA adds it automatically)

| VBA Module Name | Source File | What Changed |
|---|---|---|
| `Module1` | `vba_source/Module1_Modified.bas` | GetCompanyData/GetDeptData → Vista (bHQCO, bJCDM) |
| `Module2` | `vba_source/Module2_Modified.bas` | GetCoName → Vista bHQCO |
| `Permissions` | `vba_source/Permissions_Modified.bas` | GetSecurity → hardcoded WIPAccounting |
| `GetWIPDetailData` | `vba_source/GetWIPDetailData_Modified.bas` | Main WIP data load → Vista |
| `GetudWIPJV` | `vba_source/GetudWIPJV_Modified.bas` | JV data → Vista |
| `ThisWorkbook` | `vba_source/ThisWorkbook_Modified.cls` | Calls CloseVistaConnection on Workbook_BeforeClose |

> **ThisWorkbook** is under *Microsoft Excel Objects* in the project tree, not Modules.

---

### Surgical Module6 Patches (replace 2 subs only — everything else untouched)

**All write-path functions (CreateBatch, UpdateRow, UpdateRowJV, UpdateApprovals, CompleteCheck, MarkAllComplete, ApprCheck2) stay as Michael wrote them. Only these 2 subs get replaced:**

#### Sub 1: UseExistingBatch (top of Module6, ~line 2)

1. In Module6, find `Public Sub UseExistingBatch(ByRef NoData As Boolean)` (line ~2)
2. Select from that line down to and including its `End Sub` (~line 137)
3. Delete the selection
4. Paste the entire contents of `vba_source/Module6_UseExistingBatch_Replacement.bas`
   (Ignore the comment block header — paste the whole file including the comment block)

**What this does:** Removes the WipDb batch-check (LCGWIPBatchCheck1) and just shows the dept picker form directly.

#### Sub 2: GLCheck (around line 1412, after removing the above)

1. In Module6, find `Public Sub GLCheck()` (line numbers shift after the replacement above)
2. Select from that line down to and including its `End Sub`
3. Delete the selection
4. Paste the entire contents of `vba_source/Module6_GLCheck_Replacement.bas`

**What this does:** Removes the WipDb LCGWIPGLCheck call and queries `bGLCO.LastMthSubClsd` from Vista instead.

---

## Verification After Deploy

1. Close and reopen the workbook (or File → Close and reopen)
2. Start sheet → press **F4** in cell B3 (company picker)
   - Should show company list from Vista — 14 companies
   - No "There are no Companies Available" error
3. Select **Company 15** (WML), set Month to **12/1/2025**, pick **Dept 54**
4. Click **Get WIP Data**
   - All 4 tabs should populate — expect ~130+ jobs for Dept 54
5. **Spot-check Col W** (Jobs-GAAP, JTD Earned Rev): find a job with < 10% completion
   - Should show JTD Cost value, not (% × contract)
6. **Spot-check Col R** (Prior Proj Profit): find an in-progress job
   - Should show a non-zero value equal to bJCOR prior quarter value

---

## If You See Macro Trust / Security Errors

This is an Excel Trust Center issue (not code):
1. Right-click the .xltm file → **Properties** → check **Unblock** at bottom if present
2. File → Options → Trust Center → Trust Center Settings → Macro Settings → **Enable all macros**
3. Or add the folder to Trusted Locations

---

## Post-Deploy: Re-hide Settings Sheet

Before delivering to Nicole/Cindy, hide the Settings sheet:
1. Alt+F11 → Immediate Window (Ctrl+G)
2. Type: `Sheet2.Visible = xlSheetVeryHidden` → Enter
3. Save the workbook

---

## Expected Result

- Company/dept dropdowns pull live from Vista
- GL closed-month indicator correct (checks `bGLCO.LastMthSubClsd`)
- All 4 WIP tabs load from Vista directly — no WipDb dependency on read path
- All 4 Nicole GAAP bug fixes active (Bugs 1+2 formulas, Bugs 3+4 in SQL CTEs)
- Write-back path (save changes, approvals) works exactly as before via WipDb
