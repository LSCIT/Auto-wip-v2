# Auto-WIP Web Application — Product Requirements Document
*Version 1.0 — April 2026*
*Author: Josh Garrison, Director of Technology Innovation*

---

## 1. Background

Lyles Services Co. runs a monthly Work in Progress (WIP) schedule process across four companies
(WML, AIC, APC, NESM). The current tool is an Excel workbook with ~2,000 lines of VBA that
connects to two SQL Server databases (Vista/Viewpoint for read, LylesWIP for read/write).

After 18+ months of development and a working release (Rev 5.73p), the Excel/VBA architecture
has proven brittle: COM automation is fragile, distribution of `.xlsm` files to Project Managers
is manual and error-prone, credentials are embedded in a hidden worksheet, and every change
requires re-importing VBA modules to the workbook on a Windows VM. The process works, but
maintenance cost is disproportionately high for what is fundamentally a data-entry workflow.

This document describes Phase 1 of a web application that replaces the Excel workbook with
feature-identical functionality delivered through a browser.

---

## 2. Goals

| # | Goal |
|---|------|
| G1 | Replace the Excel/VBA tool with a browser-based application that covers the same three-stage WIP workflow |
| G2 | Eliminate manual `.xlsm` distribution — Stage 2 users log in and see their pending work |
| G3 | Replace hidden-worksheet credentials with Entra SSO + server-side SQL auth |
| G4 | Automatically notify the right people when a batch advances between stages |
| G5 | Allow Nicole and Cindy to export any WIP grid to Excel on demand |
| G6 | Preserve the existing LylesWIP database schema and stored procedures without modification |

## 3. Non-Goals (Phase 1)

- No new features beyond 5.73p parity
- No mobile or tablet support (laptop/desktop only)
- No JV sheets (JV's-Ops / JV's-GAAP — deferred, same as Excel)
- No real-time collaborative editing (last-save-wins is acceptable for Phase 1)
- No changes to Vista/Viewpoint database or stored procedures
- No replacement of the LylesWIP stored procedures (`LylesWIPSaveJobRow`, `LylesWIPCreateBatch`, etc.)
- No admin UI for managing `WipUserPermissions` (managed directly in SQL for Phase 1)

---

## 4. Users and Roles

Authorization is sourced from the existing `WipUserPermissions` table in LylesWIP.
Microsoft Entra handles authentication (identity). The application maps the authenticated
Entra UPN to a `WipUserPermissions` row to determine what the user can see and do.

| Role | Who | What they can do |
|------|-----|-----------------|
| **Accounting** | Nicole Leasure, Cindy Jordan, Brian Platten, Harbir Atwal | Create batches, load Vista data, advance state to ReadyForOps; Stage 3: edit GAAP yellow columns, mark GAAP Done, give Accounting Final Approval |
| **Operations** | Project Managers, Division Managers | Stage 2 only: edit Ops yellow columns, mark Ops Done, flag job Close, give Ops Final Approval |
| **Admin** | Josh Garrison | All of the above plus: reset batch state, view all companies/divisions |

A user's permitted company/division combinations are stored in `WipUserPermissions` and enforced
server-side on every API request. An Ops user for WML Div51 cannot read or write WML Div52
batches.

---

## 5. Technical Architecture

### 5.1 Stack

| Layer | Technology |
|-------|-----------|
| Frontend | React (Vite), ag-Grid (existing license), Axios |
| Backend | ASP.NET Core 8 Web API |
| Auth | Microsoft Entra ID via `Microsoft.Identity.Web` (MSAL) |
| ORM / Data | `Microsoft.Data.SqlClient` — direct ADO.NET; call existing stored procs |
| Vista read | Direct SQL connection to Vista (`10.112.11.8`) from the web server |
| LylesWIP read/write | Direct SQL connection to LylesWIP on same host (`cloud-apps1.ad.lylesgroup.com`) |
| Excel export | `ClosedXML` or `EPPlus` (server-side generation) |
| Email | SMTP via Exchange / Microsoft 365 (or SendGrid if relay is unavailable) |
| Hosting | IIS on `cloud-apps1.ad.lylesgroup.com` |

### 5.2 Deployment Topology

```
Browser (corporate laptop)
    │  HTTPS
    ▼
IIS — cloud-apps1.ad.lylesgroup.com
    ├── React SPA  (static files served by IIS)
    └── ASP.NET Core 8 Web API
            ├── LylesWIP DB  (localhost / same host — low latency)
            └── Vista DB     (10.112.11.8 — read-only, Stage 1 Accounting only)
```

Vista read access is only exercised during Stage 1 batch creation (same network, same as today).
Operations users never trigger a Vista query.

### 5.3 Authentication Flow

1. User visits the app — redirect to Entra login if no token
2. Entra returns JWT; `Microsoft.Identity.Web` validates it on every API request
3. The API reads `User.Identity.Name` (UPN), looks up `WipUserPermissions`, and attaches
   a `WipPrincipal` (companies, divisions, role) to the request context
4. All authorization checks are performed against `WipPrincipal`, never client-supplied values

---

## 6. Feature Requirements

### 6.1 Navigation and Batch Selection

- After login, user lands on a **Batch Dashboard** showing all batches they have permission to see
- Batches are listed by Company / Month / Division with current state badge:
  `Open` · `Ready for Ops` · `Ops Approved` · `Acct Approved`
- Accounting users see a **New Batch** button; Ops users do not
- Filtering by company, month, and state is supported

### 6.2 Stage 1 — Accounting: Create Batch and Load Data

*Matches the Start sheet + Vista load flow in Excel.*

**Create Batch**
- Accounting user selects: Company, WIP Month, Division(s)
- App calls `LylesWIPCreateBatch` stored proc → batch record created in `WipBatches` (`State = Open`)
- App immediately queries Vista (`GetWIPDataFromVista` logic) for the selected company/month/division(s)
- Vista data is written to a server-side staging structure and returned to the ag-Grid for display
- `LylesWIPGetJobOverrides` is called; overrides are merged over Vista values (same logic as `MergeOverridesOntoSheet`)
- The merged result is shown in the **Jobs-Ops** grid (read-only at this stage for Ops columns)

**Ready for Ops**
- Accounting user reviews the grid, then clicks **Ready for Ops**
- App calls `LylesWIPUpdateBatchState` → `State = ReadyForOps`
- Email notification sent to all Ops users permitted for the batch's company/division (see §6.6)

### 6.3 Stage 2 — Operations: Edit and Approve

*Matches the distributed `.xlsm` flow; no Vista connection required.*

**Grid behavior (Jobs-Ops tab)**
- ag-Grid displays all jobs for the batch; loads from LylesWIP only (no Vista call)
- **Yellow (editable) columns:** Ops Rev Override, Ops Cost Override, Ops Bonus/Profit, Completion Date
- **Non-yellow columns:** read-only (Vista-sourced values from the batch record)
- Inline editing directly in ag-Grid cells — no separate form
- **Mark Ops Done:** checkbox in Column H per row; on check, app calls `LylesWIPSaveJobRow` with `IsOpsDone=1`
  - Unchecking reverses (`IsOpsDone=0`)
- **Close job:** checkbox in Column G per row → `IsClosed=1` on save
- **Notes:** expandable notes cell per row (same semantic as Excel double-click Notes)
- Changes auto-save on cell blur (not on every keystroke); a dirty-row indicator shows unsaved rows

**Ops Final Approval**
- All rows must have `IsOpsDone=1` before the button is enabled (CompleteCheck gate)
- Ops user clicks **Ops Final Approval** → `LylesWIPUpdateBatchState` → `State = OpsApproved`
- Email notification sent to Accounting users for the batch's company

### 6.4 Stage 3 — Accounting Final Approval

*Matches Stage 3 in Excel.*

**Jobs-Ops tab** — fully read-only once `State = OpsApproved`

**Jobs-GAAP tab**
- Separate grid view showing GAAP columns
- **Yellow (editable) columns:** GAAP Rev Override, GAAP Cost Override (Accounting-only)
- **Mark GAAP Done:** checkbox in Column I per row → `LylesWIPSaveJobRow` with `IsGAAPDone=1`
- **Copy Ops to GAAP button:** calls `LylesWIP_CopyOpsToGAAP` stored proc scoped to current division; grid reloads

**Accounting Final Approval**
- All rows must have `IsGAAPDone=1` before the button is enabled
- Accounting user clicks **Accounting Final Approval** → `LylesWIPUpdateBatchState` → `State = AcctApproved`
- Batch is now **fully immutable** — no edits to any field, no state rollback (except Admin)
- If the WIP month is December: app calls `LylesWIPSaveYearEndSnapshot` after all divisions reach `AcctApproved`
- Email notification sent confirming final approval

### 6.5 Immutability and Guards

These rules are enforced server-side (API rejects the request) and reflected in the UI:

| Condition | Blocked action |
|-----------|---------------|
| `State = AcctApproved` | Any edit to any field; any state transition |
| `State < OpsApproved` | Ops Final Approval button |
| Not all rows `IsOpsDone=1` | Ops Final Approval button |
| `State < AcctApproved` | Accounting Final Approval button |
| Not all rows `IsGAAPDone=1` | Accounting Final Approval button |
| Ops user | GAAP tab is hidden; Create Batch is hidden |
| `State = ReadyForOps` or later | Ops columns are editable only by Ops role |
| `State = OpsApproved` or later | Ops columns become read-only for everyone |

### 6.6 Notifications

Triggered automatically on state transitions. All emails sent from the server; no manual step.

| Transition | Recipients | Subject template |
|------------|-----------|-----------------|
| `Open → ReadyForOps` | Ops users for company/division | "WIP Ready for Your Review — [Company] [Division] [Month]" |
| `OpsApproved` | Accounting users for company | "WIP Ops Approved — [Company] [Division] [Month]" |
| `AcctApproved` | Accounting users for company | "WIP Final Approval Complete — [Company] [Division] [Month]" |

Email body includes a direct link to the batch in the app. Recipient list is derived from
`WipUserPermissions` rows matching the batch's `JCCo` and `Department`.

> **Schema note:** An `Email` column (`nvarchar(254)`) must be added to `WipUserPermissions`
> if not already present. Migration script to be included in the project's `sql/` directory.

### 6.7 Excel Export

Available at any stage, any tab. An **Export to Excel** button above each grid generates a
server-side `.xlsx` file (via ClosedXML) and streams it to the browser as a download.

The export must:
- Match the column layout and labels of the Excel workbook's Jobs-Ops and Jobs-GAAP sheets
- Include company name, division, and WIP month in the header row
- Preserve number formatting (currency, percentages)
- Mark override cells with the same yellow fill (`FFFF00`) as the workbook

### 6.8 Jobs-Ops vs GAAP Comparison Tab

Read-only grid showing both Ops and GAAP values side by side for every job in the batch.
Available to Accounting users at any stage after batch creation. Not loadable from a
batch snapshot in Excel — the web app generates it live from `WipJobData` and the
Vista-sourced values stored in the batch record.

Columns: Job, Description, Contract Status, Ops Rev Override, GAAP Rev Override, Variance
(Rev), Ops Cost Override, GAAP Cost Override, Variance (Cost), Ops Bonus, GAAP Notes.

Export to Excel is available on this tab with the same server-side generation as §6.7.

### 6.9 Audit Trail

Every call to `LylesWIPSaveJobRow` stores `UpdatedBy` (Entra UPN) and `UpdatedAt` (UTC timestamp)
in `WipJobData`. The web app surfaces this as a **row history tooltip** on the Done checkboxes
(hover shows "Last saved by [name] on [date]"). This replicates the audit comment behavior from
the Excel workbook.

---

## 7. Data Flow (Web App)

```
Stage 1 (Accounting only):
  Browser → POST /api/batches          → LylesWIPCreateBatch (stored proc)
  Browser → GET  /api/batches/{id}/jobs → Vista SQL read + LylesWIPGetJobOverrides merge
  Browser → POST /api/batches/{id}/advance (ReadyForOps) → LylesWIPUpdateBatchState
                                                          → Email → Ops users

Stage 2 (Ops users):
  Browser → GET  /api/batches/{id}/jobs → LylesWIPGetJobOverrides (no Vista)
  Browser → PUT  /api/batches/{id}/jobs/{job} → LylesWIPSaveJobRow
  Browser → POST /api/batches/{id}/advance (OpsApproved) → LylesWIPUpdateBatchState
                                                          → Email → Accounting users

Stage 3 (Accounting):
  Browser → GET  /api/batches/{id}/jobs?view=gaap → LylesWIPGetJobOverrides (GAAP columns)
  Browser → PUT  /api/batches/{id}/jobs/{job}     → LylesWIPSaveJobRow (IsGAAPDone)
  Browser → POST /api/batches/{id}/copy-ops-to-gaap → LylesWIP_CopyOpsToGAAP
  Browser → POST /api/batches/{id}/advance (AcctApproved) → LylesWIPUpdateBatchState
                                                           → [if Dec] LylesWIPSaveYearEndSnapshot
                                                           → Email → Accounting users
```

---

## 8. UI / UX Requirements

- **ag-Grid** for all job data grids (existing license)
- Editable cells use ag-Grid's inline cell editing; yellow background (`#FFFF00`) on editable columns
- Read-only cells use the standard ag-Grid non-editable style (grey or white)
- Column layout mirrors the Excel workbook column order for Jobs-Ops and Jobs-GAAP
- Dirty rows (edited but not yet saved) show a left-border indicator
- State badge is always visible in the page header (e.g., `[Ready for Ops]`)
- Approval buttons are disabled with a tooltip explaining the blocking condition when guards are active
- No mobile breakpoints required; minimum supported viewport is 1280px wide

---

## 9. Out of Scope — Deferred to Future Phases

| Item | Reason deferred |
|------|----------------|
| JV's-Ops / JV's-GAAP sheets | Workflow not fully defined with Nicole |
| Real-time multi-user editing (WebSockets) | Not needed for Phase 1 |
| Admin UI for `WipUserPermissions` | Manage in SQL for now |
| Update Cost/Billed button | Under review — guard or remove |
| Dashboard / analytics views | Not in 5.73p |
| Viewpoint write-back (`LCGWIPMergeDetail`) | Out of scope per current architecture |

---

## 10. Resolved Decisions

| # | Question | Decision |
|---|----------|----------|
| OQ1 | SMTP relay for outbound email | `smtp.lylesgroup.com` |
| OQ2 | Email addresses for notification recipients | Add `Email` column to `WipUserPermissions` if not present |
| OQ3 | IIS app pool / SQL authentication | Connection strings with SQL auth (same accounts as today) |
| OQ4 | December year-end snapshot gate | Process runs monthly; December snapshot triggers when all divisions for a company reach `AcctApproved` in December — same gate, confirmed |
| OQ5 | Jobs-Ops vs GAAP comparison tab | **In scope for Phase 1** |

---

## 11. Success Criteria

Phase 1 is complete when:
1. All four companies (WML, AIC, APC, NESM) can complete a full three-stage WIP cycle in the web app
2. Data written by the web app matches the LylesWIP stored procedures' output (same validation as `validate_wip.py`)
3. Nicole and Cindy sign off that the ag-Grid view is an acceptable replacement for the Excel grid
4. Automatic email notifications fire correctly on each state transition
5. Excel export produces a file that Cindy can include in a board packet without modification
