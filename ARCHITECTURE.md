# Auto-WIP Web Application — Architecture
*Last updated: April 2026*

---

## Overview

Auto-WIP is a three-stage monthly workflow application that lets Lyles Services Co. manage
Work in Progress (WIP) schedules across four companies and ~25 divisions. It replaces an
Excel/VBA tool (Rev 5.73p) with a browser-based application that connects to the same
SQL Server databases.

**See [WORKFLOW.md](WORKFLOW.md) for workflow diagrams.**
**See [PRD.md](PRD.md) for full product requirements.**

---

## Deployment Topology

```
Corporate Laptop (browser, 1280px+ viewport)
    │
    │  HTTPS
    ▼
cloud-apps1.ad.lylesgroup.com  (IIS)
    ├── React SPA            ← static files served by IIS
    └── ASP.NET Core 8 API
            │
            ├── LylesWIP DB (localhost — same host, low latency)
            │     SQL auth via connection string
            │
            └── Vista DB (10.112.11.8 — read-only)
                  SQL auth via connection string
                  ONLY called during Stage 1 batch creation
```

---

## Technology Stack

| Layer | Choice | Notes |
|-------|--------|-------|
| Frontend | React (Vite) + ag-Grid | Existing ag-Grid license |
| Backend | ASP.NET Core 8 Web API | Hosted on IIS |
| Auth | Microsoft Entra ID (`Microsoft.Identity.Web`) | MSAL, JWT validation |
| Data access | `Microsoft.Data.SqlClient` — direct ADO.NET | Calls existing stored procs |
| Excel export | ClosedXML | Server-side `.xlsx` generation |
| Email | SMTP — `smtp.lylesgroup.com` | State-transition notifications |
| Vista read | Direct SQL (`10.112.11.8`) | Stage 1 only; read-only |
| LylesWIP | Direct SQL (localhost) | All stages; read + write |

---

## Authentication and Authorization

**Authentication:** Microsoft Entra ID
- User visits the app → redirect to Entra login if no valid token
- Entra returns a signed JWT
- `Microsoft.Identity.Web` validates the JWT on every API request
- `User.Identity.Name` = Entra UPN (e.g. `jgarrison@lylesgroup.com`)

**Authorization:** `WipUserPermissions` table in LylesWIP
- The API looks up the authenticated UPN in `WipUserPermissions`
- Attaches a `WipPrincipal` object to the request context: `{Role, Companies[], Divisions[]}`
- All permission checks are server-side against `WipPrincipal` — never trust client-supplied values

**Roles:**

| Role | Capabilities |
|------|-------------|
| `Accounting` | Create batches, Stage 1 review, Stage 3 GAAP edits and final approval, Excel export, comparison tab |
| `Operations` | Stage 2 only — edit Ops columns, mark Ops Done, Ops Final Approval |
| `Admin` | All of the above + reset any batch to Open |

**Schema note:** `WipUserPermissions` requires an `Email nvarchar(254)` column for notification
routing. Migration script: `sql/Add_Email_To_WipUserPermissions.sql` (to be created).

---

## Database — LylesWIP

Lives on `cloud-apps1.ad.lylesgroup.com` (same host as the web server).

### Key Tables

| Table | Purpose |
|-------|---------|
| `WipBatches` | One row per Company+Month+Division batch; tracks `State` |
| `WipJobData` | Override values per job per batch; IsOpsDone, IsGAAPDone, IsClosed |
| `WipUserPermissions` | Maps Entra UPN → Role + Company + Division |
| `WipYearEndSnapshot` | Prior-year revenue/cost baseline (populated each December) |

### Stored Procedures (do not modify)

| Procedure | Called by | Purpose |
|-----------|-----------|---------|
| `LylesWIPCreateBatch` | Stage 1 | Creates WipBatches row, State=Open |
| `LylesWIPUpdateBatchState` | Stage 1,2,3 | Advances or resets batch state |
| `LylesWIPGetJobOverrides` | All stages | Returns override rows for co+month |
| `LylesWIPSaveJobRow` | Stage 2,3 | Upserts WipJobData; records UpdatedBy+UpdatedAt |
| `LylesWIP_CopyOpsToGAAP` | Stage 3 | Copies Ops overrides to GAAP columns for a division |
| `LylesWIPSaveYearEndSnapshot` | Stage 3 (Dec) | Saves prior-year baseline |
| `LylesWIPCheckBatchState` | All | Returns current batch state |

Full schema: `sql/LylesWIP_CreateDB.sql`

---

## Database — Vista / Viewpoint

Lives on `10.112.11.8`. **Read-only.** Only accessed during Stage 1 batch creation.
Operations users never trigger a Vista query.

### Key Query

`sql/WIP_Vista_Query.sql` — 7-CTE query that reads from:
- `bJCCD` (job cost detail — 8.8M rows)
- `vrvJBProgressBills` (billing)
- `bJCCM` (contract master — source for `Department` filter)
- `bJCJM` (job master)
- `bHQCO` (company list)
- `bJCDM` (division/department list)

**Critical performance rule:** Use raw `Job` field in all JOINs and GROUP BY on `bJCCD`.
LTRIM/RTRIM in joins causes a 9-minute query. Without trimming in joins: 58 seconds.

---

## API Design (high-level)

```
GET    /api/companies                          → company list (scoped to WipPrincipal)
GET    /api/companies/{co}/divisions           → division list for company
GET    /api/batches                            → batches visible to current user
POST   /api/batches                            → create batch (Accounting only)
GET    /api/batches/{id}                       → batch header + state
GET    /api/batches/{id}/jobs                  → job rows (Ops view)
GET    /api/batches/{id}/jobs?view=gaap        → job rows (GAAP view)
GET    /api/batches/{id}/jobs?view=comparison  → Ops vs GAAP side-by-side
PUT    /api/batches/{id}/jobs/{job}            → save job row (Ops or GAAP override)
POST   /api/batches/{id}/advance               → advance state (body: {targetState})
POST   /api/batches/{id}/copy-ops-to-gaap      → copy Ops overrides → GAAP for division
GET    /api/batches/{id}/export?view=ops|gaap  → download .xlsx
```

All mutating endpoints enforce `WipPrincipal` authorization and return `403` if the user
lacks permission for the batch's company+division.

---

## Companies and Divisions

| Company | JCCo | Divisions | Notes |
|---------|------|-----------|-------|
| W. M. Lyles Co. (WML) | 15 | 50–58 | Div50 = overhead |
| Advanced Integration & Controls (AIC) | 16 | 70–78 | Div70 = overhead |
| American Paving Co. (APC) | 12 | 20, 21 | |
| New England Sheet Metal (NESM) | 13 | 31, 32, 33, 35 | No Div34 |

**Division filter:** Always use `bJCCM.Department` from Vista, not job number prefix.
APC jobs `20.252x–20.257x` are in Dept=21 despite the `20.` prefix.

---

## Override Merge Logic

When jobs are loaded for any grid view:

1. Vista baseline values are the starting point (loaded at batch creation, cached in batch record)
2. `LylesWIPGetJobOverrides` returns all `WipJobData` rows for co+month
3. For each job: if a non-NULL override exists in `WipJobData`, it takes priority
4. NULL = no override → show Vista-calculated value

**Override columns:** OpsRevOverride, OpsCostOverride, GAAPRevOverride, GAAPCostOverride,
BonusProfit, CompletionDate.

---

## Sub-job Normalization

Vista sub-jobs (`56.1009.01`, `56.1010.01`) have no trailing dot in `bJCJM`.
LylesWIP stores them with a trailing dot (`56.1010.01.`).
The API must normalize before lookup: append `.` if the last character is not `.`.
Only two sub-jobs exist across all four companies, both in WML Div56.

---

## Email Notifications

Sent via `smtp.lylesgroup.com` on every batch state transition.

| Transition | To | Subject |
|------------|----|---------|
| Open → ReadyForOps | Ops users for company+division | `WIP Ready for Your Review — {Company} {Division} {Month}` |
| ReadyForOps → OpsApproved | Accounting users for company | `WIP Ops Approved — {Company} {Division} {Month}` |
| OpsApproved → AcctApproved | Accounting users for company | `WIP Final Approval Complete — {Company} {Division} {Month}` |

Recipient list sourced from `WipUserPermissions.Email` for matching `JCCo` + `Department`.
Email body includes a direct deep-link to the batch in the app.

---

## Key Files

| File/Dir | Purpose |
|----------|---------|
| `PRD.md` | Full product requirements |
| `WORKFLOW.md` | Workflow diagrams (Mermaid) |
| `VALIDATION.md` | Data validation rules and source-of-truth hierarchy |
| `sql/LylesWIP_CreateDB.sql` | Full LylesWIP schema — tables + all stored procs |
| `sql/WIP_Vista_Query.sql` | 7-CTE Vista read query (reference for .NET port) |
| `sql/LylesWIP_CopyOpsToGAAP.sql` | Copy Ops to GAAP stored proc |
| `sql/LylesWIP_VistaWriteBack.sql` | Vista write-back tables (deferred — not Phase 1) |
| `sql/reference/` | Michael's original WipDb stored procs (reference only) |
| `LCG Automated WIP Process Steps Rev 7-15-2025.docx` | Business process document |
| `docs/bJCOP_bJCOR_Investigation.md` | Vista bJCOP/bJCOR table investigation notes |
