# Auto-WIP Web Application — Claude Code Context
*Last updated: April 2026*

## Before starting any work, read PRD.md, ARCHITECTURE.md, and WORKFLOW.md.

---

## What This Is

A web application replacing the Lyles Services Co. Automated WIP (Work in Progress) Schedule
Excel/VBA tool (last Excel version: Rev 5.73p, currently in stakeholder review).

**Stack:** ASP.NET Core 8 Web API + React (Vite) + ag-Grid
**Auth:** Microsoft Entra ID (`Microsoft.Identity.Web`)
**Host:** IIS on `cloud-apps1.ad.lylesgroup.com`
**Databases:** LylesWIP (localhost, read/write) + Vista (10.112.11.8, read-only, Stage 1 only)

---

## Stakeholders

| Person | Role |
|--------|------|
| Josh Garrison | Director of Technology Innovation (developer — you work for him) |
| Kevin Shigematsu | CEO / project sponsor |
| Cindy Jordan | CFO (final sign-off) |
| Nicole Leasure | VP Corporate Controller (primary user and validator) |
| Dane Wildey | CIO |
| Brian Platten / Harbir Atwal | Controllers (reviewers) |

---

## DB Connections

Credentials are stored in `.env` (not committed). See `.env.example` for variable names.

- **LylesWIP:** `cloud-apps1.ad.lylesgroup.com`, SQL auth — `LYLESWIP_UID` / `LYLESWIP_PWD`
- **Vista:** `10.112.11.8`, SQL auth — `VISTA_UID` / `VISTA_PWD`

---

## Critical Rules

- **SQL performance:** Never `LTRIM(RTRIM(job))` in JOIN or GROUP BY on `bJCCD`. Raw field only
  in joins; trim only in SELECT. Violating this = 9-minute queries on 8.8M rows.
- **Override priority:** LylesWIP non-NULL override wins over Vista-calculated value. Never
  silently discard an override.
- **GAAP is quarterly only:** March, June, September, December. Non-GAAP months = blank GAAP values.
- **Stage isolation:** Operations users have no Vista access. Stage 2 loads from LylesWIP only.
- **Stored procs are frozen:** Do not modify the LylesWIP stored procedures. The web app calls
  them exactly as the Excel VBA did.
- **Authorization is server-side only:** All `WipPrincipal` checks happen in the API.
  Never trust role/division values from the client.
- **Sub-job normalization:** Append `.` to job number before LylesWIP lookup if last char is not `.`
  (affects `56.1009.01` and `56.1010.01` only — both WML Div56).

---

## Key Files

| File | Purpose |
|------|---------|
| `PRD.md` | Full product requirements (Phase 1) |
| `ARCHITECTURE.md` | Stack, API design, DB schema, auth flow |
| `WORKFLOW.md` | Mermaid workflow diagrams |
| `VALIDATION.md` | Data validation rules — Nicole's source-of-truth hierarchy |
| `sql/LylesWIP_CreateDB.sql` | LylesWIP schema + all stored procs |
| `sql/WIP_Vista_Query.sql` | 7-CTE Vista read query (port to .NET) |
| `sql/LylesWIP_CopyOpsToGAAP.sql` | Copy Ops → GAAP stored proc |
| `sql/reference/` | Michael's original procs (reference only — do not call) |
| `docs/bJCOP_bJCOR_Investigation.md` | Vista table investigation notes |
| `LCG Automated WIP Process Steps Rev 7-15-2025.docx` | Business process doc |
| `.env.example` | Environment variable template |

---

## Excel Reference (Rev 5.73p)

The working Excel version is in `vm/` (not committed — gitignored). It is the source of truth
for exact column layout, override behavior, state machine guards, and formula logic that the
web app must replicate in Phase 1.

The VBA source has been removed from the repo — it is no longer the active development target.
The SQL stored procedures it called are preserved in `sql/`.
