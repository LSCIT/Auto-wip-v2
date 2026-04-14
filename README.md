# Auto-WIP

Web application for Lyles Services Co.'s monthly Work in Progress (WIP) schedule process.
Replaces an Excel/VBA tool (Rev 5.73p) with a browser-based three-stage workflow.

## Stack

- **Frontend:** React (Vite) + ag-Grid
- **Backend:** ASP.NET Core 8 Web API
- **Auth:** Microsoft Entra ID
- **Host:** IIS on `cloud-apps1.ad.lylesgroup.com`
- **DB:** LylesWIP (SQL Server, read/write) + Vista/Viewpoint (read-only, Stage 1 only)

## Documentation

| Doc | Description |
|-----|-------------|
| [PRD.md](PRD.md) | Product requirements — scope, features, decisions |
| [ARCHITECTURE.md](ARCHITECTURE.md) | Stack, API design, auth, DB schema |
| [WORKFLOW.md](WORKFLOW.md) | Three-stage workflow diagrams (Mermaid) |
| [VALIDATION.md](VALIDATION.md) | Data validation rules and source-of-truth hierarchy |

## Workflow Summary

1. **Stage 1 — Accounting:** Creates a batch, loads Vista data, reviews, advances to Ready for Ops
2. **Stage 2 — Operations:** PMs edit Ops override columns, mark jobs Done, give Ops Final Approval
3. **Stage 3 — Accounting:** Reviews GAAP columns, marks jobs Done, gives Accounting Final Approval → batch locked

See [WORKFLOW.md](WORKFLOW.md) for state machine and full flowchart.

## Setup

Copy `.env.example` to `.env` and fill in DB credentials before running locally.

## Repository

`github.com/LSCIT/auto-wip-v2` — active branch: `claude/convert-to-web-app-qVYTK`
