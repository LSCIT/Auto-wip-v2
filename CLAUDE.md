# Auto-WIP Project — Claude Code Context
*Last updated: 2026-04-03*

## Before starting any work, read ARCHITECTURE.md, MODULES.md, and VALIDATION.md.

## Development Environment (as of Apr 3, 2026)
- **Machine:** Windows VM (not Mac) — project at `E:\Auto-Wip\`
- **Python:** 3.12 — `python` command (not `python3`)
- **pywin32 installed** — use `win32com.client` for Excel COM automation (import .bas files, run macros)
- **BatchValidate output:** `C:\Trusted\validate-d3\` — set `WIP_SNAPSHOT_DIR=C:\Trusted\validate-d3` when running `validate_wip.py` to read from there directly
- **Automation script:** `patch_and_validate.py` — imports modules + runs BatchValidateAll + copies snapshots + validates
- **DB connections confirmed working on this machine:**
  - LylesWIP: `10.103.30.11`, ODBC Driver 18, UID=`wip.excel.sql`
  - Vista: `10.112.11.8`, ODBC Driver 18, UID=`WIPexcel`

---

## What This Is
Rebuilding the Automated WIP (Work In Progress) Schedule for Lyles Services Co. from a broken
Excel/VBA tool. **The product is Excel.** CEO committed this to the board. No web app, no portal.

## Stakeholders
- **Josh Garrison** — Director of Technology Innovation (developer, you work for him)
- **Kevin Shigematsu** — CEO / project sponsor
- **Cindy Jordan** — CFO (final sign-off)
- **Nicole Leasure** — VP Corporate Controller (primary user and validator)
- **Dane Wildey** — CIO
- **Brian Platten / Harbir Atwal** — Controllers (reviewers; HA initials = Harbir Atwal)
- Michael Roberts — original consultant developer (replaced; still available but not involved)

## Project State (Apr 3, 2026) — ~85–90% complete
| Area | Status |
|------|--------|
| Vista read path (all 4 sheets) | ✅ Complete — deployed Rev 5.73p |
| LylesWIP database + stored procs | ✅ Complete — 9 procs on 10.103.30.11 (added CopyOpsToGAAP) |
| Override load (40 historical files, 4,974 rows) | ✅ Complete |
| Override merge on load (LylesWIPData.bas) | ✅ Complete — with audit trail comments |
| Validation pipeline (BatchValidate + validate_wip.py) | ✅ 24/24 PASS, 0 mismatches |
| Permissions_Modified.bas | ✅ Deployed to workbook |
| Write-back: Col H (Ops Done) → SaveJobRow | ✅ Working — manually tested |
| Write-back: Col I (GAAP Done) → SaveJobRow | ✅ Working — manually tested |
| Write-back: Col G (Close) → IsClosed | ✅ Working — manually tested |
| Batch state machine (Open→RFO→OpsApproved→AcctApproved) | ✅ Working — full cycle tested |
| State machine guards (No buttons blocked after approval) | ✅ Working |
| CompleteCheck gates (Ops Done / GAAP Done before approval) | ✅ Working |
| AcctApproved immutability (edits blocked, no regression) | ✅ Working |
| Copy Ops to GAAP (scoped to current division) | ✅ Working — stored proc + VBA |
| Save & Distribute (.xlsm to C:\Trusted\) | ✅ Working |
| Audit trail comments on override cells | ✅ Working — "Changed $X to $Y by user on date" |
| December year-end snapshot (gated by all divisions) | ✅ Built — AllBatchesApproved check |
| Clear batch prompt removed from BeforeClose | ✅ Done |
| Distribution workflow (.xlsm to PMs) | ✅ Working |
| Nicole/Cindy review session | ⏳ Ready to schedule |

## Resume Point (next session)
1. Schedule Nicole/Cindy review demo — full 3-stage workflow is working
2. Confirm data questions: `54.9416.` StartMonth, NESM Div35 zero-activity
3. Test with other companies (AIC Co16, APC Co12, NESM Co13) — override data already loaded
4. Remaining polish: Update Cost/Billed button (guard or remove), JV sheets (deferred)
5. User setup: SQL driver for Brian Platten / Harbir Atwal (D4/D5)
6. Security cleanup: re-hide Settings sheet, clear test credentials (D7)

## Critical Rules (non-negotiable)
- **SQL**: Never `LTRIM(RTRIM(job))` in JOIN/GROUP BY on bJCCD. Raw field only in joins; trim in SELECT.
  Violating this: 9-minute queries. Correct: 58 seconds. 8.8M rows.
- **Save trigger**: Double-click col H (Ops Done) = save. Double-click col I (GAAP Done) = save.
  NOT auto-save on cell change.
- **OPS before GAAP**: Ops edits Jobs-Ops. GAAP is formula-driven from Ops. Only Accounting edits GAAP yellow columns in Stage 3.
- **Override priority**: Use LylesWIP override if present; otherwise show Vista-calculated value. Never silently discard an override.
- **GAAP is quarterly only**: March, June, September, December. Non-GAAP months = blank GAAP values.
- **Do not edit Project Manager** through VBA — Vista requires system-level validation for PM changes.

## Key Files
| File | Purpose |
|------|---------|
| `PLAN.md` | Sprint plan, milestone schedule, detailed resume point |
| `ARCHITECTURE.md` | Data flow, state machine, division structure |
| `MODULES.md` | Every VBA module — status, inputs/outputs, known issues |
| `VALIDATION.md` | Test date, validation results, Nicole's source-of-truth rules |
| `SCHEMA.md` | Vista tables, LylesWIP schema, stored procs |
| `vm/WIPSchedule -Rev 5.73p.xltm` | **Active workbook on VM** |
| `vba_source/LylesWIPData.bas` | LylesWIP connection + merge + SaveJobRow + audit comments |
| `vba_source/BatchValidate.bas` | 24-combo batch runner |
| `patch_and_validate.py` | COM automation: import modules + run BatchValidate + validate |
| `test_workflow_e2e.py` | COM automation: E2E 3-stage workflow test |
| `sql/LylesWIP_CopyOpsToGAAP.sql` | Stored proc: copy Ops overrides to GAAP columns |
| `validate_wip.py` | Python validation tool |
| `sql/WIP_Vista_Query.sql` | 7-CTE Vista read query |
| `sql/LylesWIP_CreateDB.sql` | Full LylesWIP schema (tables + 8 stored procs) |
| `sql/load_overrides.py` | Loads Nicole's 40 Excel files → LylesWIP |
| `Data-from-Nicole/` | 40 historical override Excel files (4 companies, Dec 2024–Dec 2025) |

## Conventions
- LOCAL repo only — no remote push. CONFIDENTIAL (DB credentials, server IPs).
- VBA target: MSOLEDBSQL provider (not ODBC). Module-level ADODB.Connection, reuse if open.
- Files ending `_Modified.bas` = deployed version replacing original. Originals kept for diff.
