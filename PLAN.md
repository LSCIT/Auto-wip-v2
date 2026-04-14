# Auto-WIP End-to-End Implementation Plan
*Updated: 2026-04-07 | Priority: HIGHEST — owner, board, senior management*
*Workflow source of truth: `LCG Automated WIP Process Steps Rev 7-15-2025.docx`*
*Product: Excel macro-enabled workbook. CEO committed to board and ownership.*

---

## ▶ RESUME HERE — April 8, 2026

**Current workbook versions:**
- `WIPSchedule -Rev 5.68p.xltm` — Nicole/Cindy demo version (frozen, don't touch)
- `WIPSchedule -Rev 5.70p.xltm` — superseded by 5.71p
- `WIPSchedule -Rev 5.71p.xltm` — Nicole feedback fixes + label changes (uploaded to SharePoint)

**VBA source:**
- `vba_source/` — frozen for 5.68p, do not edit
- `vba_source_569p/` — active development (5.70p and forward)

**Overall readiness:** ~90–93%. All Nicole feedback from 5.70p addressed. Waiting on re-test.
**Target delivery:** May 8, 2026 — on track.

### What's Ready (cumulative through Apr 7)
- Full Vista read path (all 4 sheets, all 22+ divisions)
- **JTD Cost matches Crystal Report exactly** (Mth fiscal month filtering)
- **Billed to Date matches Crystal Report exactly** (vrvJBProgressBills source)
- **Job inclusion: cost-exists as primary, StartMonth as fallback** (Nicole's rule from Apr 6 demo)
- **Prior month profit: prior month BonusProfit from LylesWIP** (via Z-column, formula-driven)
- **Prior year revenue: PYCost + bonus from LylesWIP year-end** (was missing bonus from Vista stub)
- **Prior year calc profit: sheet formula =AG-AH** (no longer VBA-overwritten)
- **Column labels: "MTD Change in Profit"** on Jobs-Ops AD and Jobs-GAAP AA
- LylesWIP database (10 stored procs, 4,973 override rows, 0 mismatches vs Nicole's files)
- Override merge on load with audit trail comments
- Write-back: Col H (Ops Done), Col I (GAAP Done), Col G (Close) all persist to LylesWIP
- Batch state machine: Open → ReadyForOps → OpsApproved → AcctApproved (full cycle tested)
- State machine guards, CompleteCheck gates, AcctApproved immutability
- Copy Ops to GAAP (stored proc + VBA, scoped to current division)
- Save & Distribute (.xlsm to C:\Trusted\ with ClearFormOnOpen=False)
- Audit trail comments on override cells ("Changed $X to $Y by user on date")
- December year-end snapshot (gated by AllBatchesApproved)
- Permissions_Modified.bas deployed (hardcoded WIPAccounting)
- **Vista write-back copy tables** (WipJCOP/WipJCOR + LylesWIPWriteBackToVista proc)
- **Push to Vista button** on Start sheet (test mode — all guards, no write)
- **Formula comparison verified**: 22/22 Jobs-Ops match, 21/22 Jobs-GAAP match (1 intentional Sprint 1 fix)

### Waiting On
- **Nicole/Cindy re-test on 5.71p** — email sent Apr 7 with status report, uploaded to SharePoint

### Questions for Nicole (sent Apr 7)
1. **Billed to Date source for GAAP** — Same JC Cost and Revenue Crystal Report, or different?
2. **Decimal display** — Should dollar columns show cents (#,##0.00) or whole dollars (#,##0)?

### Remaining Work

| # | Task | Priority | Owner | Status |
|---|------|----------|-------|--------|
| R1 | Nicole/Cindy re-test on 5.71p | Critical | Nicole + Cindy | Email sent Apr 7 — waiting |
| R2 | Verify residual rounding discrepancies | High | Josh | Spot-check after Nicole re-tests |
| R3 | Multi-company testing (AIC Co16, APC Co12, NESM Co13) | High | Josh | Not started — override data loaded |
| R4 | Circular reference on 51.1158 (pre-existing) | Low | Josh | Pre-existing in Michael's formulas — won't block delivery |
| R5 | Wire permissions to pnp.WIPSECGetRole | Medium | Josh | Not started |
| R6 | SQL driver install for Brian Platten + Harbir Atwal | Medium | Josh | Not started |
| R7 | User setup instructions | Medium | Josh | Not started |
| R8 | Security cleanup: re-hide Settings sheet, clear credentials | High | Josh | Not started — before any distribution |
| R9 | Final sign-off: Nicole + Cindy confirm numbers and complete 3-stage workflow | Critical | Nicole + Cindy | Blocked by R1 |
| R10 | Phase 2: enable Push to Vista button (uncomment write code) | Future | Josh | Infrastructure complete, button in test mode |

### Development Sequence

```
Tomorrow (Apr 8):
  1. Check for Nicole/Cindy feedback on 5.71p
  2. R3: Multi-company testing — open AIC/APC/NESM in 5.71p, compare against Crystal Reports
  3. Address any new findings from Nicole

This week (Apr 8–11):
  4. Iterate on Nicole/Cindy feedback
  5. R5: Wire permissions to P&P proc
  6. R8: Security cleanup

Next week (Apr 14–18):
  7. R6/R7: SQL driver installs + user setup docs
  8. R9: Final sign-off

Final stretch (Apr 21 – May 8):
  9. Address any remaining findings
  10. Production delivery
```

---

## Architecture — The Three-Tier Model

```
STAGE 1 — Accounting Initial Review (corporate — Vista access required)
  ├── Workbook reads live data from Vista Production (10.112.11.8)
  ├── Accounting reviews Jobs-Ops tab (read-only at this stage)
  ├── Batch created in LylesWIP on P&P → Vista data snapshot stored
  └── Start sheet → "Ready for Ops: Yes" → batch state advances

STAGE 2 — Operations Review (any location — P&P access only, no Vista required)
  ├── Workbook reads batch data from LylesWIP on P&P (no Vista connection needed)
  ├── Ops edits yellow override columns (revenue, cost, completion date, notes)
  ├── Double-click col H (Op's Done) → row saved to LylesWIP per job
  └── Start sheet → "Ops Final Approval: Yes" → batch state advances

STAGE 3 — Accounting Final Approval (corporate)
  ├── Workbook reads from LylesWIP — all Ops edits visible
  ├── Accounting edits GAAP override columns (Jobs-GAAP tab)
  ├── Double-click col I (GAAP Done) → row saved to LylesWIP per job
  ├── Start sheet → "Accounting Final Approval: Yes" → batch locked
  └── [Phase 2] Final values written back to Vista bJCOR / bJCOP
```

**Why P&P and not Vista direct:** Remote offices and job-site trailers cannot reach Vista's SQL Server, but P&P is accessible from all locations. Ops users never need a Vista connection — they operate entirely against LylesWIP on P&P.

---

## The 3-Stage User Workflow
*(Source: `LCG Automated WIP Process Steps Rev 7-15-2025.docx` — do not deviate from this)*

**Stage 1 — Accounting Initial Review (Nicole)**
1. Open workbook → Start tab → enter company (F4 button) + month (first of month format)
2. System checks for existing batch: if found → prompted to reopen; if not → division list opens
3. Ctrl/Shift-click to select division(s) → OK → batch created → Vista data loads and snapshots into LylesWIP
4. Select Jobs-Ops tab → **review only, no edits**
5. Start tab → **"Ready for Ops: Yes"** radio button → batch state advances to ReadyForOps
6. Close workbook

**Stage 2 — Ops Review & Final Approval (Operations)**
1. Open workbook → company/month (connects to LylesWIP on P&P — no Vista needed)
2. Prompted: open all divisions or select specific one
3. Jobs-Ops tab → **edit yellow override columns**
4. **Double-click col H (Op's Done)** per job = the save action (writes row to LylesWIP)
   - Double-click col H header = toggle all jobs Done/Not Done
5. **Double-click col G (Close)** if job ready to close
6. **Double-click Notes column** to expand/enter notes → double-click to collapse
7. All jobs marked Done → Start tab → **"Ops Final Approval: Yes"** → state advances to OpsApproved
8. Close workbook

**Stage 3 — Accounting Final Approval (Cindy/Nicole)**
1. Open workbook → company/month (reads from LylesWIP — sees all Ops edits)
2. Jobs-GAAP tab → **edit yellow override columns** (Accounting edits GAAP, not Ops)
3. **Double-click col I (GAAP Done)** per job = save action
4. Double-click Notes to enter notes
5. All jobs GAAP Done → Start tab → **"Accounting Final Approval: Yes"**
   - Batch state advances to AcctApproved → batch locked
6. Close workbook

**Batch state machine:**
```
Open → ReadyForOps → OpsApproved → AcctApproved
```

---

## Completed Sprints (Summary)

### Sprint 1 — Read Path ✅ (Mar 5–25)
All 4 sheets load from Vista Production. Formula bugs fixed (Col W, Col R). Live demo with Nicole + Cindy Mar 25.

### Sprint 2 — Vista Query Fixes ✅ (Mar 31, 11 days ahead)
All 8 tasks (A1–A8) complete. Date cutoff, billing fix, job inclusion logic, closed-month handling.

### Sprint 3 — Database + Override Load ✅ (Mar 31, 17 days ahead)
LylesWIP database live. 8 stored procs. 4,974 rows loaded from 40 historical files across 4 companies.

### Sprint 4 — Write Path + Override Merge ✅ (Apr 3, 21 days ahead)
SaveJobRow wired to Col H/I/G double-click. Override merge on load with dynamic row count. BatchValidate 24/24 PASS.

### Sprint 5 — Workflow + Distribution ✅ (Apr 3, 28 days ahead)
Full 3-stage state machine. CompleteCheck gates. AcctApproved immutability. Copy Ops to GAAP. Save & Distribute. Audit trail. Permissions module deployed (hardcoded). FormButtons.bas rewritten.

### Data Integrity Fix (Apr 6 AM)
- Fixed `to_decimal()` in load_overrides.py: Nicole's explicit $0 overrides now stored as 0 (not NULL)
- Reset test artifacts (Div 51 batch state, 28 workflow flags, 27 UserEdit rows)
- Restored 51.1108. override values to Nicole's originals
- Reloaded all 40 files (4,973 rows) with corrected zero handling
- Verified: 577 Dec 2025 rows, 0 mismatches against Nicole's source files

### Demo + Query Fixes — Rev 5.70p (Apr 6 PM)
- Demo with Nicole + Cindy on 5.68p — confirmed contract amounts, anticipated profit, workflow all correct
- **JTD Cost fix:** switched all bJCCD date filters from PostedDate/ActualDate to Mth (fiscal month). Exact match to Crystal Report: $87,315,159.22
- **Billed to Date fix:** switched from bARTH (AR side) to vrvJBProgressBills (JB side). Exact match: $96,918,206.90
- **Job inclusion fix:** added cost-exists fallback per Nicole's rule. Jobs with cost activity through batch month now included regardless of StartMonth. Adds 54.9416 + 57.0009
- **Prior month profit fix:** MergePriorMonthProfitsOntoSheet now computes recognized profit (earned rev - JTD cost) instead of projected profit (override rev - override cost)
- **Vista write-back tables:** WipJCOP + WipJCOR created in LylesWIP. LylesWIPWriteBackToVista proc with two-pass MERGE, quarterly + AcctApproved guards. Smoke-tested
- **Push to Vista button:** on Start sheet, test mode (all guards active, no write executed)
- **Formula comparison:** automated comparison vs Michael's 5.47p — 22/22 Jobs-Ops identical, 21/22 Jobs-GAAP (1 intentional Col Z fix from Sprint 1). Circular reference on 51.1158 confirmed pre-existing
- **`WIPexcel` account note:** may need SELECT on `vrvJBProgressBills` for production — discovered using jgarrison.sql account

### Nicole Feedback Fixes — Rev 5.71p (Apr 7)
- **Received Nicole's first feedback email** on 5.70p — 3 data issues on 51.1129, 2 label changes
- **Root cause (all 3 data issues):** VBA was overwriting visible formula columns instead of populating hidden Z-columns. Michael's template uses a two-layer architecture: Z-columns hold data from LylesWIP, visible columns have formulas reading from Z-columns. Our merge functions were writing directly to the visible columns, destroying the formulas.
- **COLJTDPriorProfit (AC) fix:** Rewrote `MergePriorMonthProfitsOntoSheet` to write prior month BonusProfit to Z-column BK (COLZPriorBonusProfit) instead of computing recognized profit and overwriting AC. Removed Vista JTD cost query (BuildPriorMonthCostLookup — now dead code).
- **COLAPYRev (AG) fix:** Expanded `MergePriorYearBonusOntoSheet` to also update AG = PYCost + bonus after writing bonus to AJ. Vista stubs PriorYrBonusProfit to 0; this backfills from LylesWIP year-end.
- **COLAPYCalcProfit (AI) fix:** Removed VBA overwrite in GetWipDetail2. Sheet formula `=AG-AH` auto-calculates once AG is correct.
- **Label changes:** Renamed COLJTDChgProfit header from "JTD Change In Profit" to "MTD Change in Profit" on Jobs-Ops (AD) and Jobs-GAAP (AA)
- **Sent email** to Nicole/Kevin/Dane/Cindy with status report + 5.71p on SharePoint + 2 outstanding questions (GAAP Billed source, decimal display)

---

## Milestone Schedule

| Milestone | Target | Status |
|-----------|--------|--------|
| Vista read path — all 4 sheets | Mar 28 | ✅ DONE |
| Vista query bugs fixed (A1–A8) | Apr 11 | ✅ DONE Mar 31 |
| LylesWIP database + stored procs | Apr 14 | ✅ DONE Mar 31 |
| Override load (40 files, 4,974 rows) | Apr 17 | ✅ DONE Mar 31 |
| Validation pipeline (24/24 PASS) | Apr 4 | ✅ DONE Apr 3 |
| Write path — SaveJobRow + Col H/I/G | Apr 11 | ✅ DONE Apr 3 |
| 3-stage workflow + state machine | Apr 14 | ✅ DONE Apr 3 |
| Audit trail + distribution | Apr 14 | ✅ DONE Apr 3 |
| Data integrity fix + full reload | Apr 6 | ✅ DONE Apr 6 |
| Nicole feedback — 5 items resolved | Apr 7–11 | ✅ DONE Apr 7 |
| Nicole/Cindy re-validation (5.71p) | Apr 7–11 | ⏳ Email sent, waiting |
| Multi-company testing | Apr 14–18 | 🔴 Not started |
| Permissions wire (P&P proc) | Apr 21–25 | 🔴 Not started |
| Security cleanup + final hardening | Apr 28 – May 2 | 🔴 Not started |
| **Production delivery** | **May 8** | **On track** |
| Phase 2: Vista bJCOP/bJCOR write-back | TBD post-delivery | |

---

## What We Are NOT Building Now

- **JV tabs** — defer until Jobs workflow is validated and signed off
- Trend/history columns (6-month, prior year snapshots) — defer
- Executive dashboards — defer
- Direct Nicole upload interface — defer
- Writing completion dates back to Vista custom field — defer (Phase 2)
- PM projection discrepancy ($2M+ PNP vs DJS gap) — defer; culture change driver
- Anything requiring a web browser — out of scope (CEO committed Excel to board)

---

## Critical Rules (Non-Negotiable)

**SQL performance:** Never use `LTRIM(RTRIM(job))` in JOIN/GROUP BY on bJCCD. Use raw field only — trim in final SELECT only. Proven 9min → 58sec on 8.8M rows.

**Save trigger:** Double-click on Done column = save. Not auto-save on every keystroke. `Sheet11.cls`/`Sheet12.cls` BeforeDoubleClick is the mechanism.

**OPS before GAAP:** Always. Ops edits Jobs-Ops yellow columns. GAAP tab is formula-driven from Ops. Only Accounting edits Jobs-GAAP yellow columns in Stage 3.

**Override priority:** Use override value from LylesWIP if present; otherwise use Vista-calculated value. Never discard a user override silently. Explicit $0 overrides are meaningful — stored as 0 with Plugged=1, not NULL.

**Do not touch Project Manager through VBA:** Vista requires system-level validation for PM changes.

**GAAP is quarterly only:** March, June, September, December. Non-GAAP months show blank GAAP values.

---

## Key Files

| File | Purpose |
|------|---------|
| `vm/WIPSchedule - Rev 5.68p.xltm` | **Current working workbook on VM** |
| `original/WIPSchedule - Rev 5.47p.xltm` | Original baseline for reference |
| `sql/WIP_Vista_Query.sql` | Main read query (7 CTEs) |
| `sql/LylesWIP_CreateDB.sql` | Full LylesWIP schema (tables + 9 stored procs) |
| `sql/load_overrides.py` | Loads Nicole's 40 Excel files → LylesWIP |
| `sql/LylesWIP_CopyOpsToGAAP.sql` | Stored proc: copy Ops overrides to GAAP columns |
| `vba_source/VistaData.bas` | Vista connection module |
| `vba_source/LylesWIPData.bas` | LylesWIP connection — merge + SaveJobRow + audit comments |
| `vba_source/BatchValidate.bas` | 24-combo unattended snapshot batch |
| `vba_source/Permissions_Modified.bas` | Deployed — eliminates Security Settings popup |
| `vba_source/FormButtons.bas` | Start sheet buttons — calls LylesWIP UpdateBatchState |
| `validate_wip.py` | Python validation framework |
| `patch_and_validate.py` | COM automation: import modules + run BatchValidate + validate |
| `test_workflow_e2e.py` | COM automation: E2E 3-stage workflow test |
| `vba_source/COLUMN_MAPPING.md` | Column letter ↔ named range ↔ Vista field |
| `LCG Automated WIP Process Steps Rev 7-15-2025.docx` | **Workflow authority** |
| `Data-from-Nicole/` | 40 historical override Excel files (Dec 2024–Dec 2025, 4 companies) |
| `Status-Reports/` | Weekly status reports (latest: Apr 3, 2026) |
