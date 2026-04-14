"""
daily_report.py — Auto-WIP Daily Status Report Generator
Produces: vm/status_reports/YYYY-MM-DD_status.pdf and .md
"""

import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

TODAY       = date.today().isoformat()
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR  = os.path.join(BASE_DIR, 'vm', 'status_reports')
PDF_FILE    = os.path.join(REPORT_DIR, f'{TODAY}_status.pdf')
MD_FILE     = os.path.join(REPORT_DIR, f'{TODAY}_status.md')

os.makedirs(REPORT_DIR, exist_ok=True)

# =============================================================================
# Styles
# =============================================================================
BASE_STYLES = getSampleStyleSheet()

def make_style(name, parent='Normal', **kwargs):
    return ParagraphStyle(name=name, parent=BASE_STYLES[parent], **kwargs)

S_TITLE    = make_style('ReportTitle', 'Title', fontSize=20, textColor=colors.HexColor('#1F3864'),
                        spaceAfter=4, alignment=TA_CENTER)
S_SUBTITLE = make_style('ReportSubtitle', fontSize=11, textColor=colors.HexColor('#555555'),
                        spaceAfter=2, alignment=TA_CENTER)
S_H1       = make_style('H1', fontSize=13, textColor=colors.HexColor('#1F3864'),
                        spaceBefore=14, spaceAfter=4, fontName='Helvetica-Bold')
S_H2       = make_style('H2', fontSize=11, textColor=colors.HexColor('#2E5090'),
                        spaceBefore=10, spaceAfter=3, fontName='Helvetica-Bold')
S_BODY     = make_style('Body', fontSize=9, spaceAfter=4, leading=13)
S_BULLET   = make_style('Bullet', fontSize=9, spaceAfter=3, leading=13,
                        leftIndent=14, bulletIndent=4)
S_SMALL    = make_style('Small', fontSize=8, textColor=colors.HexColor('#666666'), leading=11)
S_CODE     = make_style('Code', fontSize=8, fontName='Courier',
                        backColor=colors.HexColor('#F5F5F5'), leading=11)
S_WARN     = make_style('Warn', fontSize=9, textColor=colors.HexColor('#9C0006'),
                        leading=13, spaceAfter=3)
S_OK       = make_style('OK', fontSize=9, textColor=colors.HexColor('#276221'),
                        leading=13, spaceAfter=3)

def h1(text):      return Paragraph(text, S_H1)
def h2(text):      return Paragraph(text, S_H2)
def body(text):    return Paragraph(text, S_BODY)
def bullet(text):  return Paragraph(f'• {text}', S_BULLET)
def small(text):   return Paragraph(text, S_SMALL)
def warn(text):    return Paragraph(f'⚠ {text}', S_WARN)
def ok(text):      return Paragraph(f'✓ {text}', S_OK)
def sp(h=6):       return Spacer(1, h)
def rule():        return HRFlowable(width='100%', thickness=0.5,
                                     color=colors.HexColor('#CCCCCC'), spaceAfter=6)

def tbl(data, col_widths, header_rows=1):
    t = Table(data, colWidths=col_widths, repeatRows=header_rows)
    style_cmds = [
        ('BACKGROUND',   (0,0), (-1,0),           colors.HexColor('#1F3864')),
        ('TEXTCOLOR',    (0,0), (-1,0),           colors.white),
        ('FONTNAME',     (0,0), (-1,0),           'Helvetica-Bold'),
        ('FONTSIZE',     (0,0), (-1,0),           9),
        ('ALIGN',        (0,0), (-1,0),           'CENTER'),
        ('VALIGN',       (0,0), (-1,-1),          'TOP'),
        ('FONTSIZE',     (0,1), (-1,-1),          8),
        ('FONTNAME',     (0,1), (-1,-1),          'Helvetica'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),          [colors.white, colors.HexColor('#F2F6FC')]),
        ('GRID',         (0,0), (-1,-1),          0.4, colors.HexColor('#CCCCCC')),
        ('LEFTPADDING',  (0,0), (-1,-1),          6),
        ('RIGHTPADDING', (0,0), (-1,-1),          6),
        ('TOPPADDING',   (0,0), (-1,-1),          4),
        ('BOTTOMPADDING',(0,0), (-1,-1),          4),
        ('WORDWRAP',     (0,0), (-1,-1),          True),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t

def p(text, style=None):
    return Paragraph(text, style or S_BODY)

W = letter[0] - 2*inch   # usable width

# =============================================================================
# Report content
# =============================================================================
def build_story():
    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    story += [
        Paragraph('Lyles Services Co.', S_SUBTITLE),
        Paragraph('Automated WIP Schedule — Daily Status Report', S_TITLE),
        Paragraph(f'Report Date: {TODAY} &nbsp;|&nbsp; Prepared by: Josh Garrison / Claude Code',
                  S_SUBTITLE),
        sp(4), rule(), sp(4),
    ]

    # ── Executive Summary ─────────────────────────────────────────────────────
    story += [
        h1('Executive Summary'),
        body(
            'The Automated WIP Schedule rebuild is in its validation phase for the December 2025 '
            'period. Two significant bugs were identified and resolved today: a job number format '
            'mismatch in the LylesWIP override database and a critical VBA named-range bug that '
            'caused override data to silently drop for any division with more than 18 active jobs. '
            'The fix has been deployed and a fresh batch of 22 division snapshots is being '
            'regenerated now on the VM for re-validation. The read path is fully functional. '
            'Write-back to LylesWIP is the primary remaining work.'
        ),
        sp(4),
    ]

    # ── Project Snapshot ──────────────────────────────────────────────────────
    story += [h1('Project Snapshot'), sp(2)]

    snap_data = [
        [p('Item', S_SMALL), p('Detail', S_SMALL)],
        [p('Project'),       p('Automated WIP Schedule — full end-to-end rebuild')],
        [p('Sponsor'),       p('Kevin Shigematsu (CEO)')],
        [p('Primary User'),  p('Nicole Leasure (VP Corporate Controller)')],
        [p('Final Validator'),p('Cindy Jordan (CFO)')],
        [p('Developer'),     p('Josh Garrison (Director of Technology Innovation)')],
        [p('Priority'),      p('HIGHEST — escalated by owner, board, and senior management (March 2026)')],
        [p('Target Month'),  p('December 2025 (validation) → current month on go-live')],
        [p('Architecture'),  p('Excel/.xltm workbook + VBA → Vista/Viewpoint (read) + LylesWIP on PNP server (write)')],
        [p('Current Rev'),   p('WIPSchedule Rev 5.68p (deployed to VM)')],
    ]
    story += [tbl(snap_data, [1.8*inch, W - 1.8*inch]), sp(8)]

    # ── Today's Work ──────────────────────────────────────────────────────────
    story += [h1("Today's Accomplishments  (2026-04-02)"), sp(2)]

    # Bug 1
    story += [
        h2('1 — Job Number Format Fix (load_overrides.py)'),
        body(
            'Nicole\'s historical WIP import files stored 3-digit job number segments '
            '(e.g. <font name="Courier">73.105.</font>) while Vista normalizes to 4-digit '
            '(<font name="Courier">73.1050.</font>). This caused LylesWIP to hold unmatched '
            'keys for all AIC (Co16) and APC (Co12) overrides.'
        ),
        bullet('Deleted 132 short-format rows from dbo.WipJobData for Co12 and Co16'),
        bullet('Updated normalize_job() in sql/load_overrides.py to right-pad with ljust(4, "0")'),
        bullet('Re-ran load_overrides.py — 4,973 rows upserted across all 40 historical files'),
        bullet('Verified: job 78.12. → 78.1200. correctly aligns with new Vista sequential jobs'),
        ok('LylesWIP now contains correctly keyed overrides for all four companies'),
        sp(4),
    ]

    # Bug 2
    story += [
        h2('2 — Root Cause: SummaryData Static Named Range (LylesWIPData.bas)'),
        body(
            'Validation showed 344 mismatches across 11 divisions: the workbook Z columns '
            '(COLZOPsRev, COLZOPsCost, COLZGAAPRev, COLZGAAPCost) were showing 0 even though '
            'LylesWIP had the correct data with Plugged=True. Root cause identified:'
        ),
        bullet(
            'The SummaryData named range is initialized to only 18 rows ($A$7:$CM$24) in the template.'
        ),
        bullet(
            'GetWipDetail2 writes jobs using SummaryData.Cells(r, …) — VBA silently extends '
            'beyond the named range. For Div51 (28 jobs), jobs 19–28 land in sheet rows 25–34.'
        ),
        bullet(
            'MergeOverridesOntoSheet, MergePriorMonthProfitsOntoSheet, and '
            'MergePriorYearBonusOntoSheet all loop "For r = 1 To summaryRange.Rows.Count" = 18. '
            'Jobs beyond row 24 are never visited — overrides silently dropped.'
        ),
        warn(
            'Any division with more than 18 jobs lost all overrides beyond the 18th job. '
            'This affected WML Div51 (28 jobs, 12 missed), AIC Div73 (83 jobs, 67 missed), and others.'
        ),
        sp(4),
    ]

    # Fix
    story += [
        h2('3 — Fix Applied (LylesWIPData.bas)'),
        body('All three Merge* functions now dynamically detect the actual last data row:'),
        p(
            'lastDataRow = sh.Cells(sh.Rows.Count, jnColAbs).End(xlUp).Row\n'
            'totalRows = Application.Max(summaryRange.Rows.Count, lastDataRow − summaryRange.Row + 1)\n'
            'For r = 1 To totalRows',
            S_CODE
        ),
        bullet('Uses xlUp scan on the job number column to find actual last written row'),
        bullet('Application.Max() ensures the fix is additive — never shrinks below existing Rows.Count'),
        bullet('Applied identically to all three functions: no risk of partial fix'),
        ok('LylesWIPData.bas deployed to workbook Rev 5.68p on VM'),
        sp(4),
    ]

    # BatchValidate
    story += [
        h2('4 — BatchValidate.bas — Automated Snapshot Generation'),
        body(
            'Created BatchValidate.bas to automate generation of all division snapshot files '
            'instead of 20–30 minutes of manual clicking. The macro loops all 22 company/division '
            'combinations, loads data, and saves copies to C:\\Trusted\\validate-d3\\ on the VM.'
        ),
    ]

    batch_data = [
        [p('Company', S_SMALL), p('Divisions', S_SMALL), p('Count', S_SMALL)],
        [p('WML (Co15)'),  p('51, 52, 53, 54, 55, 56, 57, 58'), p('8')],
        [p('AIC (Co16)'),  p('70, 71, 72, 73, 74, 75, 76, 77, 78'), p('9')],
        [p('APC (Co12)'),  p('21'), p('1')],
        [p('NESM (Co13)'), p('31, 32, 33, 35'), p('4')],
        [p('TOTAL', S_SMALL), p(''), p('22')],
    ]
    story += [tbl(batch_data, [1.5*inch, W - 2.5*inch, 1*inch]), sp(4)]
    story += [
        bullet('Uses ResetWorkbook (not ClearForms3) to properly unprotect between iterations'),
        bullet('Non-fatal error handling — logs to Immediate window, continues to next division'),
        bullet('OUTPUT_PATH: C:\\Trusted\\validate-d3\\  (user-specified)'),
        ok('Macro is running now — 22 snapshots being generated with fixed LylesWIPData.bas'),
        sp(4),
    ]

    # validate_wip.py
    story += [
        h2('5 — validate_wip.py — Python Validation Script'),
        body(
            'Created validate_wip.py to systematically compare Nicole\'s Dec 2025 WIP History '
            'Import files against the 22 workbook snapshots. Outputs vm/validation_report.xlsx '
            'with colour-coded pass/fail per division and per field.'
        ),
        bullet('Compares COLZOPsRev, COLZOPsCost, COLZGAAPRev, COLZGAAPCost, COLZOPsBonus, CompletionDate'),
        bullet('Company-level missing tracking — jobs in Nicole\'s file found in no division snapshot'),
        bullet('$0.02 tolerance for float precision'),
        bullet('Status categories: MATCH (green), MISMATCH (red), NO_OVERRIDE (gray)'),
        ok('Script complete — will re-run once fresh 22 snapshots are copied from VM'),
        sp(4),
    ]

    # Architecture note
    story += [
        h2('6 — Distribution Workflow Issue (identified, not yet built)'),
        body(
            'Identified that distributing the workbook as .xltm to Project Managers will result '
            'in blank workbooks on open (Windows treats .xltm as a template, opens new blank copy). '
            'Even File→Open triggers Workbook_Open which may clear data via ClearFormOnOpen.'
        ),
        bullet('Fix: SaveAs .xlsm (macro-enabled workbook, not template) with ClearFormOnOpen=False'),
        bullet('Scope: part of the "Ready for Ops → Save & Distribute" action — not needed for current validation'),
        small('Documented in memory/distribution_workflow.md for Sprint 2/3 build.'),
        sp(4),
    ]

    # ── Current Status by Component ───────────────────────────────────────────
    story += [h1('Current Status by Component'), sp(2)]

    status_data = [
        [p('Component', S_SMALL), p('Status', S_SMALL), p('Notes', S_SMALL)],
        [p('Vista read path'),       p('✓ COMPLETE'),   p('All 4 sheets load. Zero WipDb dependency.')],
        [p('LylesWIP DB'),           p('✓ COMPLETE'),   p('4,973 rows for Dec 2025 across all 4 companies. Job numbers normalized.')],
        [p('Override merge (read)'), p('✓ FIXED TODAY'),p('SummaryData range bug fixed. All 22 combos re-running now.')],
        [p('Validation report'),     p('⏳ IN PROGRESS'),p('Awaiting fresh snapshots. validate_wip.py ready to run.')],
        [p('Nicole/Cindy review'),   p('⏳ PENDING'),   p('Will use validation_report.xlsx as agenda. Pending clean validation run.')],
        [p('Write-back (save row)'), p('🔴 NOT BUILT'), p('Double-click Col H → LylesWIPSaveJobRow. Sprint 2 priority.')],
        [p('Batch state workflow'),  p('🔴 NOT BUILT'), p('RFO → OpsApproved → AcctApproved state machine. Sprint 2.')],
        [p('Distribution (.xlsm)'),  p('🔴 NOT BUILT'), p('SaveAs .xlsm + ClearFormOnOpen=False on distribute.')],
        [p('Permissions module'),    p('🔴 NOT DEPLOYED'),p('Original Permissions.bas still active (pnp.WIPSECGetRole). Must replace.')],
        [p('GAAP write-back'),       p('🔴 NOT BUILT'), p('Col I double-click → save GAAP row. Sprint 3.')],
    ]
    story += [tbl(status_data, [1.8*inch, 1.3*inch, W - 3.1*inch]), sp(8)]

    # ── Immediate Next Steps ───────────────────────────────────────────────────
    story += [h1('Immediate Next Steps'), sp(2)]

    next_data = [
        [p('#', S_SMALL), p('Action', S_SMALL), p('Owner', S_SMALL), p('Blocking', S_SMALL)],
        [p('1'), p('BatchValidateAll finishes → copy 22 files from VM to Mac vm/validate-d3/'), p('Josh'), p('Yes — blocks #2')],
        [p('2'), p('Run python3 validate_wip.py — expect 344 mismatches to clear'), p('Claude'), p('Yes — blocks #3')],
        [p('3'), p('Review validation_report.xlsx — confirm PASS across all 22 divisions'), p('Josh + Claude'), p('Yes — blocks Nicole session')],
        [p('4'), p('Schedule Nicole/Cindy review session with validation_report.xlsx as agenda'), p('Josh'), p('No')],
        [p('5'), p('Deploy Permissions_Modified.bas to workbook (fixes Security Settings Not Valid popup)'), p('Josh + Claude'), p('No')],
        [p('6'), p('Build write-back: LylesWIPSaveJobRow wired to double-click Col H (Ops Done)'), p('Claude'), p('No')],
        [p('7'), p('Build batch state machine: RFO → OpsApproved → AcctApproved'), p('Claude'), p('No')],
    ]
    story += [tbl(next_data, [0.3*inch, W - 1.7*inch, 0.8*inch, 0.6*inch]), sp(8)]

    # ── Known Risks ───────────────────────────────────────────────────────────
    story += [h1('Known Risks & Open Questions'), sp(2)]

    risks_data = [
        [p('Risk / Question', S_SMALL), p('Impact', S_SMALL), p('Mitigation', S_SMALL)],
        [
            p('NESM Div35: 57 Nicole jobs in LylesWIP but 0 Vista WB jobs'),
            p('Medium'),
            p('Vista has no Dec 2025 activity for these jobs. Confirm with Nicole — likely legitimately inactive.')
        ],
        [
            p('52.8712 City of Avenal: in Vista but not in Nicole Dec 2025 file'),
            p('Low'),
            p('Nicole confirmed (2026-04-02): OK to proceed without adding. New/small job she didn\'t review.')
        ],
        [
            p('Jobs-Ops vs GAAP (Sheet13): not loaded by batch process'),
            p('Medium'),
            p('Cannot validate from snapshots. Must review manually in live workbook during Nicole session.')
        ],
        [
            p('Permissions module not deployed — Security Settings Not Valid popup on workbook open'),
            p('High'),
            p('Permissions_Modified.bas written and ready. Deploy before Nicole review session.')
        ],
        [
            p('Distribution as .xltm causes blank workbook for PMs'),
            p('High'),
            p('Identified today. Fix designed (SaveAs .xlsm + ClearFormOnOpen=False). Build in Sprint 2.')
        ],
        [
            p('16-70 (AIC Company-level overhead): snapshot not yet generated'),
            p('Low'),
            p('Added to BatchValidate.bas. Will be generated in current batch run.')
        ],
    ]
    story += [tbl(risks_data, [2.2*inch, 0.7*inch, W - 2.9*inch]), sp(8)]

    # ── Technical Reference ───────────────────────────────────────────────────
    story += [h1('Technical Reference'), sp(2)]

    story += [
        h2('Key Files Modified Today'),
    ]
    files_data = [
        [p('File', S_SMALL), p('Change', S_SMALL)],
        [p('sql/load_overrides.py'),            p('normalize_job() updated: right-pads job segment to 4 digits with ljust(4, "0")')],
        [p('vba_source/LylesWIPData.bas'),      p('All 3 Merge* functions: dynamic row count via xlUp scan instead of summaryRange.Rows.Count')],
        [p('vba_source/BatchValidate.bas'),     p('New module: automates 22-division snapshot generation. OUTPUT_PATH=C:\\Trusted\\validate-d3\\')],
        [p('validate_wip.py'),                  p('New script: compares Nicole Dec 2025 files vs 22 snapshots. Outputs validation_report.xlsx')],
        [p('memory/d2_d3_validation_notes.md'), p('Updated with job number format fix details and SummaryData range bug root cause')],
        [p('memory/distribution_workflow.md'),  p('New: documents .xltm distribution problem and .xlsm fix for Sprint 2')],
    ]
    story += [tbl(files_data, [2.4*inch, W - 2.4*inch]), sp(6)]

    story += [
        h2('Database State — LylesWIP (10.103.30.11)'),
    ]
    db_data = [
        [p('Company', S_SMALL), p('Code', S_SMALL), p('WipMonth', S_SMALL), p('Rows', S_SMALL)],
        [p('W. M. Lyles Co. (WML)'),   p('15'), p('2025-12-01'), p('311')],
        [p('Adv. Integration (AIC)'),  p('16'), p('2025-12-01'), p('TBD — re-counted after normalize fix')],
        [p('American Paving (APC)'),   p('12'), p('2025-12-01'), p('TBD')],
        [p('NESM'),                     p('13'), p('2025-12-01'), p('TBD')],
        [p('TOTAL (all months)'),       p('—'),  p('All'),        p('4,973 upserted today')],
    ]
    story += [tbl(db_data, [2.2*inch, 0.6*inch, 1.1*inch, W - 3.9*inch]), sp(6)]

    # ── Footer ────────────────────────────────────────────────────────────────
    story += [
        rule(), sp(4),
        Paragraph(
            f'Auto-WIP Daily Status Report &nbsp;|&nbsp; {TODAY} &nbsp;|&nbsp; '
            'CONFIDENTIAL — Lyles Services Co. internal use only',
            make_style('Footer', fontSize=7, textColor=colors.HexColor('#888888'),
                       alignment=TA_CENTER)
        ),
    ]

    return story


# =============================================================================
# Markdown version
# =============================================================================
MD_CONTENT = f"""# Auto-WIP Daily Status Report — {TODAY}

**Project:** Automated WIP Schedule — Lyles Services Co.
**Prepared by:** Josh Garrison / Claude Code
**Priority:** HIGHEST — escalated by owner, board, and senior management

---

## Executive Summary

The Automated WIP Schedule rebuild is in its **validation phase** for the December 2025 period. Two significant bugs were identified and resolved today:

1. Job number format mismatch in the LylesWIP override database (3-digit vs 4-digit)
2. Critical VBA named-range bug causing overrides to silently drop for divisions with >18 jobs

The fix has been deployed. A fresh batch of 22 division snapshots is being regenerated on the VM now. The read path is fully functional. Write-back to LylesWIP is the primary remaining work.

---

## Today's Accomplishments (2026-04-02)

### 1 — Job Number Format Fix (`sql/load_overrides.py`)
- Nicole's historical files stored 3-digit job segments (`73.105.`) vs Vista's 4-digit (`73.1050.`)
- Deleted 132 short-format rows from LylesWIP for Co12 and Co16
- Updated `normalize_job()` to right-pad with `ljust(4, '0')`
- Re-ran load_overrides.py → **4,973 rows upserted** across 40 historical files
- ✓ LylesWIP now contains correctly keyed overrides for all four companies

### 2 — Root Cause Found: SummaryData Static Named Range Bug
- `SummaryData` initialized to **18 rows** (`$A$7:$CM$24`) in the workbook template
- `GetWipDetail2` writes jobs using `SummaryData.Cells(r, …)` — VBA silently extends beyond the range
- `MergeOverridesOntoSheet`, `MergePriorMonthProfitsOntoSheet`, `MergePriorYearBonusOntoSheet` all looped `For r = 1 To summaryRange.Rows.Count` = 18
- ⚠ **Any division with >18 jobs lost all overrides beyond the 18th job** (WML Div51: 12 missed; AIC Div73: 67 missed)

### 3 — Fix Applied (`vba_source/LylesWIPData.bas`)
All three Merge* functions now use dynamic row detection:
```vba
lastDataRow = sh.Cells(sh.Rows.Count, jnColAbs).End(xlUp).Row
totalRows = Application.Max(summaryRange.Rows.Count, lastDataRow - summaryRange.Row + 1)
For r = 1 To totalRows
```
- ✓ Deployed to workbook Rev 5.68p on VM

### 4 — BatchValidate.bas — Automated Snapshot Generation
Created to replace 20–30 minutes of manual clicking. Loops all 22 combinations:

| Company | Divisions | Count |
|---------|-----------|-------|
| WML (Co15)  | 51–58 | 8 |
| AIC (Co16)  | 70–78 | 9 |
| APC (Co12)  | 21    | 1 |
| NESM (Co13) | 31, 32, 33, 35 | 4 |
| **TOTAL** | | **22** |

- Uses `ResetWorkbook` (not `ClearForms3`) to safely clear between iterations
- OUTPUT_PATH: `C:\\Trusted\\validate-d3\\`
- ✓ Running now with fixed LylesWIPData.bas

### 5 — validate_wip.py — Python Validation Script
Compares Nicole's Dec 2025 WIP History Import files vs 22 workbook snapshots:
- Checks: COLZOPsRev, COLZOPsCost, COLZGAAPRev, COLZGAAPCost, COLZOPsBonus, CompletionDate
- Status: MATCH / MISMATCH / NO_OVERRIDE — output to `vm/validation_report.xlsx`
- ✓ Script complete — will re-run once fresh snapshots are copied from VM

### 6 — Distribution Workflow Issue Identified
- Distributing `.xltm` to PMs causes blank workbook (Windows opens .xltm as template)
- Fix: `SaveAs .xlsm` + `ClearFormOnOpen = False` on distribute action
- Documented in `memory/distribution_workflow.md` for Sprint 2/3 build

---

## Current Status by Component

| Component | Status | Notes |
|-----------|--------|-------|
| Vista read path | ✓ COMPLETE | All 4 sheets load. Zero WipDb dependency. |
| LylesWIP DB | ✓ COMPLETE | 4,973 rows Dec 2025, all 4 companies, normalized keys |
| Override merge (read) | ✓ FIXED TODAY | SummaryData range bug fixed. 22 combos re-running. |
| Validation report | ⏳ IN PROGRESS | Awaiting fresh snapshots |
| Nicole/Cindy review | ⏳ PENDING | Pending clean validation run |
| Write-back (save row) | 🔴 NOT BUILT | Col H double-click → LylesWIPSaveJobRow (Sprint 2) |
| Batch state workflow | 🔴 NOT BUILT | RFO → OpsApproved → AcctApproved (Sprint 2) |
| Distribution (.xlsm) | 🔴 NOT BUILT | SaveAs .xlsm + ClearFormOnOpen=False (Sprint 2) |
| Permissions module | 🔴 NOT DEPLOYED | Original still active — Security popup on open |
| GAAP write-back | 🔴 NOT BUILT | Col I double-click (Sprint 3) |

---

## Immediate Next Steps

1. **BatchValidateAll finishes** → copy 22 files from VM to Mac `vm/validate-d3/`
2. **Run `python3 validate_wip.py`** — expect 344 mismatches to clear
3. **Review `validation_report.xlsx`** — confirm PASS across all 22 divisions
4. **Schedule Nicole/Cindy review session** with validation report as agenda
5. **Deploy Permissions_Modified.bas** — fixes Security Settings Not Valid popup
6. **Build write-back:** `LylesWIPSaveJobRow` wired to Col H double-click
7. **Build batch state machine:** RFO → OpsApproved → AcctApproved

---

## Known Risks & Open Questions

| Risk | Impact | Mitigation |
|------|--------|-----------|
| NESM Div35: 57 Nicole jobs, 0 Vista WB jobs | Medium | Vista has no Dec 2025 activity. Confirm with Nicole. |
| 52.8712 City of Avenal: in Vista not in Nicole file | Low | Nicole confirmed OK to proceed (2026-04-02) |
| Jobs-Ops vs GAAP (Sheet13): not batch-loadable | Medium | Review manually in live workbook during Nicole session |
| Permissions module not deployed | High | Permissions_Modified.bas ready — deploy before Nicole session |
| .xltm distribution causes blank workbook for PMs | High | Fix designed — build in Sprint 2 |

---

## Key Files Modified Today

| File | Change |
|------|--------|
| `sql/load_overrides.py` | normalize_job() 4-digit padding |
| `vba_source/LylesWIPData.bas` | Dynamic row count in all 3 Merge* functions |
| `vba_source/BatchValidate.bas` | New — 22-division batch snapshot automation |
| `validate_wip.py` | New — Python validation script |
| `memory/d2_d3_validation_notes.md` | Updated with today's bug findings |
| `memory/distribution_workflow.md` | New — .xlsm distribution fix spec |

---

*CONFIDENTIAL — Lyles Services Co. internal use only*
"""


# =============================================================================
# Build PDF
# =============================================================================
def main():
    doc = SimpleDocTemplate(
        PDF_FILE,
        pagesize=letter,
        leftMargin=inch, rightMargin=inch,
        topMargin=0.75*inch, bottomMargin=0.75*inch,
        title=f'Auto-WIP Status Report {TODAY}',
        author='Josh Garrison / Claude Code',
    )
    doc.build(build_story())
    print(f'PDF saved:      {PDF_FILE}')

    with open(MD_FILE, 'w') as f:
        f.write(MD_CONTENT)
    print(f'Markdown saved: {MD_FILE}')


if __name__ == '__main__':
    main()
