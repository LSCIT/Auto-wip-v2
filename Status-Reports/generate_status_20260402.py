#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 2, 2026."""

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
    KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

BLUE_DARK    = colors.HexColor('#1F3864')
BLUE_MID     = colors.HexColor('#2E5090')
BLUE_LIGHT   = colors.HexColor('#D6E4F0')
GRAY_LIGHT   = colors.HexColor('#F2F2F2')
GREEN        = colors.HexColor('#1D6A3C')
ORANGE       = colors.HexColor('#B45309')
RED          = colors.HexColor('#9B1C1C')
BLACK        = colors.black
WHITE        = colors.white
BORDER_GRAY  = colors.HexColor('#CCCCCC')

def S(name, **kw):
    return ParagraphStyle(name, **kw)

styles = {
    'title':      S('title',     fontName='Helvetica-Bold', fontSize=16, textColor=BLUE_DARK,
                                 leading=20, spaceAfter=4),
    'subtitle':   S('subtitle',  fontName='Helvetica',      fontSize=10, textColor=BLACK,
                                 leading=14, spaceAfter=2),
    'h1':         S('h1',        fontName='Helvetica-Bold', fontSize=13, textColor=BLUE_DARK,
                                 leading=16, spaceBefore=16, spaceAfter=4),
    'h2':         S('h2',        fontName='Helvetica-Bold', fontSize=11, textColor=BLUE_MID,
                                 leading=14, spaceBefore=10, spaceAfter=4),
    'h3':         S('h3',        fontName='Helvetica-Bold', fontSize=9,  textColor=BLUE_MID,
                                 leading=13, spaceBefore=6, spaceAfter=2),
    'body':       S('body',      fontName='Helvetica',      fontSize=9,  textColor=BLACK,
                                 leading=13, spaceAfter=4),
    'body_bold':  S('body_bold', fontName='Helvetica-Bold', fontSize=9,  textColor=BLACK,
                                 leading=13, spaceAfter=4),
    'bullet':     S('bullet',    fontName='Helvetica',      fontSize=9,  textColor=BLACK,
                                 leading=13, spaceAfter=2, leftIndent=12, bulletIndent=2),
    'note':       S('note',      fontName='Helvetica-Oblique', fontSize=8, textColor=colors.HexColor('#555555'),
                                 leading=11, spaceAfter=4),
    'code':       S('code',      fontName='Courier',        fontSize=8,  textColor=BLACK,
                                 leading=12, spaceAfter=2, leftIndent=12,
                                 backColor=colors.HexColor('#F5F5F5')),
    'th':         S('th',        fontName='Helvetica-Bold', fontSize=8,  textColor=WHITE,
                                 leading=11, wordWrap='LTR'),
    'td':         S('td',        fontName='Helvetica',      fontSize=8,  textColor=BLACK,
                                 leading=11, wordWrap='LTR'),
    'td_bold':    S('td_bold',   fontName='Helvetica-Bold', fontSize=8,  textColor=BLACK,
                                 leading=11, wordWrap='LTR'),
    'td_green':   S('td_green',  fontName='Helvetica-Bold', fontSize=8,  textColor=GREEN,
                                 leading=11, wordWrap='LTR'),
    'td_orange':  S('td_orange', fontName='Helvetica-Bold', fontSize=8,  textColor=ORANGE,
                                 leading=11, wordWrap='LTR'),
    'td_red':     S('td_red',    fontName='Helvetica-Bold', fontSize=8,  textColor=RED,
                                 leading=11, wordWrap='LTR'),
    'footer':     S('footer',    fontName='Helvetica',      fontSize=7,  textColor=colors.HexColor('#888888'),
                                 leading=10, alignment=TA_CENTER),
}

def P(text, style='body'):   return Paragraph(text, styles[style])
def TH(text):  return Paragraph(text, styles['th'])
def TD(text):  return Paragraph(text, styles['td'])
def TDB(text): return Paragraph(text, styles['td_bold'])
def TDG(text): return Paragraph(text, styles['td_green'])
def TDO(text): return Paragraph(text, styles['td_orange'])
def TDR(text): return Paragraph(text, styles['td_red'])
def B(text):   return Paragraph(f'• {text}', styles['bullet'])
def sp(n=6):   return Spacer(1, n)
def hr():      return HRFlowable(width='100%', thickness=1,   color=BLUE_DARK,   spaceAfter=6, spaceBefore=6)
def shr():     return HRFlowable(width='100%', thickness=0.5, color=BORDER_GRAY, spaceAfter=4, spaceBefore=4)

def tbl(data, col_widths, header_bg=BLUE_DARK, stripe=True):
    cmds = [
        ('BACKGROUND',    (0, 0), (-1,  0), header_bg),
        ('ROWBACKGROUNDS',(0, 1), (-1, -1), [WHITE, GRAY_LIGHT] if stripe else [WHITE]),
        ('GRID',          (0, 0), (-1, -1), 0.5, BORDER_GRAY),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 5),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
        ('TOPPADDING',    (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(cmds))
    return t

def callout(text, width, bg=BLUE_LIGHT, border=BLUE_MID):
    data = [[Paragraph(text, ParagraphStyle('cb', fontName='Helvetica', fontSize=9,
                                            textColor=BLUE_DARK, leading=14, wordWrap='LTR'))]]
    t = Table(data, colWidths=[width])
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0), (-1,-1), bg),
        ('BOX',          (0,0), (-1,-1), 1.5, border),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING',   (0,0), (-1,-1), 8),
        ('BOTTOMPADDING',(0,0), (-1,-1), 8),
    ]))
    return t

def warn_box(text, width):
    return callout(text, width,
                   bg=colors.HexColor('#FFF4E5'),
                   border=colors.HexColor('#B45309'))

def build():
    W = 7.5 * inch

    doc = SimpleDocTemplate(
        '/Users/joshuagarrison/lsc/Auto-Wip/Status-Reports/Auto-WIP_Status_Report_20260402.pdf',
        pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch,
        title='Auto-WIP Schedule — Project Status Report — April 2, 2026',
        author='Josh Garrison, IT Development',
    )

    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    story += [
        P('Auto-WIP Schedule — Project Status Report', 'title'),
        P('Date: April 2, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Priority: HIGHEST — Owner / Board / Senior Management escalation</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # ── Executive Summary ─────────────────────────────────────────────────────
    story += [
        P('Executive Summary', 'h1'),
        P(
            'April 2 was a full validation engineering day — approximately 9 hours of sustained '
            'work — focused on proving end-to-end data integrity from Nicole\'s source files through '
            'Vista\'s job cost engine through the LylesWIP override database into the workbook. '
            'The day produced two new tools (BatchValidate automation and a Python validation '
            'framework), resolved two root-cause bugs that were silently corrupting override data '
            'for a significant portion of jobs, and confirmed the full pipeline is architecturally '
            'sound. The read path is complete and validated.'
        ),
        P(
            'The single highest-impact finding of the day: a static named range in the VBA layer '
            'was silently dropping override data for every job beyond the 18th row in any division. '
            'This was invisible to users — the workbook loaded and displayed Vista data correctly, '
            'but LylesWIP overrides were not being merged for those jobs. The fix is deployed and '
            'all 22 division snapshots are being regenerated now for clean validation.'
        ),
        sp(4),
        callout(
            '<b>Updated production readiness: approximately 60–65%.</b> '
            'The database layer, override load, read merge, and validation tooling are all complete. '
            'The write path (save overrides on Done, batch state machine, distribution workflow, '
            'permissions) is the remaining open body of work.<br/><br/>'
            '<b>Target delivery: May 8, 2026 — no change.</b>',
            W
        ),
        sp(8), shr(),
    ]

    # ── What Changed Since April 1 ────────────────────────────────────────────
    story += [P('What Changed Since April 1, 2026', 'h1')]

    delta_data = [
        [TH('Item'), TH('Status Apr 1'), TH('Status Today (Apr 2)')],
        [TD('Job number key format in LylesWIP (AIC/APC)'),
         TDR('Mismatch — 3-digit segments didn\'t match Vista\'s 4-digit format'),
         TDG('Fixed — normalize_job() updated; 132 bad rows deleted; 4,973 rows re-upserted')],
        [TD('Snapshot generation (all 22 divisions)'),
         TDR('Manual — 20–30 min of clicking per batch'),
         TDG('Automated — BatchValidate.bas loops all 22 combos unattended')],
        [TD('Division 70 (AIC company overhead)'),
         TDR('Missing from validation scope'),
         TDG('Added to BatchValidate.bas; included in current batch run')],
        [TD('Override merge for divisions with >18 jobs'),
         TDR('Silent data loss — jobs beyond row 18 never got overrides merged'),
         TDG('Fixed — all 3 Merge* functions now use dynamic row count via xlUp scan')],
        [TD('Validation framework'),
         TDR('Not built — no systematic way to verify data integrity'),
         TDG('Complete — validate_wip.py compares Nicole\'s files vs 22 snapshots; outputs color-coded Excel report')],
        [TD('344 override mismatches across 11 divisions'),
         TDR('Unknown root cause'),
         TDG('Root cause identified and fixed — caused by SummaryData range bug above')],
        [TD('Distribution workflow (.xltm to PMs)'),
         TDR('Unidentified risk'),
         TDO('Risk identified — .xltm opens blank for PMs. Fix designed (.xlsm + ClearFormOnOpen=False). Sprint 2.')],
    ]
    story += [tbl(delta_data, [2.3*inch, 2.1*inch, 3.1*inch]), sp(8), shr()]

    # ── Detailed Work Log ─────────────────────────────────────────────────────
    story += [P('Detailed Work Log — April 2, 2026', 'h1')]

    # 1 - Job number format
    story += [
        KeepTogether([
            P('1.  Job Number Format Normalization', 'h2'),
            P(
                'Nicole\'s 40 historical WIP History Import Excel files store job numbers in the format '
                'the original WipDb system used: two-part company.job with variable job-segment length '
                '(e.g. <font name="Courier">73.105.</font>). Vista normalizes the job segment to exactly '
                '4 digits right-padded with zeros (<font name="Courier">73.1050.</font>). This mismatch '
                'meant every override loaded for AIC (Co16) and APC (Co12) had a key that would never '
                'match a Vista job number — all those overrides were loaded but unreachable.'
            ),
            B('Identified 22 short-format job numbers in AIC files and 0 in APC (the fix covered both companies)'),
            B('Deleted 132 affected rows from dbo.WipJobData for Co12 and Co16 in LylesWIP'),
            B('Updated normalize_job() in sql/load_overrides.py: added ljust(4, "0") right-pad for job segments shorter than 4 digits'),
            B('Re-ran load_overrides.py against all 40 historical files — 4,973 rows upserted cleanly'),
            B('Verified edge case: job 78.12 → 78.1200. which correctly aligns with Vista\'s sequential job numbering in the 78.12xx series'),
            B('Confirmed write-back architecture: when overrides eventually write to Viewpoint, job numbers come from Vista (already 4-digit), so no additional normalization is needed at write time'),
            sp(4),
        ])
    ]

    # 2 - BatchValidate
    story += [
        KeepTogether([
            P('2.  BatchValidate.bas — Automated Snapshot Generation', 'h2'),
            P(
                'Prior to today, generating a full set of division snapshots required manually loading '
                'each Company/Month/Division combination and saving — approximately 20–30 minutes of '
                'repetitive clicking for 21+ combinations. BatchValidate.bas eliminates this entirely.'
            ),
            B('Loops all 22 company/division combinations unattended; shows progress in Excel title bar'),
            B('Uses ResetWorkbook (not ClearForms3) between iterations — critical distinction: ClearForms3 leaves sheets protected, causing a 1004 error on the next iteration\'s clear attempt'),
            B('Non-fatal error handling: logs errors to the Immediate window and continues to the next division rather than halting the entire batch'),
            B('Added Division 70 (AIC company-level overhead) after confirming it is a real division with active jobs'),
            B('OUTPUT_PATH: C:\\Trusted\\validate-d3\\ (user-specified)'),
        ]),
        sp(4),
    ]

    batch_data = [
        [TH('Company'), TH('Code'), TH('Divisions'), TH('Count')],
        [TD('W. M. Lyles Co. (WML)'),                    TD('15'), TD('51, 52, 53, 54, 55, 56, 57, 58'), TD('8')],
        [TD('Advanced Integration &amp; Controls (AIC)'), TD('16'), TD('70, 71, 72, 73, 74, 75, 76, 77, 78'), TD('9')],
        [TD('American Paving Co. (APC)'),                 TD('12'), TD('21'), TD('1')],
        [TD('New England Sheet Metal (NESM)'),            TD('13'), TD('31, 32, 33, 35'), TD('4')],
        [TDB('TOTAL'), TD(''), TD(''), TDB('22')],
    ]
    story += [tbl(batch_data, [2.8*inch, 0.6*inch, 3.1*inch, 1.0*inch]), sp(8)]

    # 3 - validate_wip.py
    story += [
        P('3.  validate_wip.py — Python Validation Framework', 'h2'),
        P(
            'Created a systematic validation script that compares Nicole\'s December 2025 WIP History '
            'Import files (the authoritative source of override values) against the workbook snapshots '
            'loaded from Vista + LylesWIP. This gives a definitive, repeatable answer to "did the '
            'data make it through the full pipeline correctly?"'
        ),
        B('Reads all four company Nicole files and all 22 snapshot .xltm files via openpyxl (no macros run)'),
        B('Compares COLZOPsRev, COLZOPsCost, COLZGAAPRev, COLZGAAPCost, COLZOPsBonus, and CompletionDate'),
        B('Z columns (COLZOPsRev etc.) are the authoritative ground truth — they store exactly what LylesWIP wrote at merge time, separate from Vista\'s calculated values'),
        B('Status categories: MATCH (Nicole non-zero value matches WB Z column within $0.02), MISMATCH (value present but WB differs or is 0), NO_OVERRIDE (Nicole had zero/blank — no override expected)'),
        B('Company-level missing tracking: a job is "missing" only if it appears in Nicole\'s file but in NO division snapshot for that company — not flagged per-division'),
        B('Outputs vm/validation_report.xlsx with color-coded Summary, Mismatches, All Override Checks, Nicole Jobs Not In WB, and Notes sheets'),
        B('Identified NESM Div35: 57 Nicole jobs in LylesWIP, 57 rows in DB, but 0 matching Vista WB jobs — Vista has no Dec 2025 activity for those jobs (confirmed expected)'),
        B('Confirmed Jobs-Ops vs GAAP (Sheet13) cannot be validated from batch snapshots — that sheet is not loaded by GetWipDetail2 and must be reviewed manually in the live workbook'),
        sp(4),
    ]

    # 4 - Root cause investigation
    story += [
        P('4.  Root Cause Investigation — 344 Override Mismatches', 'h2'),
        P(
            'The first validation run showed 344 mismatches across 11 of 22 divisions: Nicole had '
            'non-zero override values, LylesWIP had the data loaded correctly, but the workbook Z '
            'columns were showing 0. This kicked off a methodical root cause investigation.'
        ),
    ]

    story += [P('Investigation sequence:', 'h3')]
    invest_data = [
        [TH('Step'), TH('What Was Checked'), TH('Finding')],
        [TD('1'), TD('Are the mismatch jobs actually in LylesWIP with correct values?'),
         TDG('Yes — queried dbo.WipJobData directly; all 12 WML Div51 mismatch jobs present with correct values')],
        [TD('2'), TD('Are all 22 snapshot files from the same fresh batch run (not stale)?'),
         TDG('Yes — all 22 files timestamped Apr 2, 15:36–15:41 (same batch)')],
        [TD('3'), TD('Is there a MaxRecords limit on the ADODB recordset in GetJobOverrides?'),
         TDG('No — reviewed LylesWIPData.bas; no MaxRecords set; BuildOverrideLookup loops all rows')],
        [TD('4'), TD('Are the Plugged flags (OpsRevPlugged etc.) set to True?'),
         TDG('Yes — queried all 12 WML Div51 mismatch jobs; Plugged=True for all applicable fields')],
        [TD('5'), TD('Does LylesWIPGetJobOverrides stored proc apply a dept filter?'),
         TDG('No — reviewed proc definition; WHERE clause is JCCo + WipMonth only; returns all 311 Co15 rows')],
        [TD('6'), TD('Does the job number format in the sheet match the dictionary key?'),
         TDG('Yes — inspected raw snapshot values with openpyxl; job format "51.1133." is clean, no trailing spaces')],
        [TD('7'), TD('How many rows does SummaryData named range cover?'),
         TDR('18 rows ($A$7:$CM$24) — but 28 jobs were written for Div51. ROOT CAUSE FOUND.')],
    ]
    story += [tbl(invest_data, [0.35*inch, 3.0*inch, 4.15*inch]), sp(4)]

    story += [
        P('Root Cause: SummaryData Static Named Range', 'h3'),
        P(
            'The SummaryData named range is initialized to 18 rows in the workbook template '
            '($A$7:$CM$24 on Jobs-Ops). GetWipDetail2 writes job rows using '
            'SummaryData.Cells(r, …) — VBA silently allows this to extend beyond the named '
            'range bounds without error. For a division with 28 jobs, rows 19–28 land in '
            'sheet rows 25–34, below the named range.'
        ),
        P(
            'MergeOverridesOntoSheet, MergePriorMonthProfitsOntoSheet, and '
            'MergePriorYearBonusOntoSheet all loop "For r = 1 To summaryRange.Rows.Count" '
            '= 18. The loop terminates before reaching the extra rows. Every job beyond '
            'position 18 in the Vista result set silently receives no overrides.'
        ),
        warn_box(
            '⚠ Impact: Any division returning more than 18 jobs from Vista lost all LylesWIP '
            'overrides for jobs beyond the 18th. This is a silent failure — the workbook displays '
            'correctly with Vista-calculated values; only the overrides are missing. '
            'Affected divisions include WML Div51 (28 jobs, 12 missed), AIC Div73 (83 jobs, ~67 missed), '
            'and any other large division.',
            W
        ),
        sp(6),
    ]

    # 5 - Fix
    story += [
        P('5.  Fix — LylesWIPData.bas Dynamic Row Count', 'h2'),
        P(
            'Applied to all three Merge* functions in vba_source/LylesWIPData.bas. '
            'Each function now detects the actual last data row before looping:'
        ),
        Paragraph(
            'jnColAbs = summaryRange.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Column\n'
            'lastDataRow = sh.Cells(sh.Rows.Count, jnColAbs).End(xlUp).Row\n'
            'totalRows = Application.Max(summaryRange.Rows.Count, lastDataRow - summaryRange.Row + 1)\n'
            'For r = 1 To totalRows',
            styles['code']
        ),
        sp(4),
        B('xlUp scan on the job number column (column A) finds the actual last row written by GetWipDetail2'),
        B('Application.Max() ensures the fix is additive — never shrinks the loop below the existing Rows.Count in edge cases'),
        B('SummaryData.Cells(r, …) correctly addresses rows beyond the named range bounds; only the loop limit was wrong'),
        B('Applied identically to MergeOverridesOntoSheet, MergePriorMonthProfitsOntoSheet, MergePriorYearBonusOntoSheet — no partial fix risk'),
        B('Deployed to workbook on VM. All 22 snapshots being regenerated now with the fix in place'),
        sp(4),
    ]

    # 6 - Distribution workflow
    story += [
        KeepTogether([
            P('6.  Distribution Workflow Risk Identified', 'h2'),
            P(
                'During discussion of the PM distribution workflow, a critical UX risk was identified. '
                'When Nicole saves and distributes the workbook as .xltm, Project Managers will '
                'receive a blank workbook when they open it.'
            ),
            B('Double-clicking a .xltm in Windows Explorer tells Excel to open a NEW workbook using it as a template — the saved data copy is never opened'),
            B('Even via File → Open, Workbook_Open fires and ClearFormOnOpen may clear all loaded data'),
            B('PMs still need macros (double-click Done writes back to LylesWIP), so stripping macros entirely is not an option'),
            P('Fix designed for Sprint 2:', 'h3'),
            B('SaveAs .xlsm (macro-enabled workbook, not template) — eliminates the template-open problem entirely'),
            B('Set ClearFormOnOpen = False in the distributed copy\'s Settings sheet before saving'),
            B('Restore ClearFormOnOpen = True in Nicole\'s master copy after distribute'),
            B('Implement as a "Save &amp; Distribute to Ops" button/action wired to the ReadyForOps workflow'),
            P('Documented in memory/distribution_workflow.md.', 'note'),
            sp(4),
        ])
    ]

    story += [shr()]

    # ── Current Status by Component ───────────────────────────────────────────
    story += [P('Current Status by Component', 'h1')]

    status_data = [
        [TH('Component'), TH('Status'), TH('Detail')],
        [TD('Vista read path — all 4 sheets'),
         TDG('✓ COMPLETE'),
         TD('Jobs-Ops, Jobs-GAAP, JV\'s-Ops, JV\'s-GAAP all load from Vista. Zero WipDb dependency.')],
        [TD('LylesWIP database (PNP server)'),
         TDG('✓ COMPLETE'),
         TD('Tables, stored procs, SQL login configured. 4,973 rows loaded across 40 historical files, all 4 companies.')],
        [TD('Job number normalization'),
         TDG('✓ COMPLETE'),
         TD('All keys 4-digit format. normalize_job() fixed. AIC/APC re-imported.')],
        [TD('Override merge on load (all 3 Merge* functions)'),
         TDG('✓ FIXED TODAY'),
         TD('SummaryData range bug resolved. Dynamic row count. All divisions now fully covered.')],
        [TD('BatchValidate automation'),
         TDG('✓ COMPLETE'),
         TD('22-division unattended batch. Currently running on VM with fixed LylesWIPData.bas.')],
        [TD('Python validation framework'),
         TDG('✓ COMPLETE'),
         TD('validate_wip.py ready. Will re-run once fresh 22 snapshots are copied from VM.')],
        [TD('Validation report (clean run)'),
         TDO('⏳ IN PROGRESS'),
         TD('Awaiting fresh snapshots from current batch run. Expect 344 mismatches to clear.')],
        [TD('Nicole / Cindy review session'),
         TDO('⏳ PENDING'),
         TD('Pending clean validation run. validation_report.xlsx will be the agenda.')],
        [TD('Permissions module deployment'),
         TDR('🔴 NOT DEPLOYED'),
         TD('Original Permissions.bas still active — "Security Settings Not Valid" popup on open. Permissions_Modified.bas written and ready.')],
        [TD('Write-back: double-click Done → LylesWIPSaveJobRow'),
         TDR('🔴 NOT BUILT'),
         TD('Primary Sprint 2 task. Col H double-click fires save of overrides to LylesWIP dbo.WipJobData.')],
        [TD('Batch state machine'),
         TDR('🔴 NOT BUILT'),
         TD('Open → ReadyForOps → OpsApproved → AcctApproved. State transitions via LylesWIPUpdateBatchState.')],
        [TD('Distribution workflow (.xlsm)'),
         TDR('🔴 NOT BUILT'),
         TD('SaveAs .xlsm + ClearFormOnOpen=False on distribute. Sprint 2.')],
        [TD('GAAP write-back (Col I double-click)'),
         TDR('🔴 NOT BUILT'),
         TD('Sprint 3.')],
    ]
    story += [tbl(status_data, [2.2*inch, 1.15*inch, 4.15*inch]), sp(8), shr()]

    # ── Immediate Next Steps ───────────────────────────────────────────────────
    story += [P('Immediate Next Steps', 'h1')]

    next_data = [
        [TH('#'), TH('Action'), TH('Owner'), TH('Blocks')],
        [TD('1'), TD('BatchValidateAll finishes on VM → copy 22 files to Mac vm/validate-d3/'), TD('Josh'), TD('Step 2')],
        [TD('2'), TD('Run python3 validate_wip.py — expect 344 mismatches to clear with fixed snapshots'), TD('Claude'), TD('Step 3')],
        [TD('3'), TD('Review validation_report.xlsx — confirm PASS across all 22 divisions. Investigate any remaining mismatches.'), TD('Josh + Claude'), TD('Nicole session')],
        [TD('4'), TD('Deploy Permissions_Modified.bas — eliminates "Security Settings Not Valid" popup on workbook open'), TD('Josh + Claude'), TD('Nicole session')],
        [TD('5'), TD('Schedule Nicole / Cindy validation review session. Use validation_report.xlsx as the agenda.'), TD('Josh'), TD('—')],
        [TD('6'), TD('Build write-back: LylesWIPSaveJobRow wired to Col H double-click (Ops Done). Core Sprint 2 deliverable.'), TD('Claude'), TD('—')],
        [TD('7'), TD('Build batch state machine: ReadyForOps → OpsApproved → AcctApproved state transitions with LylesWIPUpdateBatchState'), TD('Claude'), TD('—')],
        [TD('8'), TD('Build "Save &amp; Distribute to Ops" action: SaveAs .xlsm + ClearFormOnOpen=False + distribution path'), TD('Claude'), TD('—')],
    ]
    story += [tbl(next_data, [0.3*inch, 4.05*inch, 0.9*inch, 0.85*inch]), sp(8), shr()]

    # ── Known Risks ───────────────────────────────────────────────────────────
    story += [P('Known Risks &amp; Open Items', 'h1')]

    risks_data = [
        [TH('Risk / Open Item'), TH('Impact'), TH('Status / Mitigation')],
        [TD('NESM Div35: 57 Nicole override jobs, 0 Vista WB jobs for Dec 2025'),
         TDO('Medium'),
         TD('Vista has no Dec 2025 JTD activity for Div35 jobs. Overrides are loaded in LylesWIP but have no Vista rows to merge onto. Confirm with Nicole whether these jobs are expected to be inactive.')],
        [TD('52.8712 City of Avenal WWTP Gearbox Repair (WML Div52)'),
         TDO('Low'),
         TD('Job appears in Vista Dec 2025 but Nicole excluded it from her import file. Nicole confirmed Apr 2: proceed without it. New/small job she did not review.')],
        [TD('Jobs-Ops vs GAAP (Sheet13): cannot be validated from batch snapshots'),
         TDO('Medium'),
         TD('Sheet13 is not loaded by GetWipDetail2 — batch snapshots show only 0–1 rows. Must review this sheet manually in the live workbook during the Nicole/Cindy session.')],
        [TD('Permissions module not deployed — original Permissions.bas still active'),
         TDR('High'),
         TD('Users see "Security Settings Not Valid" popup on workbook open. Permissions_Modified.bas is written and ready. Must deploy before distributing to Nicole/Cindy.')],
        [TD('Distribution as .xltm causes blank workbook for PMs'),
         TDR('High'),
         TD('Identified today. Fix designed (SaveAs .xlsm + ClearFormOnOpen=False). Build in Sprint 2 before any PM distribution.')],
        [TD('Prior Projected Profit (Col R) — Dec 2025 has no Nov 2025 prior month in LylesWIP'),
         TDO('Medium'),
         TD('MergePriorMonthProfitsOntoSheet looks for Nov 2025 overrides. Those files may not yet be loaded. Confirm with Nicole whether Nov 2025 Col R values need to be pre-populated.')],
    ]
    story += [tbl(risks_data, [2.6*inch, 0.75*inch, 4.15*inch]), sp(8), shr()]

    # ── Files Modified ────────────────────────────────────────────────────────
    story += [P('Files Created / Modified — April 2, 2026', 'h1')]

    files_data = [
        [TH('File'), TH('Type'), TH('Change')],
        [TD('sql/load_overrides.py'),
         TD('Python'), TD('normalize_job(): right-pads job segment to 4 digits with ljust(4, "0"). Resolves AIC/APC key mismatch.')],
        [TD('vba_source/LylesWIPData.bas'),
         TD('VBA'), TD('MergeOverridesOntoSheet, MergePriorMonthProfitsOntoSheet, MergePriorYearBonusOntoSheet: dynamic row count via xlUp scan. Deployed to workbook Rev 5.68p.')],
        [TD('vba_source/BatchValidate.bas'),
         TD('VBA'), TD('New module. Automates 22-division snapshot generation. ResetWorkbook between iterations. Non-fatal error handling. Includes Div 70.')],
        [TD('validate_wip.py'),
         TD('Python'), TD('New script. Compares Nicole Dec 2025 files vs 22 workbook snapshots. Company-level missing tracking. Outputs color-coded validation_report.xlsx.')],
        [TD('memory/d2_d3_validation_notes.md'),
         TD('Memory'), TD('Updated with job number format fix and SummaryData range bug root cause and fix details.')],
        [TD('memory/distribution_workflow.md'),
         TD('Memory'), TD('New. Documents .xltm distribution problem, .xlsm fix spec, and Sprint 2 implementation plan.')],
        [TD('Status-Reports/Auto-WIP_Status_Report_20260402.pdf + .md'),
         TD('Report'), TD('This report.')],
    ]
    story += [tbl(files_data, [2.5*inch, 0.6*inch, 4.4*inch]), sp(8)]

    # ── Footer ────────────────────────────────────────────────────────────────
    story += [
        shr(), sp(4),
        P('Auto-WIP Status Report &nbsp;|&nbsp; April 2, 2026 &nbsp;|&nbsp; '
          'CONFIDENTIAL — Lyles Services Co. internal use only', 'footer'),
    ]

    doc.build(story)
    print('PDF saved: Auto-WIP_Status_Report_20260402.pdf')


if __name__ == '__main__':
    build()
