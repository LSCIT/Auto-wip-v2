#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 1, 2026."""

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
    KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── Color palette ──────────────────────────────────────────────────────────────
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

# ── Styles ─────────────────────────────────────────────────────────────────────
def S(name, **kw):
    return ParagraphStyle(name, **kw)

styles = {
    'title':      S('title',      fontName='Helvetica-Bold',    fontSize=16, textColor=BLUE_DARK,
                                  leading=20, spaceAfter=4),
    'subtitle':   S('subtitle',   fontName='Helvetica',         fontSize=10, textColor=BLACK,
                                  leading=14, spaceAfter=2),
    'h1':         S('h1',         fontName='Helvetica-Bold',    fontSize=13, textColor=BLUE_DARK,
                                  leading=16, spaceBefore=16, spaceAfter=4),
    'h2':         S('h2',         fontName='Helvetica-Bold',    fontSize=11, textColor=BLUE_MID,
                                  leading=14, spaceBefore=10, spaceAfter=4),
    'body':       S('body',       fontName='Helvetica',         fontSize=9,  textColor=BLACK,
                                  leading=13, spaceAfter=4),
    'body_bold':  S('body_bold',  fontName='Helvetica-Bold',    fontSize=9,  textColor=BLACK,
                                  leading=13, spaceAfter=4),
    'note':       S('note',       fontName='Helvetica-Oblique', fontSize=8,  textColor=colors.HexColor('#555555'),
                                  leading=11, spaceAfter=4),
    'th':         S('th',         fontName='Helvetica-Bold',    fontSize=8,  textColor=WHITE,
                                  leading=11, wordWrap='LTR'),
    'td':         S('td',         fontName='Helvetica',         fontSize=8,  textColor=BLACK,
                                  leading=11, wordWrap='LTR'),
    'td_bold':    S('td_bold',    fontName='Helvetica-Bold',    fontSize=8,  textColor=BLACK,
                                  leading=11, wordWrap='LTR'),
    'td_green':   S('td_green',   fontName='Helvetica-Bold',    fontSize=8,  textColor=GREEN,
                                  leading=11, wordWrap='LTR'),
    'td_orange':  S('td_orange',  fontName='Helvetica-Bold',    fontSize=8,  textColor=ORANGE,
                                  leading=11, wordWrap='LTR'),
    'td_red':     S('td_red',     fontName='Helvetica-Bold',    fontSize=8,  textColor=RED,
                                  leading=11, wordWrap='LTR'),
}

# ── Helpers ────────────────────────────────────────────────────────────────────
def P(text, style='body'):   return Paragraph(text, styles[style])
def TH(text):  return Paragraph(text, styles['th'])
def TD(text):  return Paragraph(text, styles['td'])
def TDB(text): return Paragraph(text, styles['td_bold'])
def TDG(text): return Paragraph(text, styles['td_green'])
def TDO(text): return Paragraph(text, styles['td_orange'])
def TDR(text): return Paragraph(text, styles['td_red'])
def sp(n=6):   return Spacer(1, n)
def hr():      return HRFlowable(width='100%', thickness=1,   color=BLUE_DARK,  spaceAfter=6, spaceBefore=6)
def shr():     return HRFlowable(width='100%', thickness=0.5, color=BORDER_GRAY, spaceAfter=4, spaceBefore=4)

def tbl(data, col_widths, header_bg=BLUE_DARK, stripe=True):
    """Word-wrapped table. data[0] must be header row of Paragraph objects."""
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

def callout(text, width):
    data = [[Paragraph(text, ParagraphStyle('cb', fontName='Helvetica', fontSize=9,
                                            textColor=BLUE_DARK, leading=14, wordWrap='LTR'))]]
    t = Table(data, colWidths=[width])
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0), (-1,-1), BLUE_LIGHT),
        ('BOX',          (0,0), (-1,-1), 1.5, BLUE_MID),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING',   (0,0), (-1,-1), 8),
        ('BOTTOMPADDING',(0,0), (-1,-1), 8),
    ]))
    return t

# ── Document ───────────────────────────────────────────────────────────────────
def build():
    W = 7.5*inch

    doc = SimpleDocTemplate(
        '/Users/joshuagarrison/lsc/Auto-Wip/Status-Reports/Auto-WIP_Status_Report_20260401.pdf',
        pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch,
        title='Auto-WIP Schedule — Project Status Report',
        author='Josh Garrison, IT Development',
    )

    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    story += [
        P('Auto-WIP Schedule — Project Status Report', 'title'),
        P('Date: April 1, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Priority: HIGHEST — Owner / Board / Senior Management escalation</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # ── Executive Summary ─────────────────────────────────────────────────────
    story += [
        P('Executive Summary', 'h1'),
        P(
            'The Automated WIP Schedule rebuild has accelerated significantly in the past 24 hours. '
            'The LylesWIP database on the P&amp;P server is live. All 40 months of historical override data '
            '(WML complete; AIC, APC, NESM in progress) are loaded and merging correctly onto the workbook '
            'at data load time. The Vista read path is producing the correct job set with correct override values. '
            'Current workbook version: Rev 5.63p / 5.64p.'
        ),
        P(
            'The remaining work is concentrated in one layer: <b>the write path</b>. When a user edits an override '
            'column and double-clicks Done today, nothing persists. That is the primary build task ahead. '
            'The database, the stored procedures, the override load, and the read merge are all done — '
            'the plumbing exists. The VBA SaveJobRow wiring is next.'
        ),
        sp(4),
        callout(
            '<b>Updated production readiness: approximately 50–55%.</b> Significant progress from 35–40% on March 31. '
            'The database layer and override data load are complete. The write path (save overrides, '
            'batch workflow, permissions) remains the open body of work.<br/><br/>'
            '<b>Target delivery: May 8, 2026 — no change to schedule.</b>',
            W
        ),
        sp(8), shr(),
    ]

    # ── What Changed Since March 31 ───────────────────────────────────────────
    story += [P('What Changed Since March 31, 2026', 'h1')]

    delta_data = [
        [TH('Item'), TH('Status as of Mar 31'), TH('Status Today (Apr 1)')],
        [TD('LylesWIP database on P&amp;P server'),
         TDR('Not built'), TDG('Live — tables, stored procs, SQL login all configured')],
        [TD('Historical override files (40 months, WML)'),
         TDR('Not loaded'), TDG('4,974 rows loaded — Dec 2024 through Dec 2025')],
        [TD('Override merge on data load'),
         TDR('Not built'), TDG('Working — LylesWIPData.bas deployed; overrides apply over Vista data at load time')],
        [TD('Prior Projected Profit (Col Q / Col R)'),
         TDO('Formula patch only — no DB source'),
         TDG('Now loads prior month\'s OpsRevOverride − OpsCostOverride from LylesWIP')],
        [TD('Col R carry-forward rule confirmed'),
         TDO('Open question'), TDG('Confirmed with Nicole: Col R = prior month Col P. Rule implemented.')],
        [TD('Vista query: future-job filter (StartMonth)'),
         TDR('Missing from VBA'), TDG('Fixed — c.StartMonth &lt;= @CutOffDate added')],
        [TD('Vista query: job inclusion for closed jobs'),
         TDR('Too broad — all Soft Closed regardless of year'),
         TDG('Fixed — Soft Closed only if closed in current year; Hard Closed with MonthClosed &gt;= batch month')],
        [TD('Vista query: billing-only closed jobs (bARTH)'),
         TDR('Missing from VBA'), TDG('Fixed — bARTH EXISTS clause added; catches retainage / late billings')],
        [TD('IIf(IsNull) Null trap in MergePriorMonth'),
         TDR('Bug — crashed on Null override fields'), TDG('Fixed — replaced with proper If/Else')],
        [TD('Clear Workbook button'),
         TDR('Crashing with 1004 errors'), TDG('Fixed — QuickReset wired; all 6 sheets clear cleanly')],
        [TD('Compile errors (ByRef mismatch, missing subs)'),
         TDR('5 compile errors blocking VBA editor'), TDG('All resolved')],
    ]
    story += [tbl(delta_data, [2.3*inch, 1.6*inch, W - 3.9*inch]), sp(4), shr()]

    # ── Validation Status ─────────────────────────────────────────────────────
    story += [
        P('D2 Validation — WML December 2025 (In Progress)', 'h1'),
        P(
            'Validation against Nicole\'s WML December 2025 WIP History Import file is underway. '
            'The import file contains 28 jobs for Division 51. The Vista query is being compared '
            'job-by-job. Rev 5.64p (pending deployment) addresses the remaining job-count discrepancy.'
        ),
    ]

    val_data = [
        [TH('Check'), TH('Nicole\'s File'), TH('Our WIP (5.63p)'), TH('Status')],
        [TD('Open job count — Div 51'),
         TD('18 jobs'), TD('18 jobs'), TDG('Match')],
        [TD('Closed job count — Div 51'),
         TD('10 jobs'), TD('17 jobs (7 extra)'),
         TDO('Fix in 5.64p — Soft Closed year filter + bARTH billing filter')],
        [TD('Job 51.1166 / 51.1167 / 51.1168 (2026 start month)'),
         TD('Not present'), TD('Were appearing'),
         TDG('Fixed in 5.63p — StartMonth filter')],
        [TD('Job 51.1074 (Hard Closed 2016)'),
         TD('Not present'), TD('Was appearing'),
         TDG('Fixed in prior sprint — JobStatus IN (1,2)')],
        [TD('Col Q — Prior Projected Profit (51.1108)'),
         TD('$858,269 expected'), TD('Now loading from LylesWIP'),
         TDG('Confirmed working after IIf fix')],
        [TD('Override values (51.1108 Ops Rev)'),
         TD('$94,196,098'), TD('Loads from LylesWIP override'),
         TDG('Confirmed — matches Nicole\'s file')],
        [TD('AIC, APC, NESM companies'),
         TD('Pending'), TD('Pending'),
         TDO('D3 — after WML validation complete')],
    ]
    story += [sp(4), tbl(val_data, [2.4*inch, 1.1*inch, 1.5*inch, W - 5.0*inch]), sp(4), shr()]

    # ── Architecture ──────────────────────────────────────────────────────────
    story += [
        P('System Architecture — Three-Stage Workflow', 'h1'),
        P(
            'The LylesWIP database on the P&amp;P server is the central persistence layer. '
            'Vista is read-only. All users — including remote offices and job trailers — '
            'connect to P&amp;P only for the write path.'
        ),
    ]

    arch_data = [
        [TH('Stage'), TH('Who'), TH('Access'), TH('What Happens'), TH('DB Status')],
        [TDB('Stage 1\nAccounting\nInitial Review'),
         TD('Nicole Leasure'),
         TD('Vista Production\n(10.112.11.8)'),
         TD('Reads live Vista data. Batch created in LylesWIP. Snapshot stored. '
            '"Ready for Ops: Yes" advances batch state.'),
         TDO('Read works.\nWrite (CreateBatch) not yet wired.')],
        [TDB('Stage 2\nOperations\nReview'),
         TD('Ops Controllers\n(any location)'),
         TD('P&amp;P Server only\n(no Vista)'),
         TD('Reads batch from LylesWIP. Ops edits yellow override columns. '
            'Double-click col H saves row. "Ops Final Approval: Yes" advances state.'),
         TDR('Not yet wired.\nSaveJobRow built;\nVBA call missing.')],
        [TDB('Stage 3\nAccounting\nFinal Approval'),
         TD('Cindy Jordan /\nNicole Leasure'),
         TD('P&amp;P Server only'),
         TD('Reads LylesWIP — all Ops edits visible. GAAP override columns editable. '
            '"Accounting Final Approval: Yes" locks batch.'),
         TDR('Not yet wired.')],
    ]
    story += [sp(4), tbl(arch_data, [0.85*inch, 0.9*inch, 0.95*inch, 2.4*inch, W - 5.1*inch]), sp(4), shr()]

    # ── Production Readiness ──────────────────────────────────────────────────
    story += [P('Production Readiness Assessment — April 1, 2026', 'h1')]

    pr_data = [
        [TH('Layer'), TH('Mar 31'), TH('Apr 1'), TH('Notes')],
        [TD('Vista read path — data loads'),
         TD('~65%'), TDG('~82%'),
         TD('Job list now correct; billing filter added; 2 minor inclusion rules remain (A7, A8)')],
        [TD('GAAP formula accuracy'),
         TD('~70%'), TDG('~80%'),
         TD('Col W and Col R both working with live override data; cascade columns validated for 51.1108')],
        [TD('Override data load — WML'),
         TD('0%'),  TDG('~85%'),
         TD('4,974 rows loaded, merging correctly; AIC/APC/NESM pending')],
        [TD('Override data load — Other companies'),
         TD('0%'),  TDO('0%'),
         TD('WML validation must complete first; then expand')],
        [TD('Write path — saving overrides'),
         TD('0%'),  TDR('0%'),
         TD('Not yet wired; database and stored procs are ready')],
        [TD('Three-stage workflow state machine'),
         TD('15%'), TDO('15%'),
         TD('UI exists; workflow state not wired to LylesWIP')],
        [TD('Permissions — role-based'),
         TD('10%'), TDO('10%'),
         TD('Hardcoded to Accounting; P&amp;P proc ready to wire')],
        [TDB('Overall system'),
         TDB('~35–40%'), TDG('~50–55%'),
         TDB('Database layer complete; write-path wiring is the remaining critical path')],
    ]
    story += [tbl(pr_data, [2.1*inch, 0.65*inch, 0.65*inch, W - 3.4*inch]), sp(4), shr()]

    # ── Nicole's Discrepancies ─────────────────────────────────────────────────
    story += [P("Nicole's Reported Discrepancies — Current Status", 'h1')]

    disc_data = [
        [TH("Nicole's Issue"), TH('Mar 31 Status'), TH('Apr 1 Status')],
        [TD('Col W (JTD Earned Revenue) — wrong for &lt;10% jobs'),
         TDG('Fixed in workbook formula'), TDG('Fixed — confirmed working')],
        [TD('Col R (Prior Projected Profit) — showing $0 instead of prior WIP value'),
         TDO('Partial — formula patch only'),
         TDG('Fixed — now loads prior month\'s (OpsRev − OpsCost) from LylesWIP')],
        [TD('Cascade columns AD, AF, AH, AJ, AL, AQ, AR — all wrong'),
         TDO('Expected to resolve with Col W fix — not yet verified'),
         TDO('Not yet verified with full override data loaded for all companies')],
        [TD('Job 54.9033 — Closed in workbook but Open in Vista'),
         TDO('Not fixed — requires close/status reconciliation'),
         TDO('Not yet built — A8 in task list')],
        [TD('Jobs 56.1022 / 56.1057 — Missing despite needing to show reversals'),
         TDO('Not fixed — requires zero-JTD activity detection'),
         TDO('Not yet built — A7 in task list')],
    ]
    story += [tbl(disc_data, [2.4*inch, 1.9*inch, W - 4.3*inch]), sp(4), shr()]

    # ── Task List ─────────────────────────────────────────────────────────────
    story += [P('Task List', 'h1'), P('Phase A — Vista Query', 'h2')]

    a_data = [
        [TH('#'), TH('Task'), TH('Detail'), TH('Status')],
        [TD('A1'), TDB('Date parameter confirmed'),
         TD('Nicole confirmed: Ending Month = batch month; Beginning Month = blank. '
            'All 7 CTEs cut off at batch month-end using @CutOffDate.'),
         TDG('DONE')],
        [TD('A2'), TDB('Date cutoff on all CTEs'),
         TD('Cost and billing queries cut off at batch month-end, not today\'s date.'),
         TDG('DONE')],
        [TD('A3'), TDB('Future-job filter'),
         TD('c.StartMonth &lt;= @CutOffDate added to JobList CTE. 51.1166/1167/1168 (2026 start) now excluded.'),
         TDG('DONE')],
        [TD('A4'), TDB('Billing activity — retainage / late invoices'),
         TD('bARTH EXISTS clause added to JobList. Catches jobs with billing but no cost activity (e.g. 51.1136).'),
         TDG('DONE')],
        [TD('A5'), TDB('Open zero-activity jobs'),
         TD('j.JobStatus = 1 always included regardless of cost activity.'),
         TDG('DONE')],
        [TD('A6'), TDB('Closed job year filter'),
         TD('Soft Closed (status 2): only if MonthClosed &gt;= @StartDate (Jan 1 of WIP year). '
            'Hard Closed (status 3): only if MonthClosed &gt;= @Month. Eliminates 2024-closed jobs.'),
         TDG('DONE — 5.64p')],
        [TD('A7'), TDB('Zero-JTD reversal jobs'),
         TD('Detect current-period activity when JTD nets to $0 (jobs 56.1022 / 56.1057 type).'),
         TDR('Pending')],
        [TD('A8'), TDB('Close status reconciliation'),
         TD('Compare workbook Close flag against Vista Contract Status; surface mismatch.'),
         TDR('Pending')],
    ]
    story += [tbl(a_data, [0.35*inch, 1.55*inch, W - 3.0*inch, 0.7*inch]), sp(8)]

    story += [P('Phase B — Database Build (P&P Server)', 'h2')]

    b_data = [
        [TH('#'), TH('Task'), TH('Detail'), TH('Status')],
        [TD('B1'), TDB('LylesWIP database on P&amp;P server'),
         TD('Database created; workbook Settings sheet (C5/C6) configured.'),
         TDG('DONE')],
        [TD('B2'), TDB('WipBatches table'),
         TD('Per company/month/division batch tracking with state machine columns.'),
         TDG('DONE')],
        [TD('B3'), TDB('WipJobData table'),
         TD('Per-job override storage: GAAP rev, Ops rev, GAAP cost, Ops cost, bonus, notes, dates, flags, approval state.'),
         TDG('DONE')],
        [TD('B4'), TDB('WipMonthlySnapshot table'),
         TD('Month-end archive; source for Prior Year columns next cycle.'),
         TDG('DONE')],
        [TD('B5'), TDB('Stored procedures'),
         TD('CreateBatch, SaveJobRow, GetJobOverrides (BuildOverrideLookup), UpdateBatchState, CheckBatchState — all live.'),
         TDG('DONE')],
        [TD('B6'), TDB('SQL login permissions'),
         TD('P&amp;P SQL auth login granted execute on all WIP procs.'),
         TDG('DONE')],
        [TD('B7'), TDB('Load historical overrides — WML'),
         TD('4,974 rows loaded from 40 monthly import files (Dec 2024–Dec 2025, WML Co 15).'),
         TDG('DONE')],
        [TD('B8'), TDB('Load historical overrides — AIC, APC, NESM'),
         TD('Same import process; pending after WML validation is confirmed.'),
         TDO('Pending — after D2')],
    ]
    story += [tbl(b_data, [0.35*inch, 1.55*inch, W - 3.0*inch, 0.7*inch]), sp(8)]

    story += [P('Phase C — Workbook Write Path (VBA)', 'h2')]

    c_data = [
        [TH('#'), TH('Task'), TH('Detail'), TH('Status')],
        [TD('C1'), TDB('LylesWIPData.bas module'),
         TD('OpenWIPConnection, CloseWIPConnection, MergeOverridesOntoSheet, '
            'MergePriorMonthProfitsOntoSheet, BuildOverrideLookup — all live.'),
         TDG('DONE')],
        [TD('C2'), TDB('Batch creation on Load'),
         TD('CreateBatch called on data load; batch record created in WipBatches.'),
         TDG('DONE')],
        [TD('C3'), TDB('Override merge on display'),
         TD('MergeOverridesOntoSheet applies LylesWIP override values over raw Vista data at load time. '
            'MergePriorMonthProfitsOntoSheet loads prior month Col Q.'),
         TDG('DONE')],
        [TD('C4'), TDB('Wire UpdateRow — Ops Done'),
         TD('Double-click col H calls SaveJobRow → writes all Ops override fields for that row to WipJobData.'),
         TDR('Next up')],
        [TD('C5'), TDB('Wire UpdateRow — GAAP Done'),
         TD('Double-click col I saves GAAP override fields to WipJobData.'),
         TDR('Pending')],
        [TD('C6'), TDB('UseExistingBatch / CreateBatch flow'),
         TD('On batch open: check for existing batch → prompt reopen or create new.'),
         TDR('Pending')],
        [TD('C7'), TDB('Start sheet state transitions'),
         TD('"Ready for Ops: Yes" → ReadyForOps. "Ops Final Approval: Yes" → OpsApproved. '
            '"Accounting Final Approval: Yes" → AcctApproved.'),
         TDR('Pending')],
        [TD('C8'), TDB('CompleteCheck before state advance'),
         TD('Verify all jobs have IsOpsDone=1 before advancing to OpsApproved; IsGAAPDone=1 before AcctApproved.'),
         TDR('Pending')],
        [TD('C9'), TDB('Cell locking on state advance'),
         TD('Jobs-Ops yellow cells lock at OpsApproved. Jobs-GAAP yellow cells lock at AcctApproved.'),
         TDR('Pending')],
        [TD('C10'), TDB('Role-based permissions'),
         TD('Replace hardcoded WIPAccounting with pnp.WIPSECGetRole call. Show/hide UI per role.'),
         TDR('Pending')],
        [TD('C11'), TDB('[Phase 2] Vista write-back'),
         TD('At AcctApproved, write final GAAP projections to Vista bJCOR / bJCOP. Deferred — not blocking delivery.'),
         TDO('Deferred')],
    ]
    story += [tbl(c_data, [0.35*inch, 1.55*inch, W - 3.0*inch, 0.7*inch]), sp(8)]

    story += [P('Phase D — Validation and Delivery', 'h2')]

    d_data = [
        [TH('#'), TH('Task'), TH('Detail'), TH('Status')],
        [TD('D1'), TDB('Confirm Col R carry-forward rule'),
         TD('Confirmed with Nicole: Col R = prior month Col P. '
            'Implemented as: prior month\'s (OpsRevOverride − OpsCostOverride) from LylesWIP.'),
         TDG('DONE')],
        [TD('D2'), TDB('WML Dec 2025 validation'),
         TD('Open job count matches (18). Closed job fix in 5.64p. '
            'Column values being verified against Nicole\'s import file.'),
         TDO('In Progress')],
        [TD('D3'), TDB('AIC, APC, NESM validation'),
         TD('Load override data for remaining 3 companies; repeat column validation.'),
         TDR('Pending — after D2')],
        [TD('D4'), TDB('SQL driver — Brian Platten'),
         TD('ODBC SQL Server driver install required for P&amp;P connection.'),
         TDR('Pending')],
        [TD('D5'), TDB('SQL driver — Harbir Atwal'),
         TD('Same install required.'),
         TDR('Pending')],
        [TD('D6'), TDB('User setup instructions'),
         TD('Trusted folder path, SQL driver install steps, first-open checklist.'),
         TDR('Pending')],
        [TD('D7'), TDB('Security cleanup'),
         TD('Re-hide Settings sheet (xlSheetVeryHidden). Clear test credentials. '
            'Verify no hardcoded usernames in VBA.'),
         TDR('Pending')],
        [TD('D8'), TDB('Final sign-off'),
         TD('Nicole and Cindy confirm numbers tie and complete three-stage workflow end-to-end.'),
         TDR('Pending')],
    ]
    story += [tbl(d_data, [0.35*inch, 1.55*inch, W - 3.0*inch, 0.7*inch]), sp(4), shr()]

    # ── Milestones ────────────────────────────────────────────────────────────
    story += [P('Milestone Schedule', 'h1')]

    ms_data = [
        [TH('Milestone'), TH('Target'), TH('Status')],
        [TD('Vista read path complete — all 4 sheets load from production'),
         TD('Mar 5, 2026'), TDG('DONE')],
        [TD('Formula fixes: Column W and Column R'),
         TD('Mar 2026'), TDG('DONE')],
        [TD('Live demo with Nicole + Cindy — bugs documented'),
         TD('Mar 25, 2026'), TDG('DONE')],
        [TD('Historical override files located and structure confirmed'),
         TD('Mar 31, 2026'), TDG('DONE')],
        [TD('LylesWIP database live on P&amp;P — tables, stored procs, SQL login'),
         TD('Apr 14, 2026'), TDG('DONE — Apr 1 (ahead of schedule)')],
        [TD('WML historical override data loaded into LylesWIP (4,974 rows)'),
         TD('Apr 17, 2026'), TDG('DONE — Apr 1 (ahead of schedule)')],
        [TD('All Vista query job-filter bugs fixed (A1–A6)'),
         TD('Apr 11, 2026'), TDG('DONE — Apr 1 (ahead of schedule)')],
        [TD('Remaining Vista query items (A7 zero-JTD reversals, A8 close reconciliation)'),
         TD('Apr 11, 2026'), TDO('In Progress')],
        [TD('WML December 2025 validation confirmed with Nicole'),
         TD('Apr 18, 2026'), TDO('In Progress — 5.64p under test')],
        [TD('Write path working — Ops overrides save, batches track (C4–C8)'),
         TD('Apr 24, 2026'), TD('')],
        [TD('AIC, APC, NESM override data loaded and validated (B8, D3)'),
         TD('Apr 25, 2026'), TD('')],
        [TD('Role-based permissions live (C10)'),
         TD('Apr 28, 2026'), TD('')],
        [TD('Three-stage workflow complete end-to-end (C7–C9)'),
         TD('May 1, 2026'), TD('')],
        [TD('Nicole validation session — all companies confirmed'),
         TD('May 5, 2026'), TD('')],
        [TD('Final fixes, security cleanup, user setup'),
         TD('May 7, 2026'), TD('')],
        [TDB('Delivery — signed off, all users set up, distributed'),
         TDB('May 8, 2026'), TD('')],
    ]
    story += [tbl(ms_data, [W - 2.15*inch, 1.15*inch, 1.0*inch]), sp(4), shr()]

    # ── Key Contacts ──────────────────────────────────────────────────────────
    story += [P('Key Contacts', 'h1')]

    contacts_data = [
        [TH('Name'), TH('Role'), TH('Notes')],
        [TDB('Kevin Shigematsu'), TD('CEO — project sponsor'), TD('')],
        [TDB('Cindy Jordan'),     TD('CFO — final approver'), TD('')],
        [TDB('Nicole Leasure'),   TD('VP Corporate Controller — primary validator'), TD('')],
        [TDB('Josh Garrison'),    TD('Director of Technology Innovation — developer'), TD('')],
        [TDB('Dane Wildey'),      TD('CIO'), TD('')],
        [TDB('Brian Platten'),    TD('Controller — maintains DJS; needs SQL driver install'), TD('Action: D4')],
        [TDB('Harbir Atwal'),     TD('Controller — needs SQL driver install'), TD('Action: D5')],
        [TDB('Michael Roberts'),  TD('Original consultant (RCS Plan)'),
         TD('Available for knowledge transfer if needed; not on critical path')],
    ]
    story += [tbl(contacts_data, [1.5*inch, 2.5*inch, W - 4.0*inch])]

    # ── Footer ────────────────────────────────────────────────────────────────
    story += [
        sp(12),
        HRFlowable(width='100%', thickness=0.5, color=BORDER_GRAY),
        sp(4),
        P('Generated: April 1, 2026  |  Confidential — Lyles Services Co. internal document', 'note'),
    ]

    doc.build(story)
    print('PDF written: Auto-WIP_Status_Report_20260401.pdf')

if __name__ == '__main__':
    build()
