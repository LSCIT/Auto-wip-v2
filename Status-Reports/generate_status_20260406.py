#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 6, 2026."""

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
def B(text):   return Paragraph(f'<bullet>&bull;</bullet> {text}', styles['bullet'])
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

def green_box(text, width):
    return callout(text, width,
                   bg=colors.HexColor('#E8F5E9'),
                   border=colors.HexColor('#1D6A3C'))

def warn_box(text, width):
    return callout(text, width,
                   bg=colors.HexColor('#FFF4E5'),
                   border=colors.HexColor('#B45309'))

def build():
    W = 7.5 * inch

    doc = SimpleDocTemplate(
        r'E:\Auto-Wip\Status-Reports\Auto-WIP_Status_Report_20260406.pdf',
        pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch,
        title='Auto-WIP Schedule — Project Status Report — April 6, 2026',
        author='Josh Garrison, IT Development',
    )

    story = []

    # == Header ================================================================
    story += [
        P('Auto-WIP Schedule — Project Status Report', 'title'),
        P('Date: April 6, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Priority: HIGHEST — Owner / Board / Senior Management escalation</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # == Executive Summary =====================================================
    story += [
        P('Executive Summary', 'h1'),
        P(
            'April 6 began with a demo session with Nicole Leasure and Cindy Jordan, validating the '
            'full 3-stage workflow on Rev 5.68p. The demo confirmed multiple values are correct '
            '(contract amounts, change in anticipated profit, override projections, workflow state '
            'machine) and identified five data accuracy issues — all traced to the same root cause: '
            'date filtering in the Vista query was using transaction posted dates instead of fiscal '
            'months, causing cost and billing values to diverge from the Crystal Report "JC Cost and '
            'Revenue" that Nicole validates against.'
        ),
        P(
            'All five issues were investigated, root-caused, and fixed the same day. The corrected '
            'query was verified against the Crystal Report with <b>exact penny-match</b> on both JTD '
            'Cost ($87,315,159.22) and Billed to Date ($96,918,206.90) for the benchmark job 51.1129. '
            'These fixes, along with several other improvements, were deployed as Rev 5.70p.'
        ),
        sp(4),
        green_box(
            '<b>Production readiness: approximately 88-92%.</b> '
            'JTD Cost and Billed to Date now match Crystal Report exactly. Prior month profit, '
            'job inclusion, and Vista write-back copy tables all addressed. Remaining: circular '
            'reference edge case (pre-existing), fine-tuning prior profit denominator, '
            'multi-company validation.<br/><br/>'
            '<b>Target delivery: May 8, 2026 — no change.</b>',
            W
        ),
        sp(8), shr(),
    ]

    # == Demo Session ==========================================================
    story += [
        P('Demo Session — April 6, 2026', 'h1'),
        P('<b>Attendees:</b> Josh Garrison, Nicole Leasure, Cindy Jordan', 'body'),
        P('<b>Workbook:</b> WIPSchedule - Rev 5.68p, Company 15 (WML), Division 51, December 2025', 'body'),
        P('<b>Duration:</b> ~36 minutes (recorded with transcript)', 'body'),
        sp(4),
        P('What Validated Successfully', 'h2'),
    ]

    validated_data = [
        [TH('Item'), TH('Detail')],
        [TD('Contract amounts'), TD('Confirmed correct across multiple jobs')],
        [TD('Change in anticipated profit'), TD("All values matched Nicole's WIP")],
        [TD('Override values (GAAP projections)'), TD('51.1129 contract amount 126,483,581 confirmed')],
        [TD('Ready for OPS button'), TD('Permission check working — blocked unauthorized roles')],
        [TD('GAAP Done (double-click save)'), TD('Saved to database with user name attached')],
        [TD('Audit trail comments'), TD('Override changes recorded with from/to/who/when')],
        [TD('Closed jobs filtering'), TD('Improved — fewer/correct jobs in closed section')],
    ]
    story += [tbl(validated_data, [2.5*inch, 5.0*inch]), sp(8)]

    story += [P('Issues Identified', 'h2')]

    issues_data = [
        [TH('#'), TH('Issue'), TH('Root Cause'), TH('Status')],
        [TD('1'),
         TD('JTD Cost pulling beyond batch month'),
         TD('Query used PostedDate instead of Mth (fiscal month)'),
         TDG('Fixed — exact match')],
        [TD('2'),
         TD('Billed to Date pulling beyond batch month'),
         TD('Query used bARTH (AR side) instead of JB Progress Bills'),
         TDG('Fixed — exact match')],
        [TD('3'),
         TD('Circular reference on job 51.1158'),
         TD("Pre-existing formula issue in Michael's workbook (zero-profit edge case)"),
         TDO('Confirmed not our bug')],
        [TD('4'),
         TD('Prior month profit showing projected instead of recognized'),
         TD('MergePriorMonthProfits wrote OpsRev-OpsCost instead of earned rev - JTD cost'),
         TDG('Fixed')],
        [TD('5'),
         TD('Job 54.9416 not appearing (StartMonth = Jan 2026)'),
         TD('Query excluded jobs by StartMonth only'),
         TDG('Fixed — cost-exists inclusion')],
        [TD('6'),
         TD('Completion Date source'),
         TD("Should come from Nicole's override files"),
         TDG('Already working')],
    ]
    story += [tbl(issues_data, [0.3*inch, 1.8*inch, 2.8*inch, 2.6*inch]), sp(6)]

    story += [
        P("Nicole's New Business Rule — Job Inclusion", 'h3'),
        P(
            "Nicole clarified that job inclusion should be driven by <b>cost activity first</b>, "
            "not just contract StartMonth:"
        ),
        B('If any cost has hit the job through the batch month — include it'),
        B('If no cost, fall back to StartMonth &le; batch month (for backlog)'),
        P(
            'This resolves job 54.9416, which had preliminary design costs in Nov/Dec 2025 but an '
            'official StartMonth of January 2026.',
            'note'
        ),
        sp(4), shr(),
    ]

    # == What Changed Since April 3 ============================================
    story += [P('What Changed Since April 3, 2026', 'h1')]

    delta_data = [
        [TH('Item'), TH('Status Apr 3'), TH('Status Today (Apr 6)')],
        [TD('JTD Cost accuracy'),
         TDO('Off by ~$36K vs Crystal Report (PostedDate filter)'),
         TDG('Exact match (Mth filter)')],
        [TD('Billed to Date accuracy'),
         TDO('Off by ~$236K vs Crystal Report (bARTH source)'),
         TDG('Exact match (JB Progress Bills)')],
        [TD('Job inclusion (54.9416)'),
         TDO('Excluded (StartMonth filter only)'),
         TDG('Included (cost-exists fallback)')],
        [TD('Prior month profit'),
         TDO('Showed projected ($11.25M)'),
         TDG('Shows recognized (~$8.3M)')],
        [TD('LylesWIP override data'),
         TDO('GAAP values missing, zeros stored as NULL'),
         TDG('All 4,973 rows reloaded, 0 mismatches')],
        [TD('DB state for demo'),
         TDO('Ops Done flags set from testing'),
         TDG('Reset — clean for Nicole/Cindy')],
        [TD('Vista write-back tables'),
         TDR('Not built'),
         TDG('WipJCOP + WipJCOR created, proc tested')],
        [TD('Push to Vista button'),
         TDR('Not built'),
         TDG('Built (test mode — guards only, no write)')],
        [TD('Workbook versions'),
         TD('5.68p only'),
         TDG('5.68p (demo), 5.70p (fixes), source separated')],
        [TD('Formula comparison'),
         TDR('Not verified'),
         TDG('22/22 Ops match, 21/22 GAAP match (1 intentional fix)')],
    ]
    story += [tbl(delta_data, [1.7*inch, 2.5*inch, 3.3*inch]), sp(8), shr()]

    # == Data Integrity Fix ====================================================
    story += [
        P('Data Integrity Fix — April 6, 2026', 'h1'),
        P('Before the demo, identified and fixed a data loading issue:'),
        B("<b>to_decimal() in load_overrides.py</b> — Nicole's explicit $0 overrides were stored as NULL "
          "instead of 0. Fixed: zeros now stored as 0 with Plugged=1."),
        B('<b>GAAP override values</b> — 541 GAAP overrides were missing from December 2025. Root cause: '
          'values were wiped during April 3 testing. Fixed: all 40 files reloaded (4,973 rows).'),
        B('<b>Test artifacts</b> — Reset Div 51 batch state, cleared 28 workflow flags, restored '
          "51.1108 override values to Nicole's originals."),
        B('<b>Verification:</b> 577 Dec 2025 rows compared against Nicole\'s source files — '
          '<b>0 mismatches</b>.'),
        sp(8), shr(),
    ]

    # == Technical Detail — Query Fixes ========================================
    story += [
        P('Technical Detail — Query Fixes', 'h1'),
        P('Fix 1: Fiscal Month Filtering (JTD Cost)', 'h2'),
        P(
            'The Crystal Report "JC Cost and Revenue" filters by Mth (fiscal month), not PostedDate '
            'or ActualDate. A transaction entered in January 2026 can have a fiscal month of December '
            '2025 (or vice versa). Switching all bJCCD date filters from PostedDate/ActualDate to Mth '
            'produces an exact match.'
        ),
    ]

    cost_data = [
        [TH('Source'), TH('JTD Cost for 51.1129')],
        [TD('Crystal Report (Ending Month 12/25)'), TDB('$87,315,159.22')],
        [TD('Our query (Mth <= @Month)'), TDG('$87,315,159.22')],
        [TD('Old query (PostedDate <= @CutOffDate)'), TDR('$87,279,469 (off by ~$36K)')],
    ]
    story += [tbl(cost_data, [3.5*inch, 4.0*inch]), sp(6)]

    story += [
        P('Fix 2: JB Progress Bills (Billed to Date)', 'h2'),
        P(
            'The Crystal Report sources "Billed Amount" from the JB (Job Billing) module via '
            'vrvJBProgressBills, not from AR (Accounts Receivable) via bARTH. AR tracks invoices/'
            'payments/retainage separately; JB tracks cumulative billed per progress bill.'
        ),
    ]

    billed_data = [
        [TH('Source'), TH('Billed Amount for 51.1129')],
        [TD('Crystal Report'), TDB('$96,918,206.90')],
        [TD('vrvJBProgressBills (JB side)'), TDG('$96,918,206.90')],
        [TD('bARTH Invoiced (AR side)'), TDR('$96,682,005.02 (off by ~$236K)')],
    ]
    story += [tbl(billed_data, [3.5*inch, 4.0*inch]), sp(6)]

    story += [
        P('Fix 3: Cost-Exists Job Inclusion', 'h2'),
        P(
            'Added a secondary inclusion path: if a job has actual cost activity (bJCCD) through '
            'the batch month, include it even if StartMonth is after the batch month. Adds 2 jobs '
            'for Dec 2025: 54.9416 (Div 54) and 57.0009 (Div 57).'
        ),
        P('Fix 4: Recognized Profit (Prior Month)', 'h2'),
        P(
            'Changed MergePriorMonthProfitsOntoSheet to compute recognized profit (earned revenue '
            '- JTD actual cost) instead of projected profit (override rev - override cost). Added '
            'BuildPriorMonthCostLookup function that queries Vista for prior month JTD costs.'
        ),
        sp(4),
        P('Circular Reference (51.1158) — Not Our Bug', 'h2'),
        P(
            "Formula comparison between Michael's 5.47p and our 5.70p confirmed: <b>zero formula "
            "differences on Jobs-Ops</b> (22/22 match), <b>one intentional fix on Jobs-GAAP</b> "
            "(Col Z — removed erroneous 10% threshold, approved by Nicole in Sprint 1). The circular "
            "reference is pre-existing and occurs when OpsRev = OpsCost (zero expected profit)."
        ),
        sp(8), shr(),
    ]

    # == Current Workbook Versions =============================================
    story += [P('Current Workbook Versions', 'h1')]

    version_data = [
        [TH('Version'), TH('Purpose'), TH('Status')],
        [TDB('5.68p'), TD('Nicole/Cindy demo version'), TDG('Frozen — validated, DB reset for testing')],
        [TDB('5.70p'), TD('Query fixes + Push to Vista button'), TDG('Built — Crystal Report match verified')],
    ]
    story += [
        tbl(version_data, [1.0*inch, 3.25*inch, 3.25*inch]),
        sp(4),
        P('Source code separated: vba_source/ (5.68p frozen) and vba_source_569p/ (5.70p development).', 'note'),
        sp(8), shr(),
    ]

    # == Vista Write-Back Infrastructure =======================================
    story += [
        P('Vista Write-Back Infrastructure (Phase 2 Prep)', 'h1'),
        P('Built local copies of Vista override tables in LylesWIP for safe testing:'),
    ]

    writeback_data = [
        [TH('Table'), TH('Mirrors'), TH('Purpose')],
        [TD('WipJCOP'), TD('Vista bJCOP'), TD('GAAP/OPS cost overrides')],
        [TD('WipJCOR'), TD('Vista bJCOR'), TD('GAAP/OPS revenue overrides')],
    ]
    story += [tbl(writeback_data, [1.5*inch, 2.0*inch, 4.0*inch]), sp(6)]

    story += [
        P('Stored procedure <b>LylesWIPWriteBackToVista</b> created with:', 'body'),
        B("Two-pass MERGE pattern (matches Michael's LCGWIPMergeDetail)"),
        B('Quarterly guard (Mar/Jun/Sep/Dec only)'),
        B('AcctApproved guard (all departments must be approved)'),
        B('Smoke-tested: 51.1108 values match Nicole\'s source exactly'),
        sp(4),
        P(
            '"Push to Vista" button on Start sheet — currently in test mode (validates all guards, '
            'shows confirmation, does not execute write).',
            'note'
        ),
        sp(8), shr(),
    ]

    # == Remaining for Delivery ================================================
    story += [P('Remaining for Delivery', 'h1')]

    remain_data = [
        [TH('#'), TH('Item'), TH('Priority'), TH('Status')],
        [TD('R1'), TD('Nicole/Cindy validation on 5.70p'), TDR('Critical'), TDO('Ready — fixes address all demo findings')],
        [TD('R2'), TD('Multi-company testing (AIC, APC, NESM)'), TDR('High'), TDO('Not started')],
        [TD('R3'), TD('Wire permissions to pnp.WIPSECGetRole'), TDO('Medium'), TDO('Not started')],
        [TD('R4'), TD('SQL driver install for Brian Platten + Harbir Atwal'), TDO('Medium'), TDO('Not started')],
        [TD('R5'), TD('Security cleanup (Settings sheet, test credentials)'), TDR('High'), TDO('Not started')],
        [TD('R6'), TD('Circular reference on 51.1158 (pre-existing)'), TD('Low'), TD("Pre-existing — won't block delivery")],
        [TD('R7'), TD('Fine-tune prior profit denominator (~$312K gap)'), TDO('Medium'), TD('Close — may resolve with full data reload')],
        [TD('R8'), TD('Confirm: decimal display preference (cents vs whole dollars)'), TD('Low'), TD('Question for Nicole')],
        [TD('R9'), TD('Confirm: GAAP Billed to Date source (same Crystal Report?)'), TD('Low'), TD('Question for Nicole')],
    ]
    story += [tbl(remain_data, [0.4*inch, 3.3*inch, 0.8*inch, 3.0*inch]), sp(8), shr()]

    # == Milestone Schedule ====================================================
    story += [P('Milestone Schedule', 'h1')]

    miles_data = [
        [TH('Milestone'), TH('Target'), TH('Status')],
        [TD('Vista read path — all 4 sheets'),
         TD('Mar 28'), TDG('Done')],
        [TD('Vista query bugs fixed (A1-A8)'),
         TD('Apr 11'), TDG('Done Mar 31')],
        [TD('LylesWIP database + stored procs'),
         TD('Apr 14'), TDG('Done Mar 31')],
        [TD('Override load (40 files, 4,974 rows)'),
         TD('Apr 17'), TDG('Done Mar 31')],
        [TD('Validation pipeline (24/24 PASS)'),
         TD('Apr 4'), TDG('Done Apr 3')],
        [TD('Write path + 3-stage workflow'),
         TD('Apr 14'), TDG('Done Apr 3')],
        [TD('Data integrity fix + full reload'),
         TD('Apr 6'), TDG('Done Apr 6')],
        [TDB('Query fixes — Crystal Report exact match'),
         TDB('Apr 7'), TDG('Done Apr 6')],
        [TDB('Vista write-back copy tables'),
         TDB('Apr 7'), TDG('Done Apr 6')],
        [TD('Nicole/Cindy validation (5.70p)'),
         TD('Apr 7-11'), TDO('Ready to schedule')],
        [TD('Multi-company testing'),
         TD('Apr 14-18'), TDO('Not started')],
        [TD('Permissions wire (P&amp;P proc)'),
         TD('Apr 21-25'), TDO('Not started')],
        [TD('Security cleanup'),
         TD('Apr 28 - May 2'), TDO('Not started')],
        [TDB('Production delivery'),
         TDB('May 8'), TDB('On track')],
    ]
    story += [tbl(miles_data, [3.0*inch, 1.2*inch, 3.3*inch]), sp(8), shr()]

    # == Key Contacts ==========================================================
    story += [P('Key Contacts', 'h1')]

    contacts_data = [
        [TH('Name'), TH('Role'), TH('Relevance')],
        [TD('Kevin Shigematsu'), TD('CEO / Project Sponsor'), TD('Board commitment. Final escalation.')],
        [TD('Cindy Jordan'), TD('CFO'), TD('Final sign-off on WIP schedule accuracy.')],
        [TD('Nicole Leasure'), TD('VP Corporate Controller'), TD('Primary user and validator. Source of truth for override data.')],
        [TD('Dane Wildey'), TD('CIO'), TD('IT oversight, infrastructure.')],
        [TD('Brian Platten'), TD('Controller'), TD('Reviewer. Needs SQL driver install.')],
        [TD('Harbir Atwal'), TD('Controller'), TD('Reviewer. Needs SQL driver install.')],
        [TD('Josh Garrison'), TD('Dir. of Technology Innovation'), TD('Developer. Report author.')],
    ]
    story += [tbl(contacts_data, [1.5*inch, 2.0*inch, 4.0*inch]), sp(8)]

    # == Footer ================================================================
    story += [
        shr(), sp(4),
        P('Auto-WIP Status Report &nbsp;|&nbsp; April 6, 2026 &nbsp;|&nbsp; '
          'CONFIDENTIAL — Lyles Services Co. internal use only', 'footer'),
    ]

    doc.build(story)
    print('PDF saved: Auto-WIP_Status_Report_20260406.pdf')


if __name__ == '__main__':
    build()
