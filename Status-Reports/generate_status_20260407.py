#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 7, 2026."""

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

def build():
    W = 7.5 * inch

    doc = SimpleDocTemplate(
        r'E:\Auto-Wip\Status-Reports\Auto-WIP_Status_Report_20260407.pdf',
        pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch,
        title='Auto-WIP Schedule — Project Status Report — April 7, 2026',
        author='Josh Garrison, IT Development',
    )

    story = []

    # == Header ================================================================
    story += [
        P('Auto-WIP Schedule — Project Status Report', 'title'),
        P('Date: April 7, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Priority: HIGHEST — Owner / Board / Senior Management escalation</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # == Executive Summary =====================================================
    story += [
        P('Executive Summary', 'h1'),
        P(
            'Received first round of user feedback from Nicole Leasure on Rev 5.70p. Nicole identified '
            'three data discrepancies on the Jobs-Ops tab for job 51.1129, all traced to the same root '
            'cause: VBA code was overwriting sheet formula cells with computed values, bypassing the '
            "original template's architecture where visible columns read from hidden Z-columns populated "
            'by LylesWIP. Two column label changes were also requested. All five items were investigated, '
            'root-caused, and fixed in Rev 5.71p the same day.'
        ),
        sp(4),
        green_box(
            '<b>Production readiness: approximately 90-93%.</b> '
            'All Nicole feedback items from 5.70p addressed. Label changes applied. '
            'VBA modules imported and workbook ready for re-test.<br/><br/>'
            '<b>Target delivery: May 8, 2026 — no change.</b>',
            W
        ),
        sp(8), shr(),
    ]

    # == Nicole's Feedback =====================================================
    story += [
        P("Nicole's Feedback — April 7, 2026", 'h1'),
        P('<b>Source:</b> Email from Nicole Leasure, reviewing Rev 5.70p', 'body'),
        P('<b>Job referenced:</b> 51.1129, Company 15 (WML), Division 51', 'body'),
        sp(4),
        P('Issues Reported and Resolved', 'h2'),
    ]

    issues_data = [
        [TH('#'), TH("Nicole's Report"), TH('Column'), TH('Root Cause'), TH('Fix')],
        [TD('1'),
         TD('AC should be $8,997,122 — "seems to be pulling from GAAP schedule"'),
         TD('COLJTDPriorProfit (AC)'),
         TD('MergePriorMonthProfitsOntoSheet was overwriting the AC formula with a computed '
            '"recognized profit." The template formula =BK (hidden Z-column) was correct.'),
         TDG('Rewrote to populate Z-column BK with prior month BonusProfit from LylesWIP')],
        [TD('2'),
         TD("AG should be $71,707,694 — Michael's spreadsheet correct for 06/01/25"),
         TD('COLAPYRev (AG)'),
         TD('PY Revenue = PYCost + PriorYrBonusProfit, but Vista stubs bonus to 0. '
            'Backfill function wrote bonus to AJ but never corrected AG.'),
         TDG('Expanded MergePriorYearBonusOntoSheet to also update AG = PYCost + bonus')],
        [TD('3'),
         TD('AI should be $6,957,477'),
         TD('COLAPYCalcProfit (AI)'),
         TD('Template has formula =AG-AH but VBA overwrote it with GAAP-based values.'),
         TDG('Removed VBA overwrite; formula auto-calculates')],
        [TD('4'),
         TD('Rename AD on Jobs-Ops'),
         TD('COLJTDChgProfit (AD)'),
         TD('Label change request'),
         TDG('"MTD Change in Profit"')],
        [TD('5'),
         TD('Rename AA on Jobs-GAAP'),
         TD('COLJTDChgProfit (AA)'),
         TD('Label change request'),
         TDG('"MTD Change in Profit"')],
    ]
    story += [tbl(issues_data, [0.3*inch, 1.5*inch, 1.1*inch, 2.2*inch, 2.4*inch]), sp(8)]

    # == Root Cause Pattern ====================================================
    story += [
        P('Root Cause Pattern', 'h2'),
        P(
            "All three data issues share the same architectural root cause: <b>VBA code was writing "
            "directly to visible formula columns, destroying the template formulas that read from "
            "hidden Z-columns.</b>"
        ),
        P(
            "Michael's original template architecture uses a two-layer pattern:"
        ),
        B('<b>Hidden Z-columns</b> (BK, BN, etc.) are populated by VBA with data from LylesWIP'),
        B('<b>Visible columns</b> (AC, Q, AI, etc.) contain formulas that read from the Z-columns'),
        sp(4),
        P(
            'The Vista direct-read query (added in Sprint 1) stubs all LylesWIP-sourced fields to 0. '
            'The merge functions (MergePriorMonthProfitsOntoSheet, MergePriorYearBonusOntoSheet) were '
            'created to backfill these from LylesWIP after the Vista data loads. The bug was that these '
            'functions wrote to the visible columns instead of (or in addition to) the Z-columns, '
            'overwriting the formulas.',
            'note'
        ),
        sp(8), shr(),
    ]

    # == What Changed Since April 6 ============================================
    story += [P('What Changed Since April 6, 2026', 'h1')]

    delta_data = [
        [TH('Item'), TH('Status Apr 6'), TH('Status Today (Apr 7)')],
        [TD('Nicole/Cindy feedback'),
         TDO('Email sent, waiting'),
         TDG('Received — 5 items, all resolved')],
        [TD('COLJTDPriorProfit (AC)'),
         TDO('Showed computed recognized profit'),
         TDG('Shows prior month BonusProfit via Z-column')],
        [TD('COLAPYRev (AG)'),
         TDO('Missing bonus (Vista stub = 0)'),
         TDG('PYCost + bonus from LylesWIP year-end')],
        [TD('COLAPYCalcProfit (AI)'),
         TDO('VBA-overwritten with GAAP values'),
         TDG('Sheet formula =AG-AH restored')],
        [TD('Column labels (AD, AA)'),
         TD('"JTD Change In Profit"'),
         TDG('"MTD Change in Profit"')],
        [TD('Workbook version'),
         TD('5.70p'),
         TDG('5.71p (all fixes applied, modules imported)')],
    ]
    story += [tbl(delta_data, [1.7*inch, 2.5*inch, 3.3*inch]), sp(8), shr()]

    # == Technical Detail ======================================================
    story += [
        P('Technical Detail — Fixes Applied', 'h1'),
        P('Fix 1: MergePriorMonthProfitsOntoSheet (LylesWIPData.bas)', 'h2'),
        P(
            '<b>Before:</b> Computed recognizedProfit = earnedRev - priorJTDCost using a Vista JTD cost '
            'query, then wrote it directly to COLJTDPriorProfit (AC), destroying the formula.'
        ),
        P(
            '<b>After:</b> Looks up prior month BonusProfit (ov(8)) and OpsRev - OpsCost from LylesWIP, '
            'writes them to the hidden Z-columns:'
        ),
        B('COLZPriorBonusProfit (BK) — feeds AC via formula =BK'),
        B('COLZPriorJTDOPsProfit (BN) — feeds Q via formula =BN'),
        P(
            'The BuildPriorMonthCostLookup Vista query is no longer called (dead code).',
            'note'
        ),
        sp(4),
        P('Fix 2: MergePriorYearBonusOntoSheet (LylesWIPData.bas)', 'h2'),
        P(
            '<b>Before:</b> Wrote prior year-end BonusProfit to COLAPYBonusProfit (AJ) only.'
        ),
        P(
            '<b>After:</b> Also reads PYCost from COLAPYCost (AH, already written by GetWipDetail2) '
            'and updates COLAPYRev (AG) = PYCost + bonus. This corrects the baseline that GetWipDetail2 '
            'wrote with a 0 bonus stub.'
        ),
        sp(4),
        P('Fix 3: GetWIPDetailData_Modified.bas', 'h2'),
        P(
            '<b>Before:</b> Line 598 wrote PYEarnedRev - PYCost to COLAPYCalcProfit (AI), overwriting '
            'the sheet formula =AG6-AH6.'
        ),
        P(
            '<b>After:</b> Line removed. The sheet formula auto-calculates once AG and AH are correct.'
        ),
        sp(4),
        P('Fixes 4 &amp; 5: Workbook Label Changes', 'h2'),
        P(
            'Changed via COM automation (unprotect, edit, reprotect, save):'
        ),
        B('Jobs-Ops AD3: "JTD Change In Profit" &rarr; "MTD Change in Profit"'),
        B('Jobs-GAAP AA3: "JTD Change In Profit" &rarr; "MTD Change in Profit"'),
        sp(8), shr(),
    ]

    # == Current Workbook Versions =============================================
    story += [P('Current Workbook Versions', 'h1')]

    version_data = [
        [TH('Version'), TH('Purpose'), TH('Status')],
        [TD('5.68p'), TD('Nicole/Cindy demo version'), TD('Frozen')],
        [TD('5.70p'), TD('Query fixes + Push to Vista button'), TD('Superseded by 5.71p')],
        [TDB('5.71p'), TD('Nicole feedback fixes + label changes'), TDG('Active — ready for re-test')],
    ]
    story += [
        tbl(version_data, [1.0*inch, 3.25*inch, 3.25*inch]),
        sp(4),
        P('Source code: vba_source_569p/ (active development for 5.70p+).', 'note'),
        sp(8), shr(),
    ]

    # == Remaining for Delivery ================================================
    story += [P('Remaining for Delivery', 'h1')]

    remain_data = [
        [TH('#'), TH('Item'), TH('Priority'), TH('Status')],
        [TD('R1'), TD('Nicole/Cindy re-test on 5.71p'), TDR('Critical'), TDO('Ready — all 5 feedback items addressed')],
        [TD('R2'), TD('Multi-company testing (AIC, APC, NESM)'), TDR('High'), TDO('Not started')],
        [TD('R3'), TD('Verify residual rounding discrepancies'), TDR('High'), TDO('Spot-check after Nicole re-tests')],
        [TD('R4'), TD('Wire permissions to pnp.WIPSECGetRole'), TDO('Medium'), TDO('Not started')],
        [TD('R5'), TD('Security cleanup (Settings sheet, test credentials)'), TDR('High'), TDO('Not started')],
        [TD('R6'), TD('SQL driver install for Brian Platten + Harbir Atwal'), TDO('Medium'), TDO('Not started')],
        [TD('R7'), TD('Nicole questions: GAAP Billed to Date source, decimal display'), TD('Low'), TD('Pending response')],
        [TD('R8'), TD('Circular reference on 51.1158 (pre-existing)'), TD('Low'), TD("Won't block delivery")],
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
        [TD('Query fixes — Crystal Report exact match'),
         TD('Apr 7'), TDG('Done Apr 6')],
        [TDB('Nicole feedback — 5 items resolved'),
         TDB('Apr 7-11'), TDG('Done Apr 7')],
        [TD('Nicole/Cindy re-validation (5.71p)'),
         TD('Apr 7-11'), TDO('Ready')],
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
        P('Auto-WIP Status Report &nbsp;|&nbsp; April 7, 2026 &nbsp;|&nbsp; '
          'CONFIDENTIAL — Lyles Services Co. internal use only', 'footer'),
    ]

    doc.build(story)
    print('PDF saved: Auto-WIP_Status_Report_20260407.pdf')


if __name__ == '__main__':
    build()
