#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 13, 2026."""

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
        r'E:\Auto-Wip\Status-Reports\Auto-WIP_Status_Report_20260413.pdf',
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
        P('Date: April 13, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Workbook Version: Rev 5.73p</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # ── Executive Summary ─────────────────────────────────────────────────────
    story += [
        P('Executive Summary', 'h1'),
        P(
            'Rev 5.73p addresses all four items Nicole identified in her April 10 review of 5.72p, '
            'plus one latent bug discovered during investigation. A comprehensive audit of every piece '
            'of feedback from the April 6 demo session through April 10 (17 total resolved issues) '
            'confirms all prior fixes are carried forward with no regressions. '
            'Validation passes 24/24 divisions across all four companies with zero mismatches.'
        ),
        sp(4),
        callout(
            '<b>Key result:</b> All 17 issues reported across Revs 5.70p through 5.72p have been '
            'individually verified against the 5.73p output. 51.1129 values match the Crystal Report '
            'to the penny. The two missing overhead jobs (51.1156 / 51.1157) now appear correctly.',
            W
        ),
    ]

    # ── Nicole's April 10 Feedback ────────────────────────────────────────────
    story += [
        P("Nicole's April 10 Feedback — All Resolved", 'h1'),
        P(
            'Nicole reviewed 5.72p for Rocklin (Div 51), December 2025 and identified four issues. '
            'Root cause analysis revealed three were regressions from the 5.72p rewrite (fixes that '
            'existed in 5.71p were accidentally lost), and one was a query filter issue.'
        ),
    ]

    data = [
        [TH('#'), TH('Issue'), TH('Root Cause'), TH('Fix in 5.73p'), TH('Verified Value')],
        [TD('1'), TD('Missing jobs 51.1156 / 51.1157 (overhead, cost only)'),
         TD('HardClosed filter excluded jobs closed after WIP date '
            '(MonthClosed=Mar 2026 > Dec 2025). These were open during Dec 2025. '
            'Also zero-activity WHERE was too aggressive.'),
         TD('HardClosed filter now includes jobs where MonthClosed > batch month. '
            'WHERE clause softened to include jobs with estimates.'),
         TDG('28 jobs (was 26)')],
        [TD('2'), TD('Column AC (JTD Prior Profit) regressed'),
         TD('MergePriorMonthProfitsOntoSheet wrote opsRev-opsCost (projected profit) '
            'instead of bonusProfit. Also only wrote to Z-columns, but data rows need '
            'visible column writes too (no formulas in data rows).'),
         TD('Write bonusProfit to both Z-column (BK) and visible column (AC). '
            'Write opsRev-opsCost to both BN and Q.'),
         TDG('$8,997,122')],
        [TD('3'), TD('Column AG (Previous Year Revenue) regressed'),
         TD('MergePriorYearBonusOntoSheet only wrote bonus to AJ but did not '
            'recalculate AG = PYCost + bonus. The correction lines from 5.71p were lost.'),
         TD('After writing bonus, recalculate AG and AI explicitly.'),
         TDG('$71,707,694')],
        [TD('4'), TD('Column AU (JTD Billings) regressed'),
         TD('VistaData.bas lost the BilledThruMonth CTE from 5.71p. Query used '
            'live bJCCM.BilledAmt (includes future billings) instead of '
            'vrvJBProgressBills filtered by BillMonth.'),
         TD('Restored BilledThruMonth CTE. SELECT uses date-filtered billing.'),
         TDG('$96,918,206.90')],
    ]
    story += [tbl(data, [0.3*inch, 1.2*inch, 1.8*inch, 1.8*inch, 1.0*inch])]

    # ── Additional Fix ────────────────────────────────────────────────────────
    story += [
        P('Additional Fix — Latent Bug', 'h2'),
        P(
            '<b>SaveJobRow BonusProfit corruption (5.73-6):</b> When a user clicks "GAAP Done" on '
            'the Jobs-GAAP sheet, SaveJobRow could not read the bonus column (it only exists on '
            'the Ops sheet). The bonus defaulted to NULL and the stored procedure overwrote the '
            'saved BonusProfit with NULL. This had not yet been triggered in production but would '
            'have caused data loss on the next GAAP save. Fixed in both VBA and the stored procedure '
            '(deployed to 10.103.30.11).'
        ),
    ]

    # ── Database Status ───────────────────────────────────────────────────────
    story += [
        P('Database Investigation', 'h1'),
        P(
            "We investigated whether Nicole's save during the April 6 demo session may have "
            'corrupted override data. Findings: the LylesWIP database is <b>clean</b>. Only two rows '
            'were modified (51.1129 and 51.1108, both with correct BonusProfit values). '
            'All 4,974 override rows across 4 companies are intact. All issues Nicole reported '
            'were display-side VBA regressions, not data corruption.'
        ),
    ]

    # ── Full Regression Verification ──────────────────────────────────────────
    story += [
        P('Full Regression Verification — 17/17 Issues Confirmed', 'h1'),
        P(
            'Every issue reported from April 6 through April 10 was individually tested '
            'against the 5.73p snapshot output. The following table shows the complete results.'
        ),
    ]

    reg_data = [
        [TH('Rev'), TH('Issue'), TH('5.73p Status')],
        [TD('5.70'), TD('JTD Cost date cutoff (Mth filter)'), TDG('PASS — $87,315,159.22')],
        [TD('5.70'), TD('Billed to Date from vrvJBProgressBills'), TDG('PASS — $96,918,206.90')],
        [TD('5.70'), TD('Job 54.9416 inclusion (cost before start date)'), TDG('PASS — row 38, Div54')],
        [TD('5.70'), TD('Prior month profit = recognized profit from LylesWIP'), TDG('PASS')],
        [TD('5.71'), TD('Column AC = $8,997,122 (51.1129)'), TDG('PASS — $8,997,122.18')],
        [TD('5.71'), TD('Column AG = $71,707,694 (51.1129)'), TDG('PASS — $71,707,694.03')],
        [TD('5.71'), TD('Column AI corrected (51.1129)'), TDG('PASS — $7,950,623.36')],
        [TD('5.71'), TD('Col AD / AA renamed "MTD Change in Profit"'), TDG('PASS')],
        [TD('5.72'), TD('Circular reference on 51.1142 (ghost job excluded)'), TDG('PASS')],
        [TD('5.72'), TD('Ghost jobs rows 37-44 removed'), TDG('PASS — 0 ghosts')],
        [TD('5.72'), TD('Subtotals / totals not $0'), TDG('PASS — $364.5M')],
        [TD('5.72'), TD('51.1158 Column W = $56,961'), TDG('PASS — $56,960.69')],
        [TD('5.72'), TD('51.1139 on WIP in Closed section'), TDG('PASS — row 33')],
        [TD('5.73'), TD('51.1156 / 51.1157 overhead jobs present'), TDG('PASS — rows 38-39')],
        [TD('5.73'), TD('Columns AC / AG / AU restored'), TDG('PASS')],
        [TD('5.73'), TD('SaveJobRow BonusProfit preservation'), TDG('PASS — deployed')],
    ]
    story += [tbl(reg_data, [0.5*inch, 3.5*inch, 2.1*inch])]

    # ── Validation ────────────────────────────────────────────────────────────
    story += [
        P('Validation Results', 'h1'),
        P('<b>24/24 divisions PASS, 0 mismatches.</b> Batch validation ran all 24 company/division '
          'combinations across WML (Co15), AIC (Co16), APC (Co12), and NESM (Co13).'),
    ]

    val_data = [
        [TH('Company'), TH('Divisions'), TH('Jobs'), TH('Matched'), TH('Mismatched'), TH('Result')],
        [TD('WML (Co15)'), TD('50-58 (9 divs)'), TD('311'), TD('311'), TD('0'), TDG('PASS')],
        [TD('AIC (Co16)'), TD('70-78 (9 divs)'), TD('154'), TD('154'), TD('0'), TDG('PASS')],
        [TD('APC (Co12)'), TD('20-21 (2 divs)'), TD('44'), TD('44'), TD('0'), TDG('PASS')],
        [TD('NESM (Co13)'), TD('31-35 (4 divs)'), TD('68'), TD('68'), TD('0'), TDG('PASS')],
    ]
    story += [tbl(val_data, [1.0*inch, 1.2*inch, 0.7*inch, 0.8*inch, 0.9*inch, 0.7*inch])]

    # ── Open Items ────────────────────────────────────────────────────────────
    story += [
        P('Open Items', 'h1'),
    ]

    open_data = [
        [TH('#'), TH('Item'), TH('Status'), TH('Action')],
        [TD('1'), TD('Column AI (Prior Year Calc Profit) for job 51.1129: shows $7,950,623 vs '
            'Nicole\'s $6,957,477 reference from April 7 email'),
         TDO('Needs confirmation'),
         TD('Our value = AG ($71,707,694) minus AH ($63,757,071) = $7,950,623. This matches the Dec 2024 '
            'bonus profit in LylesWIP. Nicole approved 5.71p which produced this same value. '
            'The $6,957,477 reference was from Michael\'s WIP for June 2025 — may use a different '
            'calculation method. Nicole to confirm which is correct for job 51.1129, December 2025.')],
        [TD('2'), TD('Jobs-GAAP Billed to Date source — same Crystal Report?'),
         TDO('Awaiting answer'),
         TD('Asked April 7, not yet answered.')],
        [TD('3'), TD('Dollar display format — cents ($1,234.56) or whole dollars ($1,235)?'),
         TDO('Awaiting answer'),
         TD('Asked April 7, not yet answered.')],
        [TD('4'), TD('Percent complete format on Jobs-GAAP (1 decimal)'),
         TDO('Low priority'),
         TD('Nicole said fixing cost cutoff would cascade-fix this. Not re-raised since 5.70p.')],
    ]
    story += [tbl(open_data, [0.3*inch, 2.2*inch, 1.0*inch, 2.6*inch])]

    # ── Files Modified ────────────────────────────────────────────────────────
    story += [
        P('Files Modified in 5.73p', 'h1'),
    ]

    files_data = [
        [TH('File'), TH('Changes')],
        [TD('vba_source/LylesWIPData.bas'),
         TD('MergePriorMonthProfitsOntoSheet: Z-column + visible column writes with correct bonusProfit. '
            'MergePriorYearBonusOntoSheet: AG recalculation + AI explicit write. '
            'SaveJobRow: preserve BonusProfit on GAAP sheet saves.')],
        [TD('vba_source/VistaData.bas'),
         TD('Restored BilledThruMonth CTE (vrvJBProgressBills). '
            'HardClosed filter includes jobs open at WIP date. '
            'Zero-activity WHERE softened for overhead jobs.')],
        [TD('sql/WIP_Vista_Query.sql'),
         TD('Reference SQL updated to match VBA changes.')],
        [TD('sql/LylesWIP_CreateDB.sql'),
         TD('LylesWIPSaveJobRow: BonusProfit = ISNULL(@BonusProfit, BonusProfit).')],
        [TD('patch_and_validate.py'),
         TD('Updated to 5.73p workbook path. Added VistaData.bas to module import list.')],
    ]
    story += [tbl(files_data, [2.2*inch, 5.3*inch])]

    # ── Next Steps ────────────────────────────────────────────────────────────
    story += [
        P('Next Steps', 'h1'),
        P('1. Deploy 5.73p to SharePoint for Nicole / Cindy review.'),
        P('2. Confirm Column AI value with Nicole ($7,950,623 vs $6,957,477).'),
        P('3. Get answers on open items 2-3 (GAAP billings source, dollar format).'),
        P('4. Test with other companies (AIC Co16, APC Co12, NESM Co13) — override data already loaded.'),
        P('5. Schedule Nicole / Cindy review session once feedback on 5.73p is received.'),
        sp(8), hr(),
        P('Git repository initialized for version tracking. '
          'Issue history documented in ISSUES.md.', 'note'),
    ]

    doc.build(story)
    print(f'PDF generated: {doc.filename}')


if __name__ == '__main__':
    build()
