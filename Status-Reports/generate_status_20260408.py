#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 8, 2026."""

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
        r'E:\Auto-Wip\Status-Reports\Auto-WIP_Status_Report_20260408.pdf',
        pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch,
        title='Auto-WIP Schedule -- Project Status Report -- April 8, 2026',
        author='Josh Garrison, IT Development',
    )

    story = []

    # == Header ================================================================
    story += [
        P('Auto-WIP Schedule -- Project Status Report', 'title'),
        P('Date: April 8, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Priority: HIGHEST -- Owner / Board / Senior Management escalation</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # == Executive Summary =====================================================
    story += [
        P('Executive Summary', 'h1'),
        P(
            'Received second round of detailed feedback from Nicole Leasure on Rev 5.71p (five issues on '
            'Jobs-Ops tab for Rocklin Div 51). Retrieved and analyzed Michael Roberts\' original stored '
            'procedures from WipDb (LCGWIPCreateBatch, LCGWIPUpdateCostBill, LCGWIPGetDetailPM) which '
            'revealed the root cause of all five issues: differences in job filtering, override column '
            'defaulting, and contract status mapping between Michael\'s pre-populated WIPDetail table and '
            'our live Vista query approach. All five issues resolved in Rev 5.72p, plus eight additional '
            'improvements including 6-month trend history tooltips, fiscal month date filtering, and '
            'comparison sheet restoration.'
        ),
        sp(4),
        green_box(
            '<b>Production readiness: approximately 92-95%.</b> '
            'All Nicole feedback items from 5.71p addressed. Michael\'s architecture fully understood '
            'and incorporated. 6-month trend tooltips implemented. Jobs-Ops vs GAAP comparison tab '
            'formulas restored (pending validation).<br/><br/>'
            '<b>Target delivery: May 8, 2026 -- no change.</b>',
            W
        ),
        sp(8), shr(),
    ]

    # == Key Discovery =========================================================
    story += [
        P("Key Discovery -- Michael's Original Architecture", 'h1'),
        P(
            'Retrieved Michael Roberts\' stored procedures from WipDb using SA access. This was the '
            '"Rosetta Stone" for understanding all of Nicole\'s reported issues.'
        ),
        sp(4),
    ]

    arch_data = [
        [TH('Proc'), TH('Purpose'), TH('Key Finding')],
        [TD('LCGWIPCreateBatch'),
         TD('Populates WIPDetail from Vista'),
         TD('Job filter: HardClosed only if MonthClosed in WIP year. '
            'Override columns default to Vista projected values when no explicit override. '
            'ContractStatus mapped to 1/2 only.')],
        [TD('LCGWIPUpdateCostBill'),
         TD('Refreshes cost/billing in WIPDetail'),
         TD('Uses bJCCP for cost (not bJCCD). Uses Mth for date filtering (not PostedDate).')],
        [TD('LCGWIPGetDetailPM'),
         TD('Reads WIPDetail for Excel'),
         TD('Simple SELECT with permissions filter. No Vista queries -- all pre-populated.')],
    ]
    story += [tbl(arch_data, [1.5*inch, 1.8*inch, 4.2*inch]), sp(8), shr()]

    # == Nicole's Feedback =====================================================
    story += [
        P("Nicole's Feedback -- April 7-8, 2026 (All Resolved)", 'h1'),
        P('<b>Source:</b> Email from Nicole Leasure reviewing Rev 5.71p, plus phone call April 8', 'body'),
        P('<b>Division:</b> Company 15 (WML), Division 51 (Rocklin)', 'body'),
        sp(4),
    ]

    issues_data = [
        [TH('#'), TH("Nicole's Report"), TH('Root Cause'), TH('Fix'), TH('Status')],
        [TD('1'),
         TD('Circular reference on 51.1142'),
         TD('Vista ContractStatus=3 (HardClosed) passed raw to VBA. '
            'VBA only handles 1/2; the 2-to-3 transition fired subtotal logic incorrectly, '
            'creating SUM ranges that included themselves.'),
         TD('ContractStatus mapped to 1 (Open) or 2 (Closed) in SQL, matching Michael\'s mapping.'),
         TDG('Fixed')],
        [TD('2'),
         TD('Closed jobs in rows 37-44 should not show on WIP'),
         TD('EXISTS clauses caught $0 ghost transactions (SL records with ActualCost=0) '
            'pulling in jobs closed in prior years.'),
         TD('HardClosed jobs now only included if MonthClosed within the WIP year (Jan 1 - Dec 1). '
            'Zero-activity/zero-contract exclusion added.'),
         TDG('Fixed')],
        [TD('3'),
         TD('Subtotals and totals showing $0 for some columns (e.g. Column W)'),
         TD('Override columns I (Revenue) and M (Cost) left at $0 when no explicit override. '
            'All downstream formulas (% Complete, Earned Revenue, Projected Profit) depend on I/M.'),
         TD('When no override: I defaults to Vista projected revenue, '
            'M defaults to MAX(Vista projected cost, actual cost). Matches Michael\'s defaulting.'),
         TDG('Fixed')],
        [TD('4'),
         TD('Job 51.1158 Column W shows $0, should be $56,961'),
         TD('Same as #3, plus template circular formula in Column Z (W-Z-Y-W chain) '
            'prevented recalculation.'),
         TD('Override defaulting (#3) + VBA always writes value to Column Z, breaking circular chain. '
            'Column W now shows $56,960.69.'),
         TDG('Fixed')],
        [TD('5'),
         TD('Job 51.1139 not on Auto WIP. Had $481 profit change in 2025 to make job '
            '100% complete and closed. Rule: if open on prior year WIP, must appear.'),
         TD('Job was included by HardClosed-in-WIP-year filter (MonthClosed = Dec 2025) '
            'but sorted into Open section. Nicole confirmed: belongs in Closed.'),
         TD('ContractStatus mapping: Vista-closed jobs shown in Closed section. '
            '$481 CY revenue visible in Column AL. Bonus = $4,907.56.'),
         TDG('Fixed')],
    ]
    story += [tbl(issues_data, [0.3*inch, 1.5*inch, 1.8*inch, 2.1*inch, 0.6*inch]), sp(8), shr()]

    # == Additional Fixes ======================================================
    story += [P('Additional Improvements in 5.72p', 'h1')]

    addl_data = [
        [TH('#'), TH('Item'), TH('Detail')],
        [TD('6'),
         TD('Fiscal month date filtering'),
         TD('Inline VBA SQL was using PostedDate/ActualDate. Michael used Mth (fiscal month) via bJCCP. '
            'Now consistent -- 51.1129 JTD Cost matches Crystal Report exactly ($87,315,159.22).')],
        [TD('7'),
         TD('Jobs-GAAP Column Z threshold restored'),
         TD('10% completion threshold (IF V>0.1) had been removed. Restored across Z6:Z436.')],
        [TD('8'),
         TD('Save &amp; Distribute timestamp'),
         TD('Output filename includes timestamp: "WIP Schedule - 15 Div51 Dec2025 20260408-1541.xlsm"')],
        [TD('9'),
         TD('6-month override trend tooltips'),
         TD('Hover tooltips on Columns I/M showing 6 months of override history from LylesWIP.')],
        [TD('10'),
         TD('6-month Vista projected trend tooltips'),
         TD('Hover tooltips on Columns F/L showing 6 months of bJCOR/bJCOP projection history.')],
        [TD('11'),
         TD('Prior month projected profit'),
         TD('Column Q now populates from LylesWIP prior month OpsRev - OpsCost. Was showing $0.')],
        [TD('12'),
         TD('Database index for performance'),
         TD('Covering index IX_WipJobData_CoMonth on LylesWIP WipJobData (JCCo, WipMonth). '
            'Supports 6 extra trend queries per sheet load without performance degradation.')],
    ]
    story += [tbl(addl_data, [0.3*inch, 2.0*inch, 5.2*inch]), sp(8), shr()]

    # == Validated Data Points =================================================
    story += [
        P('Validated Data Points', 'h1'),
        B('<b>51.1129</b> JTD Cost = $87,315,159.22 -- Crystal Report exact match (Nicole confirmed "looks good")'),
        B('<b>51.1139</b> in Closed section with bonus $4,907.56 and $481 CY revenue change'),
        B('<b>51.1158</b> Column W = $56,960.69 ($56,961 as Nicole expected)'),
        B('<b>26 jobs total</b> for Co 15 Div 51 Dec 2025 (was ~46 with ghost-transaction jobs)'),
        B('<b>All historical override data valid</b> from December 2024 forward (Nicole confirmed)'),
        sp(8), shr(),
    ]

    # == What Changed Since April 7 ============================================
    story += [P('What Changed Since April 7, 2026', 'h1')]

    delta_data = [
        [TH('Item'), TH('Status Apr 7'), TH('Status Today (Apr 8)')],
        [TD('Nicole feedback (5 items)'),
         TDO('Received, investigating'),
         TDG('All 5 resolved in 5.72p')],
        [TD('Michael\'s stored procs'),
         TDO('Architecture unknown'),
         TDG('Retrieved, analyzed, incorporated')],
        [TD('Job filtering'),
         TDR('Ghost transactions pulling in old closed jobs'),
         TDG('HardClosed-in-WIP-year filter matching Michael\'s logic')],
        [TD('Override defaulting'),
         TDR('Columns I/M empty for non-override jobs'),
         TDG('Defaults to Vista projected values')],
        [TD('ContractStatus mapping'),
         TDR('Raw 1/2/3 causing circular refs'),
         TDG('Mapped to 1/2 -- no circular references')],
        [TD('Date filtering'),
         TDO('PostedDate/ActualDate (diverges from Crystal Report)'),
         TDG('Fiscal Mth -- exact Crystal Report match')],
        [TD('6-month trend tooltips'),
         TDR('All blank'),
         TDG('Override + Vista projected trends populated')],
        [TD('Prior month profit (Col Q)'),
         TDR('$0 for all jobs'),
         TDG('Populated from LylesWIP')],
        [TD('Workbook version'),
         TD('5.71p'),
         TDG('5.72p')],
    ]
    story += [tbl(delta_data, [2.0*inch, 2.5*inch, 3.0*inch]), sp(8), shr()]

    # == Current Workbook Versions =============================================
    story += [P('Current Workbook Versions', 'h1')]

    version_data = [
        [TH('Version'), TH('Purpose'), TH('Status')],
        [TD('5.47p'), TD('Michael Roberts\' original (baseline for formula comparison)'), TD('Reference only')],
        [TD('5.68p'), TD('Validation baseline (24/24 PASS)'), TD('Frozen')],
        [TD('5.71p'), TD('Nicole feedback round 1 fixes'), TD('Superseded by 5.72p')],
        [TDB('5.72p'), TD('Nicole feedback round 2 + Michael\'s architecture alignment'), TDG('Active -- ready for review')],
    ]
    story += [tbl(version_data, [1.0*inch, 3.25*inch, 3.25*inch]), sp(8), shr()]

    # == Remaining for Delivery ================================================
    story += [P('Remaining for Delivery', 'h1')]

    remain_data = [
        [TH('#'), TH('Item'), TH('Priority'), TH('Status')],
        [TD('R1'), TD('Nicole/Cindy re-test on 5.72p'), TDR('Critical'), TDO('Ready for review')],
        [TD('R2'), TD('Jobs-Ops vs GAAP comparison tab -- formulas restored, needs debugging'), TDR('High'), TDO('Pending')],
        [TD('R3'), TD('Multi-company testing (AIC, APC, NESM)'), TDR('High'), TDO('Not started')],
        [TD('R4'), TD('Row 19 Col R floating point display ($0 vs dash)'), TD('Low'), TDO('Template format issue')],
        [TD('R5'), TD('Wire permissions to pnp.WIPSECGetRole'), TDO('Medium'), TDO('Not started')],
        [TD('R6'), TD('Security cleanup (Settings sheet, test credentials)'), TDR('High'), TDO('Not started')],
        [TD('R7'), TD('SQL driver install for Brian Platten + Harbir Atwal'), TDO('Medium'), TDO('Not started')],
        [TD('R8'), TD('Clean up stale Phase 1/Phase 2 comments in source'), TD('Low'), TDO('Not started')],
    ]
    story += [tbl(remain_data, [0.4*inch, 3.3*inch, 0.8*inch, 3.0*inch]), sp(8), shr()]

    # == Milestone Schedule ====================================================
    story += [P('Milestone Schedule', 'h1')]

    miles_data = [
        [TH('Milestone'), TH('Target'), TH('Status')],
        [TD('Vista read path -- all 4 sheets'), TD('Mar 28'), TDG('Done')],
        [TD('LylesWIP database + stored procs'), TD('Apr 14'), TDG('Done Mar 31')],
        [TD('Override load (40 files, 4,974 rows)'), TD('Apr 17'), TDG('Done Mar 31')],
        [TD('Validation pipeline (24/24 PASS)'), TD('Apr 4'), TDG('Done Apr 3')],
        [TD('Write path + 3-stage workflow'), TD('Apr 14'), TDG('Done Apr 3')],
        [TD('Query fixes -- Crystal Report exact match'), TD('Apr 7'), TDG('Done Apr 6')],
        [TD('Nicole feedback round 1 (5 items)'), TD('Apr 7-11'), TDG('Done Apr 7')],
        [TDB('Nicole feedback round 2 (5 items) + Michael\'s procs'), TDB('Apr 8'), TDG('Done Apr 8')],
        [TD('Nicole/Cindy re-validation (5.72p)'), TD('Apr 9-11'), TDO('Ready')],
        [TD('Multi-company testing'), TD('Apr 14-18'), TDO('Not started')],
        [TD('Permissions wire (P&amp;P proc)'), TD('Apr 21-25'), TDO('Not started')],
        [TD('Security cleanup'), TD('Apr 28 - May 2'), TDO('Not started')],
        [TDB('Production delivery'), TDB('May 8'), TDB('On track')],
    ]
    story += [tbl(miles_data, [3.0*inch, 1.2*inch, 3.3*inch]), sp(8), shr()]

    # == Files Modified ========================================================
    story += [P('Files Modified', 'h1')]

    files_data = [
        [TH('File'), TH('Changes')],
        [TD('vba_source/VistaData.bas'), TD('SQL filter, Mth dating, ContractStatus mapping, Vista trend tooltips')],
        [TD('vba_source/GetWIPDetailData_Modified.bas'), TD('Override defaulting, bonus write, trend calls, Dim fix')],
        [TD('vba_source/LylesWIPData.bas'), TD('Override trend tooltips, timestamp filename')],
        [TD('sql/WIP_Vista_Query.sql'), TD('Reference SQL updated (ContractStatus, HardClosed filter, zero-activity WHERE)')],
        [TD('sql/reference/'), TD('Michael\'s procs archived (LCGWIPCreateBatch, UpdateCostBill, GetDetailPM)')],
        [TD('ARCHITECTURE.md'), TD('Updated with Michael\'s architecture findings and override defaulting rules')],
        [TD('vm/WIPSchedule -Rev 5.72p.xltm'), TD('Active workbook with all fixes')],
    ]
    story += [tbl(files_data, [2.5*inch, 5.0*inch]), sp(8)]

    # == Footer ================================================================
    story += [
        shr(), sp(4),
        P('Auto-WIP Status Report &nbsp;|&nbsp; April 8, 2026 &nbsp;|&nbsp; '
          'CONFIDENTIAL -- Lyles Services Co. internal use only', 'footer'),
    ]

    doc.build(story)
    print('PDF saved: Auto-WIP_Status_Report_20260408.pdf')


if __name__ == '__main__':
    build()
