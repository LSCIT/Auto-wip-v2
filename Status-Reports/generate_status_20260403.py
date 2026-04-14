#!/usr/bin/env python3
"""Generate Auto-WIP Status Report PDF — April 3, 2026."""

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
        r'E:\Auto-Wip\Status-Reports\Auto-WIP_Status_Report_20260403.pdf',
        pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch,
        title='Auto-WIP Schedule — Project Status Report — April 3, 2026',
        author='Josh Garrison, IT Development',
    )

    story = []

    # == Header ================================================================
    story += [
        P('Auto-WIP Schedule — Project Status Report', 'title'),
        P('Date: April 3, 2026 &nbsp;&nbsp;|&nbsp;&nbsp; Prepared by: Josh Garrison, Director of Technology Innovation', 'subtitle'),
        P('Distribution: Kevin Shigematsu, Cindy Jordan, Nicole Leasure, Dane Wildey', 'subtitle'),
        P('<b>Priority: HIGHEST — Owner / Board / Senior Management escalation</b>', 'subtitle'),
        sp(4), hr(),
    ]

    # == Executive Summary =====================================================
    story += [
        P('Executive Summary', 'h1'),
        P(
            'April 3 was a breakthrough day. The full 3-stage workflow — from Accounting Initial '
            'Review through Operations Review to Accounting Final Approval — is now working '
            'end-to-end with live database persistence. Every feature required for the Nicole/Cindy '
            'demo is operational. Production readiness jumped from ~60-65% to ~85-90% in a single day.'
        ),
        P(
            'The day also marked a significant infrastructure change: development moved from Mac to '
            'the Windows VM, enabling direct Excel COM automation via Python/pywin32. This eliminated '
            'the manual VBA editor workflow and allowed rapid iteration — 15+ module imports and test '
            'cycles in one session.'
        ),
        sp(4),
        green_box(
            '<b>Updated production readiness: approximately 85-90%.</b> '
            'Write path (save overrides), 3-stage workflow, audit trail, and distribution are all '
            'working. Remaining work: multi-company testing, permissions wire, Nicole/Cindy demo, '
            'and Phase 2 design.<br/><br/>'
            '<b>Target delivery: May 8, 2026 — no change.</b>',
            W
        ),
        sp(8), shr(),
    ]

    # == What Changed Since April 2 ============================================
    story += [P('What Changed Since April 2, 2026', 'h1')]

    delta_data = [
        [TH('Item'), TH('Status Apr 2'), TH('Status Today (Apr 3)')],
        [TD('D2 Validation'),
         TDO('21/22 pass, 2 snapshots missing, 1 mismatch'),
         TDG('24/24 PASS, 0 mismatches — all 4 companies validated')],
        [TD('Write-back: Ops Done (Col H)'),
         TDR('Not built'),
         TDG('Working — SaveJobRow persists all override fields to LylesWIP')],
        [TD('Write-back: GAAP Done (Col I)'),
         TDR('Not built'),
         TDG('Working — same SaveJobRow path')],
        [TD('Write-back: Close (Col G)'),
         TDR('Not built'),
         TDG('Working — IsClosed flag persists')],
        [TD('Batch state machine'),
         TDR('Not wired'),
         TDG('Working — Open > ReadyForOps > OpsApproved > AcctApproved all tested')],
        [TD('CompleteCheck gates'),
         TDR('Not wired'),
         TDG('Working — blocks approval until all jobs marked Done')],
        [TD('State regression guards'),
         TDR('Not built'),
         TDG('Working — all No buttons blocked after approval set')],
        [TD('AcctApproved immutability'),
         TDR('Not built'),
         TDG('Working — edits blocked on both sheets after final approval')],
        [TD('Save &amp; Distribute to Ops'),
         TDR('Not built'),
         TDG('Working — saves .xlsm to C:\\Trusted\\ with ClearFormOnOpen=False')],
        [TD('Copy Ops to GAAP'),
         TDR('Not built'),
         TDG('Working — stored proc + VBA, scoped to current division')],
        [TD('December year-end snapshot'),
         TDR('Not wired'),
         TDO('Built — gated by AllBatchesApproved for company/month')],
        [TD('Audit trail comments'),
         TDR('Not built'),
         TDG('Working — "Changed $X to $Y by user on date" on all override cells')],
        [TD('Permissions module'),
         TDR('Not deployed'),
         TDO('Deployed — hardcoded WIPAccounting (P&amp;P proc wire deferred to Sprint 5)')],
        [TD('FormButtons.bas'),
         TDR("Old version (Michael's WipDb calls)"),
         TDG('Replaced — now calls LylesWIP UpdateBatchState correctly')],
        [TD('Application.Undo guards'),
         TDO('Not needed (manual only)'),
         TDG('36 calls guarded across 6 sheet modules')],
        [TD('Clear batch prompt'),
         TDO('Existed (destructive)'),
         TDG('Removed — batches are immutable after approval')],
        [TD('Development environment'),
         TDO('Mac + manual VBA editor'),
         TDG('Windows VM + pywin32 COM automation')],
    ]
    story += [tbl(delta_data, [1.9*inch, 2.2*inch, 3.4*inch]), sp(8), shr()]

    # == Bugs Fixed ============================================================
    story += [P('Bugs Fixed — April 3, 2026', 'h1')]

    bugs_data = [
        [TH('#'), TH('Bug'), TH('Impact / Fix')],
        [TD('1'),
         TD("FormButtons.bas called LCGWIPUpdateApprovals (Michael's old Vista proc) instead of Module6.UpdateApprovals"),
         TD('Batch state transitions silently failed. Rewired to LylesWIPData.UpdateBatchState.')],
        [TD('2'),
         TD('SaveJobRow used wrong column key "COLJob" instead of "COLJobNumber"'),
         TD('Every save crashed silently. Fixed key name.')],
        [TD('3'),
         TD('SaveJobRow crashed on GAAP sheet due to missing COLZOPsBonusNew column'),
         TD('Added On Error Resume Next guard for optional columns.')],
        [TD('4'),
         TD('ByRef argument type mismatch: GetOpsRevPlug and 8 other helpers declared row As Integer but SaveJobRow passes Long'),
         TD('Changed all 9 function signatures to Long.')],
        [TD('5'),
         TD('AFA (Accounting Final Approval) set AcctAppr flag BEFORE CompleteCheck'),
         TD('Allowed approval with incomplete jobs. Moved flag inside CompleteCheck gate.')],
        [TD('6'),
         TD('AFA radio button stayed on "Yes" visually even when CompleteCheck failed'),
         TD('Sheet12 checked the radio button shape directly, blocking all edits. Added Else branch to reset radio.')],
        [TD('7'),
         TD('Copy Ops to GAAP operated on all jobs for company/month instead of current division'),
         TD('Scoped SQL UPDATE to jobs on the current sheet only.')],
        [TD('8'),
         TD('Application.Undo calls in Worksheet_Change crashed with 1004 when called via COM'),
         TD('No user edit to undo via COM. Added On Error Resume Next guards to all 36 occurrences.')],
        [TD('9'),
         TD('GetWipDetail2 reload after Copy Ops to GAAP failed (Vista connection state issue)'),
         TD('Replaced with MergeOverridesOntoSheet (LylesWIP only, no Vista needed).')],
    ]
    story += [tbl(bugs_data, [0.3*inch, 3.0*inch, 4.2*inch]), sp(8), shr()]

    # == Nicole Confirmations ==================================================
    story += [
        P('Nicole Confirmations — April 3, 2026', 'h1'),
        B('<b>Close column (Col G)</b> = Ops recommends job be closed out. Accounting manually runs JC Contract Close in Vista (soft close if 100% complete + billed including retainage). Not an automated sync.'),
        B('<b>Copy Ops to GAAP</b> = Yes, Ops override values should carry over as starting point for GAAP overrides.'),
        sp(8), shr(),
    ]

    # == Production Readiness ==================================================
    story += [P('Production Readiness Assessment — April 3, 2026', 'h1')]

    readiness_data = [
        [TH('Layer'), TH('Apr 2'), TH('Apr 3'), TH('Notes')],
        [TD('Vista read path'),
         TDO('~82%'),
         TDG('~90%'),
         TD('A7 resolved implicitly (bJCCD EXISTS). A8 confirmed not needed (Close is workflow marker).')],
        [TD('GAAP formula accuracy'),
         TDO('~80%'),
         TDO('~85%'),
         TD('Col W and R working. Cascade columns need full-company verification.')],
        [TD('Override data load'),
         TDO('~85%'),
         TDG('~90%'),
         TD('All 4 companies loaded (4,974 rows). Validated 24/24 PASS.')],
        [TD('Write path — saving overrides'),
         TDR('0%'),
         TDG('~95%'),
         TD('Ops Done, GAAP Done, Close all persist. Copy Ops to GAAP working.')],
        [TD('Three-stage workflow'),
         TDR('15%'),
         TDG('~95%'),
         TD('Full cycle tested. State machine, guards, immutability all working.')],
        [TD('Permissions'),
         TDR('10%'),
         TDO('30%'),
         TD('Hardcoded WIPAccounting deployed. P&amp;P proc wire deferred.')],
        [TD('Distribution'),
         TDR('0%'),
         TDG('~90%'),
         TD('Save &amp; Distribute working. Multi-division distribution not tested.')],
        [TD('Audit trail'),
         TDR('0%'),
         TDG('~90%'),
         TD('Override comments with from/to/who/when on all cells.')],
        [TDB('Overall system'),
         TDR('~60-65%'),
         TDG('~85-90%'),
         TDB('Write path working. Demo-ready.')],
    ]
    story += [tbl(readiness_data, [1.5*inch, 0.7*inch, 0.7*inch, 4.6*inch]), sp(8), shr()]

    # == Remaining for Delivery ================================================
    story += [P('Remaining for Delivery', 'h1')]

    remain_data = [
        [TH('#'), TH('Item'), TH('Priority'), TH('Sprint')],
        [TD('1'), TD('Nicole / Cindy demo session — full 3-stage walkthrough'), TDR('Critical'), TD('Next')],
        [TD('2'), TD('Multi-company testing (AIC Co16, APC Co12, NESM Co13)'), TDR('High'), TD('4')],
        [TD('3'), TD('SQL driver install for Brian Platten and Harbir Atwal'), TDO('Medium'), TD('4')],
        [TD('4'), TD('Wire permissions to pnp.WIPSECGetRole'), TDO('Medium'), TD('5')],
        [TD('5'), TD('Security cleanup: re-hide Settings sheet, clear test credentials'), TDR('High'), TD('5')],
        [TD('6'), TD('Phase 2 design: Vista bJCOP/bJCOR write-back — separate button, gated by AcctApproved + GL period open'), TDO('Medium'), TD('Future')],
        [TD('7'), TD("Update Cost/Billed button — guard or remove (Michael's legacy)"), TDO('Low'), TD('Future')],
    ]
    story += [tbl(remain_data, [0.3*inch, 4.5*inch, 0.8*inch, 0.9*inch]), sp(8), shr()]

    # == Milestone Schedule ====================================================
    story += [P('Milestone Schedule', 'h1')]

    miles_data = [
        [TH('Milestone'), TH('Target'), TH('Status')],
        [TD('Vista read path — all 4 sheets'),
         TD('Mar 28'), TDG('DONE (Rev 5.68p deployed)')],
        [TD('LylesWIP database + stored procs'),
         TD('Mar 31'), TDG('DONE (8 procs on 10.103.30.11)')],
        [TD('Override load (40 files, 4,974 rows)'),
         TD('Apr 1'), TDG('DONE')],
        [TD('Validation pipeline (24/24 PASS)'),
         TD('Apr 4'), TDG('DONE Apr 3 — ahead of schedule')],
        [TD('Write path — SaveJobRow + Col H/I/G'),
         TD('Apr 11'), TDG('DONE Apr 3 — ahead of schedule')],
        [TD('3-stage workflow + state machine'),
         TD('Apr 14'), TDG('DONE Apr 3 — ahead of schedule')],
        [TD('Audit trail + distribution'),
         TD('Apr 14'), TDG('DONE Apr 3 — ahead of schedule')],
        [TD('Nicole / Cindy demo session'),
         TD('Apr 7-11'), TDO('PENDING — all features demo-ready')],
        [TD('Multi-company testing'),
         TD('Apr 14-18'), TDO('Not started')],
        [TD('Permissions wire (P&amp;P proc)'),
         TD('Apr 21-25'), TDO('Sprint 5 — deferred')],
        [TD('Security cleanup + final hardening'),
         TD('Apr 28 - May 2'), TDO('Not started')],
        [TD('Production delivery'),
         TD('May 8'), TDB('On track')],
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
        P('Auto-WIP Status Report &nbsp;|&nbsp; April 3, 2026 &nbsp;|&nbsp; '
          'CONFIDENTIAL — Lyles Services Co. internal use only', 'footer'),
    ]

    doc.build(story)
    print('PDF saved: Auto-WIP_Status_Report_20260403.pdf')


if __name__ == '__main__':
    build()
