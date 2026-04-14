
Public Sub ResetWorkbook()
    ' "Clear All" button handler — full reset for testing a new company/division.
    ' Uses Workbook_Open path so it also clears Start page selections.
    Call CloseVistaConnection
    Caller = "Workbook_Open"
    Call ClearForms3
    Caller = ""
End Sub


Public Sub ClearWIPDetailTable()
Application.EnableEvents = False

Call ClearWIPDetail
Caller = "Workbook_Open"
Call ClearForms3

Caller = ""

Application.EnableEvents = True

End Sub



Public Sub ClearForms3()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If Caller = "" Then
    
    
  
    Sheet11.Unprotect "password"
    Sheet12.Unprotect "password"
    Sheet13.Unprotect "password"
    Sheet14.Unprotect "password"
    Sheet15.Unprotect "password"
    Sheet16.Unprotect "password"

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End If

'Application.ScreenUpdating = False


'Ops Final Approval
'Yes
Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
'No
Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn

'Ready For OPs
'Yes
Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
'No
Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn

'Acct Final Approval
'Yes
Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
'No
Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn


Sheet2.Range("Sent").Value = "False"
Sheet2.Range("ReadyForOpsAppr1").Value = "N"
Sheet2.Range("InitAppr").Value = "N"
Sheet2.Range("FinalAppr").Value = "N"
Sheet2.Range("AcctAppr").Value = "N"
Sheet2.Range("SendAppr").Value = "False"
Sheet2.Range("SendJV").Value = "False"
Sheet2.Range("CompleteAll").Value = "False"
Sheet2.Range("CompleteAllGAAP").Value = "False"

'With Sheet11.Range("Done").Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'End With



ClearForm Sheet11
ClearForm Sheet12

ClearFormJV Sheet14

' JV's-GAAP (Sheet15) is SHA-512 protected with an unknown password.
' Unprotect "password" silently fails, so ClearContents would error.
' Since udWIPJV doesn't exist in production, Sheet15 always loads blank anyway.
On Error Resume Next
ClearFormJV Sheet15
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
Else
    On Error GoTo 0
End If


Select Case Caller

    Case "":

        Sheet2.Range("LastClosedMth").Value = ""

    Case "Workbook_Open":
    
        Sheet2.Range("LastClosedMth").Value = ""
        Sheet2.Range("GAAPView").Value = "N"
        Sheet17.Range("StartCompany").Value = ""
        Sheet17.Range("StartMonth").Value = ""
        Sheet17.Range("StartDept").Value = ""
        Sheet17.Range("StartCoName").Value = "<<< F4 will open Lookup"
        Sheet17.Range("StartDeptName").Value = "<<< F4 will open Lookup"
        Sheet17.Range("StartMonthStatus").Value = ""
        With Sheet17.Range("StartMonth").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 5287936
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With



End Select

Sheet17.Activate

If Caller = "" Then
    If Sheet2.Range("ProtectSheet").Value = "True" Then
        Sheet11.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet12.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet13.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet14.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet15.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet16.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True

    End If
    
    'Application.EnableEvents = True
    'Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End If

Sheet17.Range("StartCompany").Select


GoTo 9999
errexit:
MsgBox "There was an error In the ClearForms3 Routine" & Err, vbOKOnly

9999:

End Sub


Public Sub ClearForm(sh As Worksheet)
'Application.ScreenUpdating = False

If sh.CodeName = "Sheet11" Then

    With sh.Range("ExecOvrRevAdj").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With sh.Range("ExecOvrCostAdj").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With sh.Range("ExecOvrBonusAdj").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

 sh.Range("SummaryData").ClearComments
 sh.Range("SummaryData").ClearContents


With sh.Range("Done").Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With sh.Range("Done").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("Done").Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("Done").Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("Done").Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With


On Error Resume Next
With sh.Range("DoneGAAP").Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("DoneGAAP").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("DoneGAAP").Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("DoneGAAP").Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With sh.Range("DoneGAAP").Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With



 sh.Range("NotesCost").WrapText = False
 sh.Range("NotesRev").WrapText = False
 sh.Columns(LetDict(sh.CodeName)("COLOvrRevProjNotes")).Hidden = True
 sh.Columns(LetDict(sh.CodeName)("COLOvrCostProjNotes")).Hidden = True
 If sh.CodeName = "Sheet11" Then
    sh.Range("Done").ClearContent
 End If
 
 sh.Range("SummaryDataInput").ClearContents
 sh.Range("SummaryDataInput").ClearComments
 sh.Range("SummaryData").Font.Bold = False
 sh.Range("CalcCells").ClearContents
 sh.Range("CalcCells").Font.Bold = False
 sh.Activate


'Reset formulas
 sh.Range("Formulas").Copy
 sh.Range("SummaryDataInput").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    

 sh.Range("SummaryDataInput").Font.Bold = False

 sh.Range("Done").Font.Bold = True
 sh.Range("DoneGAAP").Font.Bold = True

Sheet2.Range("Sent").Value = "FALSE"
 sh.Range("B7").Select

End Sub


Public Sub ClearFormOpsVsGAAP()
'Application.ScreenUpdating = False

' SummaryData is sheet-scoped to Sheet11/Sheet12 — ignore if not found on Sheet13
On Error Resume Next
Sheet13.Range("SummaryData").ClearComments
Sheet13.Range("SummaryData").Font.Bold = False
On Error GoTo 0

End Sub


Public Sub ClearFormJV(sh)
'Application.ScreenUpdating = False


sh.Range("SummaryDataJV").ClearContents
sh.Range("ChangedJV").ClearContents
With sh.Range("SummaryDataJV")
    .ClearComments
End With

'Reset formulas

sh.Range("FormulasJV").Copy
sh.Range("SummaryDataJV").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'sh.Range("SummaryDataJV").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False



End Sub


Public Sub SetColWidth(sh As Worksheet)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.ScreenUpdating = False

Dim Dollars As Integer
Dim SP As Integer
Dim Note As Integer
Dim Profit As Integer
Dollars = 14
SP = 2
Note = 2
Profit = 13

'Ops
If sh.CodeName = "Sheet11" Then
    With sh
        .Columns(LetDict(sh.CodeName)("COLJTDBonusProfit")).ColumnWidth = 4
        .Columns(LetDict(sh.CodeName)("COLJTDBonusProfitNotesShow")).ColumnWidth = 4
        .Columns(LetDict(sh.CodeName)("COLAPYBonusProfit")).ColumnWidth = 4
        .Columns(LetDict(sh.CodeName)("COLCYBonusProfit")).ColumnWidth = 4
    End With
End If

'GAAP
If sh.CodeName = "Sheet12" Then
    With sh
        .Columns(LetDict(sh.CodeName)("COLGAAPDone")).ColumnWidth = 4
    End With
End If



 ' ALL Sheets
 
 With sh
    .Columns(LetDict(sh.CodeName)("COLJobNumber")).ColumnWidth = 9
    .Columns(LetDict(sh.CodeName)("COLJobDesc")).ColumnWidth = 60
    .Columns(LetDict(sh.CodeName)("COLPrjMngr")).ColumnWidth = 17
    .Columns(LetDict(sh.CodeName)("COLCurConAmt")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLPMProjRev")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLDone")).ColumnWidth = 4
    .Columns(LetDict(sh.CodeName)("COLOvrRevProj")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLOvrRevProjNotes2")).ColumnWidth = Note
    .Columns(LetDict(sh.CodeName)("COLPMProjCost")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLOvrCostProj")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLOvrCostProjNotes2")).ColumnWidth = Note
    .Columns(LetDict(sh.CodeName)("COLProjProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLPriorProjProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLChgProjProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLCompDate")).ColumnWidth = 11
    .Columns(LetDict(sh.CodeName)("COLSP1")).ColumnWidth = SP
    .Columns(LetDict(sh.CodeName)("COLJTDPctComp")).ColumnWidth = 9
    .Columns(LetDict(sh.CodeName)("COLJTDEarnedRev")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLJTDCost")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLJTDProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLJTDPriorProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLJTDChgProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLSP2")).ColumnWidth = SP
    .Columns(LetDict(sh.CodeName)("COLAPYPctComp")).ColumnWidth = 11
    .Columns(LetDict(sh.CodeName)("COLAPYRev")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLAPYCost")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLAPYCalcProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLSP3")).ColumnWidth = SP
    .Columns(LetDict(sh.CodeName)("COLCYRev")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLCYCost")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLCYCalcProfit")).ColumnWidth = Profit
    .Columns(LetDict(sh.CodeName)("COLSP4")).ColumnWidth = SP
    .Columns(LetDict(sh.CodeName)("COLBRev")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLBCost")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLBProfit")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLSP5")).ColumnWidth = SP
    .Columns(LetDict(sh.CodeName)("COLBILLBillings")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLBILLRevExBill")).ColumnWidth = Dollars
    .Columns(LetDict(sh.CodeName)("COLBILLBillExRev")).ColumnWidth = Dollars
End With






If sh.CodeName = "Sheet11" Then
    With sh
        .Columns(LetDict(sh.CodeName)("COLJTDBonusProfit")).ColumnWidth = Profit
        .Columns(LetDict(sh.CodeName)("COLJTDBonusProfitNotesShow")).ColumnWidth = Note
        .Columns(LetDict(sh.CodeName)("COLAPYBonusProfit")).ColumnWidth = Profit
        .Columns(LetDict(sh.CodeName)("COLCYBonusProfit")).ColumnWidth = Profit
    End With
End If

If sh.CodeName = "Sheet12" Then
    With sh
        .Columns(LetDict(sh.CodeName)("COLGAAPDone")).ColumnWidth = 4
    End With
End If


GoTo 9999
errexit:
MsgBox "There was an error in the SetColWidth Routine. " & Err, vbOKOnly
9999:

Application.ScreenUpdating = True
End Sub

Public Sub UnProtectAll()
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Sheet2.Unprotect "password"

Sheet11.Unprotect "password"
Sheet12.Unprotect "password"
Sheet13.Unprotect "password"
Sheet14.Unprotect "password"
Sheet15.Unprotect "password"
Sheet16.Unprotect "password"




Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub

Sub PrepForPublish()
    ActiveSheet.Shapes.Range(Array("Button 38")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 40")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 43")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 39")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 37")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 45")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 44")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 42")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 46")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 36")).Select
    Selection.Delete
    Selection.Cut
'    ActiveSheet.Shapes.Range(Array("Button 1")).Select
'    Selection.Delete
'    Selection.Cut
    Sheets("Settings").Select
    ActiveWindow.SmallScroll Down:=3
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "TRUE"
    Range("C16").Select
    ActiveCell.FormulaR1C1 = "TRUE"
    Range("C29").Select
    Selection.ClearContents
    Range("C33").Select
    Selection.ClearContents
    Range("C35").Select
    ActiveCell.FormulaR1C1 = "FALSE"
    Range("C37").Select
    Sheets("Start").Select
    
    Sheet11.Activate
    ActiveWindow.Zoom = 100
    
    Sheet12.Activate
    ActiveWindow.Zoom = 100

    Sheet13.Activate
    ActiveWindow.Zoom = 100
    
    Sheet14.Activate
    ActiveWindow.Zoom = 100
    
    Sheet15.Activate
    ActiveWindow.Zoom = 100
    
    Sheet16.Activate
    ActiveWindow.Zoom = 100
End Sub

