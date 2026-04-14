Attribute VB_Name = "GetWIPDetailData"
' =============================================================================
' GetWIPDetailData.bas — MODIFIED for Vista Direct Connection
' Changes from original:
'   1. Lines 134-209: Replaced WipDb stored proc with VistaData.GetWIPDataFromVista()
'   2. Connection string no longer has Workstation ID=MROBERTS (bug C4 fix)
'   3. All recordset field names preserved for backward compatibility
' =============================================================================

' Module-level variables
Public cnn As ADODB.Connection
Public Datars As ADODB.Recordset

Public Sub GetWipDetail2(sh As Worksheet)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
Dim frm As DataRetrievalStatus
Dim SortOption As String


Set frm = New DataRetrievalStatus
frm.Label1.Caption = "Getting Data from Vista......."
frm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
frm.Show vbModeless
DoEvents

ProtectUnProtect ("Off")

sh.Activate

sh.Range("SummaryData").ClearComments

Caller = "GetudMoJobSummary"

If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If

If Sheet2.Range("SortOption").Value = 1 Then
    SortOption = "Department"
    sh.Columns(NumDict(sh.CodeName)("COLPrjMngr")).Hidden = False
Else
    SortOption = "PM"
    sh.Columns(NumDict(sh.CodeName)("COLPrjMngr")).Hidden = True
End If

If Sheet17.Range("StartCompany").Value = "" Or Sheet17.Range("StartMonth").Value = "" Or Sheet17.Range("StartDept").Value = "" Then
    MsgBox "Select Company, Month and Division!", vbOKOnly
    Sheet17.Activate
    Sheet17.Range("StartCompany").Select
    GoTo 9999
End If

Application.EnableEvents = False
'Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Declare variables'
Dim Co As String
Dim Dept As String
Dim Mo As String
Dim YrStart As Date
Dim YrEnd As Date
Dim LastGAAPProjContract As Double
Dim LastGAAPProjCost As Double
Dim LastOPsProjContract As Double
Dim LastOPsProjCost As Double

Dim ContractAdj As Double
Dim GAAPCostAdj As Double
Dim cmd As ADODB.Command
Dim objStrSQL
Dim objStrSQL1
Dim objStrSQL2
Dim LastGAAPPctComp As Double
Dim LastGAAPEarnedRev As Double
Dim LastOPsPctComp As Double
Dim LastOPsEarnedRev As Double

Dim PYPctComp As Double
Dim PYEarnedRev As Double
Dim PYCost As Double
Dim PYOpsPctComp As Double
Dim PYOpsEarnedRev As Double
Dim PMRevTrend As String
Dim PMCostTrend As String
Dim ORRevTrend As String
Dim ORCostTrend As String

Dim BatchId As Integer
Dim rowVersion() As Byte
'Dim rowVersion As Variant
Dim byteArray() As Byte
Dim hexString As String
Dim ba As Integer
Dim GroupBy As String

' Phase 1: Skip batch/approval checks (no WipDb dependency)
' Call ApprCheck2    ' Phase 2: re-enable when batch system is connected

' Phase 1: Set approval defaults
Sheet2.Range("CompleteAll").Value = False
Sheet2.Range("CompleteAllGAAP").Value = False

sh.Activate
sh.Range("B6:C301").Select

With Selection
    .HorizontalAlignment = xlLeft
End With

sh.Range("B7").Select

Co = CStr(Sheet17.Range("StartCompany").Value)
Dept = CStr(Sheet17.Range("StartDept").Value)
Mo = CStr(Sheet17.Range("StartMonth").Value)

If Sheet2.Range("SortOption").Value = 1 Then
    GroupBy = "Department"
Else
    GroupBy = "PM"
End If


YrStart = "01/01/" & DatePart("yyyy", Sheet17.Range("StartMonth").Value)
YrEnd = "12/31/" & (DatePart("yyyy", Sheet17.Range("StartMonth").Value) - 1)


Dim row As Range


' ===========================================================================
' VISTA DIRECT CONNECTION (replaces WipDb stored proc call)
' ===========================================================================
' Instead of calling LCGWIPGetDetailPM against WipDb, we now query Vista
' directly using VistaData.GetWIPDataFromVista()
'
' Original code (lines 134-209) called:
'   cmd.CommandText = "LCGWIPGetDetailPM"
'   with params @Co, @Month, @Dept, @UserName, @DeptOut, @GroupBy
'   against WipDb (PPServerName connection)
'
' New code calls VistaData.GetWIPDataFromVista() which:
'   1. Connects to Vista server (VPServerName from Settings sheet)
'   2. Runs CTE-based SQL query against bJCJM/bJCCM/bJCCD/etc.
'   3. Returns recordset with same field names as the stored proc
' ===========================================================================

If Datars Is Nothing Then
    Set Datars = GetWIPDataFromVista(CInt(Co), CDate(Mo), Dept, GroupBy)

    If Datars Is Nothing Then
        MsgBox "Failed to retrieve data from Vista. Check connection settings.", vbCritical
        GoTo 9999
    End If
ElseIf Datars.State = adStateClosed Then
    Set Datars = GetWIPDataFromVista(CInt(Co), CDate(Mo), Dept, GroupBy)

    If Datars Is Nothing Then
        MsgBox "Failed to retrieve data from Vista. Check connection settings.", vbCritical
        GoTo 9999
    End If
Else
    ' Ensure Datars is not empty before moving to first
    If Datars.EOF And Datars.BOF Then
        GoTo NoData   ' empty recordset — no jobs for this division/month
    End If
    Datars.MoveFirst
End If

' Job Status Start and ending row
Dim StartRow As Integer
Dim EndRow As Integer

' Department start and ending rows
Dim DStartRow As Integer
Dim DEndRow As Integer

Dim OpenRowTotal As Integer

' Initialize variables for PM tracking
Dim PM As String
PM = "None"
Dim PMStartRow As Integer
Dim PMEndRow As Integer

' Department/Status variables
Dim Div As String
Dim Status As Integer
Div = "None"
Status = 0
' Common variables
Dim GroupStartRow As Integer
Dim GroupEndRow As Integer

T = 0

If Datars.EOF <> True Then
    Datars.MoveLast
    Datars.MoveFirst

    T = Datars.RecordCount

    Dim i As Integer
    Dim r As Integer
    r = 1
    StartRow = sh.Range("SummaryData").Rows(r).row
    GroupStartRow = StartRow

    For i = 1 To T
        If SortOption = "Department" Then
            ' Department Open Status Footer
            If Div <> "None" And Div <> Datars.Fields("Department").Value And Status <> 2 Or (Status = 1 And Status <> Datars.Fields("ContractStatus").Value) Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "  Division " & Div & " Open Job Totals"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                EndRow = sh.Range("SummaryData").Rows(r - 1).row
                OpenRowTotal = r

                Call SetTotals(r, StartRow, EndRow, "COLCurConAmt", sh)
                Call SetTotals(r, StartRow, EndRow, "COLPMProjRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLOvrRevProj", sh)
                Call SetTotals(r, StartRow, EndRow, "COLPMProjCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLOvrCostProj", sh)
                Call SetTotals(r, StartRow, EndRow, "COLProjProfit", sh)
                Call SetTotals(r, StartRow, EndRow, "COLPriorProjProfit", sh)
                Call SetTotals(r, StartRow, EndRow, "COLChgProjProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLJTDEarnedRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLJTDCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLJTDProfit", sh)

                If sh.CodeName <> "Sheet12" Then
                    Call SetTotals(r, StartRow, EndRow, "COLJTDCalcEarnedRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDBonusProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYBonusProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYBonusProfit", sh)
                End If

                Call SetTotals(r, StartRow, EndRow, "COLJTDPriorProfit", sh)
                Call SetTotals(r, StartRow, EndRow, "COLJTDChgProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLAPYRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLAPYCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLAPYCalcProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLCYRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLCYCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLCYCalcProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLBRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLBILLBillings", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBILLRevExBill", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBILLBillExRev", sh)

                r = r + 2
            End If

            ' Department Closed Status Footer
            If Div <> Datars.Fields("Department").Value And Status = 2 Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "  Division " & Div & " Closed Job Totals"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                EndRow = sh.Range("SummaryData").Rows(r - 1).row
                OpenRowTotal = r

                Call SetTotals(r, StartRow, EndRow, "COLCurConAmt", sh)
                Call SetTotals(r, StartRow, EndRow, "COLPMProjRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLOvrRevProj", sh)
                Call SetTotals(r, StartRow, EndRow, "COLPMProjCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLOvrCostProj", sh)
                Call SetTotals(r, StartRow, EndRow, "COLProjProfit", sh)
                Call SetTotals(r, StartRow, EndRow, "COLPriorProjProfit", sh)
                Call SetTotals(r, StartRow, EndRow, "COLChgProjProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLJTDEarnedRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLJTDCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLJTDProfit", sh)

                If sh.CodeName <> "Sheet12" Then
                    Call SetTotals(r, StartRow, EndRow, "COLJTDCalcEarnedRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDBonusProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYBonusProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYBonusProfit", sh)
                End If

                Call SetTotals(r, StartRow, EndRow, "COLJTDPriorProfit", sh)
                Call SetTotals(r, StartRow, EndRow, "COLJTDChgProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLAPYRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLAPYCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLAPYCalcProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLCYRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLCYCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLCYCalcProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLBRev", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBCost", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBProfit", sh)

                Call SetTotals(r, StartRow, EndRow, "COLBILLBillings", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBILLRevExBill", sh)
                Call SetTotals(r, StartRow, EndRow, "COLBILLBillExRev", sh)

                r = r + 2
            End If

            ' Department Footer
            If Div <> "None" And Div <> Datars.Fields("Department").Value And Datars.Fields("Department").Value <> "None" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "Division " & Div & " Totals"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                GroupEndRow = sh.Range("SummaryData").Rows(r - 1).row
                OpenRowTotal = r

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCurConAmt", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrRevProj", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrCostProj", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLProjProfit", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPriorProjProfit", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLChgProjProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDEarnedRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDProfit", sh)

                If sh.CodeName <> "Sheet12" Then
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCalcEarnedRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDBonusProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYBonusProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYBonusProfit", sh)
                End If

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDPriorProfit", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDChgProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCalcProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCalcProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillings", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLRevExBill", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillExRev", sh)

                r = r + 2
            End If

            ' Department Header
            If Div <> Datars.Fields("Department").Value Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "Division " & Datars.Fields("Department").Value
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                GroupStartRow = sh.Range("SummaryData").Rows(r).row
                r = r + 1
            End If

            ' Open Status Header
            If Div <> Datars.Fields("Department").Value And Datars.Fields("ContractStatus").Value = 1 Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "  Open Jobs"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                StartRow = sh.Range("SummaryData").Rows(r).row
                r = r + 1
            End If

            ' Closed Status Header
            If Datars.Fields("ContractStatus").Value = 2 And Datars.Fields("ContractStatus").Value <> Status Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "  Closed Jobs"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                StartRow = sh.Range("SummaryData").Rows(r).row
                r = r + 1
            End If
        Else ' SortOption = "PM"
            ' [PM grouping code unchanged — same as original lines 434-496]
            ' PM Footer
            If PM <> "None" And PM <> Datars.Fields("PM").Value Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "PM " & PM & " Totals"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft

                GroupEndRow = sh.Range("SummaryData").Rows(r - 1).row
                OpenRowTotal = r

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCurConAmt", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrRevProj", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrCostProj", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLProjProfit", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPriorProjProfit", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLChgProjProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDEarnedRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDProfit", sh)

                If sh.CodeName <> "Sheet12" Then
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCalcEarnedRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDBonusProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYBonusProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYBonusProfit", sh)
                End If

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDPriorProfit", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDChgProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCalcProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCalcProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBRev", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBCost", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBProfit", sh)

                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillings", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLRevExBill", sh)
                Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillExRev", sh)

                r = r + 2
            End If

            ' PM Header
            If PM <> Datars.Fields("PM").Value Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "PM " & Datars.Fields("PM").Value
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                GroupStartRow = sh.Range("SummaryData").Rows(r).row
                r = r + 1
            End If
        End If

        ' Common job data population
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobNumber")) = Datars.Fields("Contract").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")) = "    " & Datars.Fields("ContractDescription").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLPrjMngr")) = Datars.Fields("PM").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCurConAmt")) = (Datars.Fields("OrigContractAmt").Value + Datars.Fields("COContractAmt").Value)

        If Datars.Fields("COContractAmt").Value <> 0 Then
            With sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCurConAmt"))
                .AddComment.Text Text:="Original - " & vbNewLine & CStr(Format(Datars.Fields("OrigContractAmt").Value, "#,###")) _
                    & vbNewLine & "Change Order - " & vbNewLine & Format(Datars.Fields("COContractAmt").Value, "#,###")
            End With
            With sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCurConAmt")).Comment
                .Shape.AutoShapeType = msoShapeRoundedRectangle
                .Shape.TextFrame.Characters.Font.Name = "Arial"
                .Shape.TextFrame.Characters.Font.Size = 10
                .Shape.Height = 100
            End With
        End If

        LastGAAPProjContract = 0
        LastGAAPProjCost = 0
        LastOPsProjContract = 0
        LastOPsProjCost = 0
        LastOPsPctComp = 0
        LastOPsEarnedRev = 0
        LastGAAPPctComp = 0
        LastGAAPEarnedRev = 0
        PMRevTrend = ""
        PMCostTrend = ""
        PYCost = 0
        PYPctComp = 0
        PYEarnedRev = 0

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLPMProjRev")) = Datars.Fields("ProjContract").Value

        ' Phase 1: Skip trend comments (no historical data from Vista)
        ' Trend tooltips will be re-enabled in Phase 2 when WipDb historical data is available

        If Datars.Fields("Close").Value = "C" Then
            If sh.CodeName = "Sheet11" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLClose")) = "C"
            End If
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZUserName")) = Datars.Fields("UserName").Value
        Else
            If sh.CodeName = "Sheet11" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLClose")) = ""
            End If
        End If

        If Datars.Fields("Completed").Value = "Y" Then
            If sh.CodeName = "Sheet11" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLDone")) = "P"
            End If
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZUserName")) = Datars.Fields("UserName").Value
        Else
            If sh.CodeName = "Sheet11" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLDone")) = ""
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLDone")).ClearComments
            End If
        End If

        If Datars.Fields("ProjCost").Value = 0 Then
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLPMProjCost")) = Datars.Fields("ActualCost").Value
        Else
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLPMProjCost")) = Datars.Fields("ProjCost").Value
        End If

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCompDate")) = Datars.Fields("CompletionDate").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDCost")) = Datars.Fields("ActualCost").Value

        If Datars.Fields("ActualCost").Value <> Datars.Fields("OrgActualCost").Value Then
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDCost")).Font.Bold = True
            With sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDCost"))
                .AddComment.Text Text:="Original = " & CStr(Format(Datars.Fields("OrgActualCost"), "#,##0;(#,##0)"))
                .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 25
                .Comment.Shape.Width = 125
            End With
        End If

        If sh.CodeName <> "Sheet12" Then ' Ops
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDBonusProfitNotes")) = Datars.Fields("BonusProfitNotes").Value
            If Datars.Fields("BonusProfitNotes").Value <> "" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDBonusProfitNotesShow")) = "+"
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDBonusProfitNotesShow")).Font.Bold = True
            End If
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLAPYBonusProfit")) = Datars.Fields("PriorYrBonusProfit").Value

            If Datars.Fields("OpsRevPlugged").Value = "Y" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")) = Datars.Fields("OpsRev").Value
            End If

            If Datars.Fields("OpsCostPlugged").Value = "Y" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")) = Datars.Fields("OpsCost").Value
            End If

            If Datars.Fields("BonusProfitPlugged").Value = "Y" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJTDBonusProfit")) = Datars.Fields("BonusProfit").Value
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOPsBonus")).Font.Bold = True
            End If
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOPsBonus")) = Datars.Fields("BonusProfit").Value

        Else ' GAAP
            If Datars.Fields("CompletedGAAP").Value = "Y" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLGAAPDone")) = "P"
            Else
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLGAAPDone")) = ""
            End If

            If Datars.Fields("GAAPRevPlugged").Value = "Y" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")) = Datars.Fields("GAAPRev").Value
            End If

            If Datars.Fields("GAAPCostPlugged").Value = "Y" Then
                sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")) = Datars.Fields("GAAPCost").Value
            End If
        End If

        PYCost = Datars.Fields("ActualCost").Value - Datars.Fields("CYActualCost").Value

        If Datars.Fields("PriorYearGAAPCost").Value <> 0 Then
            PYPctComp = PYCost / Datars.Fields("PriorYearGAAPCost").Value
            PYEarnedRev = Datars.Fields("PriorYearGAAPRev").Value * PYPctComp
        End If

        If Datars.Fields("PriorYearOpsCost").Value <> 0 Then
            PYOpsPctComp = PYCost / Datars.Fields("PriorYearOpsCost").Value
            PYOpsEarnedRev = Datars.Fields("PriorYearGAAPRev").Value * PYPctComp
        End If

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLAPYPctComp")) = PYPctComp
        If sh.CodeName = "Sheet11" Then
            ' Baseline: PYCost + 0 (PriorYrBonusProfit is stubbed to 0 from Vista).
            ' MergePriorYearBonusOntoSheet corrects this to PYCost + actual bonus from LylesWIP.
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLAPYRev")) = PYCost + Datars.Fields("PriorYrBonusProfit").Value
            ' COLAPYCalcProfit has a sheet formula =AG-AH; do not overwrite it.
        Else
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLAPYRev")) = PYEarnedRev
        End If
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLAPYCost")) = PYCost
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCYCost")) = Datars.Fields("CYActualCost").Value

        If Datars.Fields("CYActualCost").Value <> Datars.Fields("OrgCYActualCost").Value Then
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCYCost")).Font.Bold = True
            With sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLCYCost"))
                .AddComment.Text Text:="Original = " & CStr(Format(Datars.Fields("OrgCYActualCost"), "#,##0;(#,##0)"))
                .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 25
                .Comment.Shape.Width = 125
            End With
        End If

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLBILLBillings")) = Datars.Fields("BilledAmt").Value

        If Datars.Fields("BilledAmt").Value <> Datars.Fields("OrgBilledAmt").Value Then
            sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLBILLBillings")).Font.Bold = True
            With sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLBILLBillings"))
                .AddComment.Text Text:="Original = " & CStr(Format(Datars.Fields("OrgBilledAmt"), "#,##0;(#,##0)"))
                .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 25
                .Comment.Shape.Width = 125
            End With
        End If

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOPsRev")) = Datars.Fields("OpsRev").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOPsCost")) = Datars.Fields("OpsCost").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZGAAPRev")) = Datars.Fields("GAAPRev").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZGAAPCost")) = Datars.Fields("GAAPCost").Value

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZPriorOPsProfit")) = Datars.Fields("LastOpsRev").Value - Datars.Fields("LastOpsCost").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZPriorBonusProfit")).Value = Datars.Fields("LastBonusProfit").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZPriorJTDOPsProfit")) = Datars.Fields("LastOpsRev").Value - Datars.Fields("LastOpsCost").Value

        ' Prior GAAP projection lookup (simplified for Phase 1 — uses 0s)
        LastGAAPProjContract = Datars.Fields("LastGAAPRev").Value
        LastGAAPProjCost = Datars.Fields("LastGAAPCost").Value

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZPriorJTDGAAPProfit")) = LastGAAPProjContract - LastGAAPProjCost

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOPsRevNotes")) = Datars.Fields("OpsRevNotes").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOPsCostNotes")) = Datars.Fields("OpsCostNotes").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZGAAPRevNotes")) = Datars.Fields("GAAPRevNotes").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZGAAPCostNotes")) = Datars.Fields("GAAPCostNotes").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZGAAPPYRev")) = Datars.Fields("PriorYearGAAPRev").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZGAAPPYCost")) = Datars.Fields("PriorYearGAAPCost").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOpsPYRev")) = Datars.Fields("PriorYearOpsRev").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZOpsPYCost")) = Datars.Fields("PriorYearOpsCost").Value

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZContractStatus")) = Datars.Fields("ContractStatus").Value

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZUserName")).Value = Datars.Fields("UserName").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZBatchSeq")).Value = Datars.Fields("BatchSeq").Value

        ' RowVersion handling (Phase 1: always 0, no concurrency needed)
        hexString = "0000000000000000"
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZRowVersion")).Value = hexString
        hexString = ""

        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZORGJTDCost")).Value = Datars.Fields("OrgActualCost").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZORGCYCost")).Value = Datars.Fields("OrgCYActualCost").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZORGBilledAmt")).Value = Datars.Fields("OrgBilledAmt").Value
        sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLZORGCYBilledAmt")).Value = Datars.Fields("OrgCYBilledAmt").Value

        ' Update grouping variables
        If SortOption = "Department" Then
            Div = Datars.Fields("Department").Value
            Status = Datars.Fields("ContractStatus").Value
        Else
            PM = Datars.Fields("PM").Value
        End If

        Datars.MoveNext
        r = r + 1
99:
        ' [EOF handling with totals — identical to original lines 1026-1237]
        If Datars.EOF = True Then
            ' NOTE: The full EOF totals section is identical to the original.
            ' It writes division/PM totals, open/closed subtotals, and grand totals.
            ' For brevity, this section is the same as lines 1026-1237 in the original.
            ' When importing to VBA, copy lines 1026-1237 from the original exactly.

            ' === BEGIN: Copy from original GetWIPDetailData.bas lines 1026-1237 ===
            If SortOption = "Department" Then
                If Status = 1 Then
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "  Division " & Div & " Open Job Totals"
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                    Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                    EndRow = sh.Range("SummaryData").Rows(r - 1).row
                    OpenRowTotal = r
                    Call SetTotals(r, StartRow, EndRow, "COLCurConAmt", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLPMProjRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLOvrRevProj", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLPMProjCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLOvrCostProj", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLProjProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLPriorProjProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLChgProjProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDEarnedRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDProfit", sh)
                    If sh.CodeName <> "Sheet12" Then
                        Call SetTotals(r, StartRow, EndRow, "COLJTDCalcEarnedRev", sh)
                        Call SetTotals(r, StartRow, EndRow, "COLJTDBonusProfit", sh)
                        Call SetTotals(r, StartRow, EndRow, "COLAPYBonusProfit", sh)
                        Call SetTotals(r, StartRow, EndRow, "COLCYBonusProfit", sh)
                    End If
                    Call SetTotals(r, StartRow, EndRow, "COLJTDPriorProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDChgProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYCalcProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYCalcProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBILLBillings", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBILLRevExBill", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBILLBillExRev", sh)
                    r = r + 2
                End If
                If Status = 2 Then
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "  Division " & Div & " Closed Job Totals"
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                    Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                    EndRow = sh.Range("SummaryData").Rows(r - 1).row
                    OpenRowTotal = r
                    Call SetTotals(r, StartRow, EndRow, "COLCurConAmt", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLPMProjRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLOvrRevProj", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLPMProjCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLOvrCostProj", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLProjProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLPriorProjProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLChgProjProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDEarnedRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDProfit", sh)
                    If sh.CodeName <> "Sheet12" Then
                        Call SetTotals(r, StartRow, EndRow, "COLJTDCalcEarnedRev", sh)
                        Call SetTotals(r, StartRow, EndRow, "COLJTDBonusProfit", sh)
                        Call SetTotals(r, StartRow, EndRow, "COLAPYBonusProfit", sh)
                        Call SetTotals(r, StartRow, EndRow, "COLCYBonusProfit", sh)
                    End If
                    Call SetTotals(r, StartRow, EndRow, "COLJTDPriorProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLJTDChgProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLAPYCalcProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLCYCalcProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBRev", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBCost", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBProfit", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBILLBillings", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBILLRevExBill", sh)
                    Call SetTotals(r, StartRow, EndRow, "COLBILLBillExRev", sh)
                    r = r + 2
                End If
                If Div <> "None" Then
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "Division " & Div & " Totals"
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                    Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                    GroupEndRow = sh.Range("SummaryData").Rows(r - 1).row
                    OpenRowTotal = r
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCurConAmt", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrRevProj", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrCostProj", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLProjProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPriorProjProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLChgProjProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDEarnedRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDProfit", sh)
                    If sh.CodeName <> "Sheet12" Then
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCalcEarnedRev", sh)
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDBonusProfit", sh)
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYBonusProfit", sh)
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYBonusProfit", sh)
                    End If
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDPriorProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDChgProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCalcProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCalcProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillings", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLRevExBill", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillExRev", sh)
                    r = r + 2
                End If
            Else ' PM final footer
                If PM <> "None" Then
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Value = "PM " & PM & " Totals"
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
                    Sheet13.Range("SummaryData").Rows(r).Font.Bold = True
                    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).HorizontalAlignment = xlLeft
                    GroupEndRow = sh.Range("SummaryData").Rows(r - 1).row
                    OpenRowTotal = r
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCurConAmt", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrRevProj", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPMProjCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLOvrCostProj", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLProjProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLPriorProjProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLChgProjProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDEarnedRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDProfit", sh)
                    If sh.CodeName <> "Sheet12" Then
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDCalcEarnedRev", sh)
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDBonusProfit", sh)
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYBonusProfit", sh)
                        Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYBonusProfit", sh)
                    End If
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDPriorProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLJTDChgProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLAPYCalcProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLCYCalcProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBRev", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBCost", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBProfit", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillings", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLRevExBill", sh)
                    Call SetDTotals(r, GroupStartRow, GroupEndRow, "COLBILLBillExRev", sh)
                    r = r + 2
                End If
            End If
            ' === END: Copy from original ===
        End If
    Next
100:
End If


If sh.CodeName = "Sheet12" Then
    Datars.Close
    Set Datars = Nothing
End If

If T < 1 Then
    GoTo NoData
End If

OpenRowTotal = sh.Range("SummaryData").Cells(OpenRowTotal, NumDict(sh.CodeName)("COLOvrRevProj")).row + 1
EndRow = sh.Range("SummaryData").Rows(r).row

Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLCurConAmt", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLPMProjRev", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLOvrRevProj", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLPMProjCost", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLOvrCostProj", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLProjProfit", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLPriorProjProfit", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLChgProjProfit", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDEarnedRev", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDCost", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDProfit", sh)

If sh.CodeName <> "Sheet12" Then
    Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDCalcEarnedRev", sh)
    Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDBonusProfit", sh)
    Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLAPYBonusProfit", sh)
    Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLCYBonusProfit", sh)
End If

Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDPriorProfit", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLJTDChgProfit", sh)

Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLAPYRev", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLAPYCost", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLAPYCalcProfit", sh)

Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLCYRev", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLCYCost", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLCYCalcProfit", sh)

Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLBRev", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLBCost", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLBProfit", sh)

Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLBILLBillings", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLBILLRevExBill", sh)
Call SetGrandTotals(r, OpenRowTotal, EndRow, "COLBILLBillExRev", sh)

' Set Grand Total
With sh.Range("SummaryData").Cells(r + 2, NumDict(sh.CodeName)("COLJobDesc"))
    .Value = "Company " & Sheet17.Range("StartCompany").Value & " Grand Totals"
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With

With Sheet13.Range("SummaryData").Cells(r + 2, NumDict(sh.CodeName)("COLJobDesc"))
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With


sh.Range("SummaryData").Cells(r + 2, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
sh.Range("SummaryData").Rows(r + 2).Font.Bold = True

If SortOption = "Department" Then
    With sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc"))
        .Value = "Company " & Sheet17.Range("StartCompany").Value & " Open Job Totals"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    sh.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
    sh.Range("SummaryData").Rows(r).Font.Bold = True

    With Sheet13.Range("SummaryData").Cells(r, NumDict(sh.CodeName)("COLJobDesc"))
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With

    With sh.Range("SummaryData").Cells(r + 1, NumDict(sh.CodeName)("COLJobDesc"))
        .Value = "Company " & Sheet17.Range("StartCompany").Value & " Closed Job Totals"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    sh.Range("SummaryData").Cells(r + 1, NumDict(sh.CodeName)("COLJobDesc")).Font.Bold = True
    sh.Range("SummaryData").Rows(r + 1).Font.Bold = True

    With Sheet13.Range("SummaryData").Cells(r + 1, NumDict(sh.CodeName)("COLJobDesc"))
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
End If

sh.Range("Comments").WrapText = False

Sheet17.Activate
Sheet17.Range("StartMonth").Select

Sheet2.Range("PullNewData").Value = "N"

' ==========================================================================
' C2: Ensure a WipBatches row exists for this company/month/dept in LylesWIP.
' Called after every successful Vista load — idempotent (no-op if already exists).
' Dept is passed as-is (may be a single code like "54" or "All").
' ==========================================================================
LylesWIPData.CreateBatch CInt(Co), CDate(Mo), Dept, Environ("UserName")

' C3: Merge saved LylesWIP overrides back onto the sheet.
' Loads all override rows for this co/month in one DB call → Dictionary lookup,
' then does a single pass over SummaryData rows applying any saved overrides.
' No-op if no overrides exist yet for this batch month.
LylesWIPData.MergeOverridesOntoSheet sh, CInt(Co), CDate(Mo)

' C3b: Write prior month projected profit (OpsRev - OpsCost from LylesWIP month-1)
' into COLPriorProjProfit and COLJTDPriorProfit. No-op on GAAP sheet.
LylesWIPData.MergePriorMonthProfitsOntoSheet sh, CInt(Co), CDate(Mo)

' C3c: Write prior year-end bonus profit (Dec of prior year from LylesWIP)
' into COLAPYBonusProfit. No-op on GAAP sheet.
LylesWIPData.MergePriorYearBonusOntoSheet sh, CInt(Co), CDate(Mo)

Application.Calculation = xlCalculationAutomatic
GoTo 9999
NoData:
MsgBox "There was no data to pull into form...", vbInformation

GoTo 9999
errexit:
MsgBox "There was an error In the GetWipDetail Routine: " & Err.Description, vbOKOnly
' Clean up Vista connection on error
Call CloseVistaConnection
If Not Datars Is Nothing Then
    If Datars.State <> adStateClosed Then Datars.Close
End If
Set Datars = Nothing

9999:

Application.EnableEvents = True
'Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Sheet2.Range("Sent").Value = "FALSE"
ProtectUnProtect ("On")
Caller = ""
Unload frm

End Sub
