
Public Sub GetudWIPJVSub(sh As Worksheet)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
'Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

sh.Unprotect "password"

Dim Co As String
Dim Dept As String
Dim Mo As String
Dim YrStart As Date
Dim YrEnd As Date
Dim LastGAAPProjContract As Double
Dim PrevProjCost As Double
Dim ExecGAAPContractAdj As Double
Dim ExecGAAPCostAdj As Double
Dim cmd As ADODB.Command
Dim objStrSQL
Dim objStrSQL1
Dim objStrSQL2
Dim ba As Integer

sh.Activate
sh.Range("C11:C35").HorizontalAlignment = xlLeft
sh.Range("B11:B35").HorizontalAlignment = xlLeft

Co = CStr(Sheet17.Range("StartCompany").Value)
Dept = CStr(Sheet17.Range("StartDept").Value)
Mo = CStr(Sheet17.Range("StartMonth").Value)

YrStart = "01/01/" & DatePart("yyyy", Sheet17.Range("StartMonth").Value)
YrEnd = "12/31/" & (DatePart("yyyy", Sheet17.Range("StartMonth").Value) - 1)

Dim row As Range

' Open Connection if not already open
If cnn Is Nothing Then
    Set cnn = New ADODB.Connection
    Dim connectionString As String
    If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
        connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
    Else
        connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
    End If
    cnn.connectionString = connectionString
    cnn.Open
Else
    If cnn.State <> adStateOpen Then
        cnn.Open
    End If
End If

' Job Status Start and ending row
Dim StartRow As Integer
Dim EndRow As Integer

' Department start and ending rows
Dim DStartRow As Integer
Dim DEndRow As Integer

Dim OpenRowTotal As Integer

' Check if Datars is Nothing or closed, then open it
If Datars Is Nothing Then
    Set Datars = New ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "LCGWIPGetDetailJV"
    cmd.CommandTimeout = 180
    
    Set cmdCo = cmd.CreateParameter("@Co", adTinyInt, adParamInput, 1, Co)
    cmd.Parameters.Append cmdCo
    Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 0, CDate(Mo))
    cmd.Parameters.Append cmdMonth
    cnn.CursorLocation = adUseClient
    
    Datars.Open cmd.Execute
Else
        Datars.MoveFirst
End If


If Datars.EOF <> True Then

    Datars.MoveLast
    Datars.MoveFirst

    T = Datars.RecordCount
    Dim i As Integer
    Dim r As Integer
    Dim Div As String
    Div = "None"
    Dim Status As Integer
    Status = 0
    r = 1
    StartRow = sh.Range("SummaryDataJV").Rows(r).row
    DStartRow = StartRow
    
    Dim OCompleted As String
    Dim GCompleted As String

    For i = 1 To T
    
        
        If Datars.Fields("OCompleted").Value = "Y" Then
            OCompleted = "P"
        Else
            OCompleted = ""
        End If
        
        If Datars.Fields("GCompleted").Value = "Y" Then
            GCompleted = "P"
        Else
            GCompleted = ""
        End If

        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVMMJobNo")) = Datars.Fields("JVJobNum").Value
        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJobNo")) = Datars.Fields("IntJobNum").Value
        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVSupJobNo")) = Datars.Fields("SupJobNumber").Value
        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJobDesc")) = Datars.Fields("JVJobDesc").Value
        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVPartners")) = Datars.Fields("JVPartners").Value
        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVSharePct")) = Datars.Fields("OurJVPct").Value
        sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLZBatchSeq")) = Datars.Fields("BatchSeq").Value
        
        If sh.CodeName = "Sheet14" Then ' Ops
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLDone")) = OCompleted
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVCurContAmt")) = Datars.Fields("OpsContractAmt").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVOvrRevProj")) = Datars.Fields("OpsProjectedRevenue").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVOvrCostProj")) = Datars.Fields("OpsProjectedCost").Value
            
            If Datars.Fields("OpsEarnedRev").Value <> 0 Then
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJTDEarnedRev")) = Datars.Fields("OpsEarnedRev").Value
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVZJTDER2")) = Datars.Fields("OpsEarnedRev").Value
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVZJTDER")) = "T"
                With sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJTDEarnedRev")).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            Else
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVZudChg")) = 0
            End If
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJTDCost")) = Datars.Fields("OpsJTDCost").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVBILLBillings")) = Datars.Fields("OpsJTDBillings").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVAPYRev")) = Datars.Fields("PYOpsEarnedRevenue").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVAPYCost")) = Datars.Fields("PYOpsPJTDCost").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLZUserName")) = Datars.Fields("OUserName").Value
            For ba = LBound(Datars.Fields("ORowVersion").Value) To UBound(Datars.Fields("ORowVersion").Value)
                hexString = hexString & Right("0" & Hex(Datars.Fields("ORowVersion").Value(ba)), 2)
            Next ba
            
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLZRowVersion")) = hexString
            hexString = ""
        
        Else 'GAAP
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLDone")) = OCompleted
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLDoneGAAP")) = GCompleted
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVCurContAmt")) = Datars.Fields("GAAPContractAmt").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVProjFinalProfit")) = Datars.Fields("GAAPProjectedFinalProfit").Value
            If Datars.Fields("GAAPEarnedRev").Value <> 0 Then
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJTDEarnedRev")) = Datars.Fields("GAAPEarnedRev").Value
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVZJTDER2")) = Datars.Fields("GAAPEarnedRev").Value
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVZJTDER")) = "T"
                With sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJTDEarnedRev")).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
        
            Else
                sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVZudChg")) = 0
            End If
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVJTDCost")) = Datars.Fields("GAAPJTDCost").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVBILLBillings")) = Datars.Fields("GAAPJTDBillings").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVAPYRev")) = Datars.Fields("PYGAAPEarnedRevenue").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLJVAPYCost")) = Datars.Fields("PYGAAPJTDCost").Value
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLZUserName")) = Datars.Fields("GUserName").Value
            For ba = LBound(Datars.Fields("GRowVersion").Value) To UBound(Datars.Fields("GRowVersion").Value)
                hexString = hexString & Right("0" & Hex(Datars.Fields("GRowVersion").Value(ba)), 2)
            Next ba
            
            sh.Range("SummaryDataJV").Cells(r, NumDict(sh.CodeName)("COLZRowVersion")) = hexString
            hexString = ""
        
        End If
        
        Datars.MoveNext
        r = r + 1
    
    Next

End If

If sh.CodeName = "Sheet15" Then
    Datars.Close
    cnn.Close
End If

Application.Calculation = xlCalculationAutomatic

GoTo 9999
errexit:
MsgBox ("There was an error geting Data From Viewpoint")
If Datars.State <> 0 Then
Datars.Close
End If
cnn.Close
Application.Calculation = xlCalculationAutomatic
9999:

If Sheet2.Range("ProtectSheet").Value = True Then
    sh.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If

sh.Range("A1").Select
Application.EnableEvents = True
'Application.ScreenUpdating = True


Sheet17.Activate

End Sub



Public Sub SetTotalsJV(r As Integer, StartRow As Integer, EndRow As Integer, Col As String, sh As Worksheet)

sh.Range("SummaryDataJV").Cells(r, ColLetToNum(Col)).Formula = "=Sum(" & Col & StartRow & ":" & Col & EndRow & ")"
sh.Range("SummaryDataJV").Cells(r, ColLetToNum(Col)).Font.Bold = True

End Sub

Public Sub SetDTotalsJV(r As Integer, StartRow As Integer, EndRow As Integer, Col As String, sh As Worksheet)

sh.Range("SummaryDataJV").Cells(r, ColLetToNum(Col)).Formula = "=Sum(" & Col & StartRow & ":" & Col & EndRow & ")/2"
sh.Range("SummaryDataJV").Cells(r, ColLetToNum(Col)).Font.Bold = True

End Sub



Public Sub SetGrandTotalsJV(r As Integer, OpenRowTotal As Integer, EndRow As Integer, Col As String, sh As Worksheet)

sh.Range("SummaryDataJV").Cells(r, ColLetToNum(Col)).Formula = "=Sum(" & Col & "5" & ":" & Col & EndRow + 1 & ")/4"
' Open Job Totals
sh.Range("SummaryDataJV").Cells(r - 2, ColLetToNum(Col)).Formula = "=SUMIF(AX5:AX" & EndRow & ",1, " & Col & "5" & ":" & Col & EndRow & ")"

' Closed Job Totals
sh.Range("SummaryDataJV").Cells(r - 1, ColLetToNum(Col)).Formula = "=SUMIF(AX5:AX" & EndRow & ",2, " & Col & "5" & ":" & Col & EndRow & ")"


sh.Range("SummaryDataJV").Cells(r, ColLetToNum(Col)).Font.Bold = True

End Sub
