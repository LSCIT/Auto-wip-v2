Attribute VB_Name = "GetudWIPJV"
' GetudWIPJV — Joint Venture data retrieval
' Modified: March 2026 — Phase 1 Vista Direct Connection
' Changes:
'   - Now queries Vista directly (udWIPJV + JC tables) instead of WipDb (LCGWIPGetDetailJV)
'   - JV master data (partners, share %) from udWIPJV table in Viewpoint
'   - Financial data (cost, billings, projections) from bJCCD/bJCCM
'   - Completion/workflow fields default to empty (Phase 1)

Public Sub GetudWIPJVSub(sh As Worksheet)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
Application.Calculation = xlCalculationManual

sh.Unprotect "password"

Dim Co As String
Dim Mo As String

sh.Activate
sh.Range("C11:C35").HorizontalAlignment = xlLeft
sh.Range("B11:B35").HorizontalAlignment = xlLeft

Co = CStr(Sheet17.Range("StartCompany").Value)
Mo = CStr(Sheet17.Range("StartMonth").Value)

Dim Datars As ADODB.Recordset

' Get JV data from Vista (replaces WipDb LCGWIPGetDetailJV)
If sh.CodeName = "Sheet14" Then
    ' First call (Sheet14 = JV's-Ops) — execute the query
    Set Datars = GetJVDataFromVista(CInt(Co), CDate(Mo))

    If Datars Is Nothing Then
        ' No JV data or connection error — skip silently
        GoTo 9999
    End If
Else
    ' Second call (Sheet15 = JV's-GAAP) — reuse recordset
    ' In the original, Datars was module-level and reused between calls.
    ' For Phase 1, we re-query (only 5 rows, fast)
    Set Datars = GetJVDataFromVista(CInt(Co), CDate(Mo))

    If Datars Is Nothing Then
        GoTo 9999
    End If
End If

If Datars.EOF <> True Then

    Datars.MoveLast
    Datars.MoveFirst

    T = Datars.RecordCount
    Dim i As Integer
    Dim r As Integer
    Dim hexString As String
    Dim ba As Integer
    r = 1

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

Datars.Close
Set Datars = Nothing

Application.Calculation = xlCalculationAutomatic

GoTo 9999
errexit:
MsgBox ("There was an error getting JV Data from Vista: " & Err.Description)
Application.Calculation = xlCalculationAutomatic
9999:

If Sheet2.Range("ProtectSheet").Value = True Then
    sh.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If

sh.Range("A1").Select
Application.EnableEvents = True

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
