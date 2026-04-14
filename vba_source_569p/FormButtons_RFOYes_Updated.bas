Attribute VB_Name = "FormButtons_RFOYes_Updated"
' =============================================================================
' FormButtons — RFO Click Handlers (Updated)
' REPLACES RFOYes_Click and RFONo_Click in FormButtons.bas
'
' Change from original:
'   Original: called UpdateApprovals → which called LCGWIPUpdateApprovals1 on WipDb
'   Updated:  calls LCGWIPUpdateApprovals (Viewpoint, no "1" suffix) directly
'             → writes to budMoJobSumAppr in Viewpoint (10.112.11.8)
'             No new tables or procs needed — uses Michael's existing Viewpoint proc.
'
' LCGWIPUpdateApprovals (Viewpoint) parameters:
'   @Co TINYINT, @Month DATE, @Dept VARCHAR(100),
'   @InitApproval CHAR(1), @FinalApproval CHAR(1),
'   @ReadyForOps CHAR(1), @AcctApproval CHAR(1),
'   @RetMsg VARCHAR(512) OUTPUT
'
' TO DEPLOY:
'   In VBA editor, open FormButtons module.
'   Find "Sub RFOYes_Click()" (near the top).
'   Select from there down through its "End Sub" and replace with the sub below.
'   Do the same for RFONo_Click.
' =============================================================================

Sub RFOYes_Click()
' ACCT - Ready For Ops (Yes)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If Len(Sheet17.Range("StartCompany").Value) <> 0 And _
   Len(Sheet17.Range("StartMonth").Value) <> 0 And _
   Len(Sheet17.Range("StartDept").Value) <> 0 Then

    Select Case Sheet2.Range("Role").Value

        Case "":
            ' No role loaded — do nothing

        Case "WIPAccounting":

            ' Update radio buttons and Settings flag (original behavior preserved)
            Sheet2.Range("ReadyForOpsAppr1").Value = "Y"
            Sheet2.Range("SendAppr").Value = "True"
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOn

            ' Call LCGWIPUpdateApprovals in Viewpoint (no "1" suffix)
            ' This writes ReadyForOPs1='Y' to budMoJobSumAppr for each dept.
            Dim co       As Integer
            Dim wipMonth As Date
            Dim deptList As String
            Dim retMsg   As String

            co       = CInt(Sheet17.Range("StartCompany").Value)
            wipMonth = CDate(Sheet17.Range("StartMonth").Value)
            deptList = CStr(Sheet17.Range("StartDept").Value)

            Dim conn As ADODB.Connection
            Set conn = GetVistaConnection()

            If conn Is Nothing Then
                MsgBox "Could not connect to Vista to save approval status.", vbCritical
                GoTo errexit
            End If

            Dim cmd As New ADODB.Command
            cmd.ActiveConnection = conn
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "LCGWIPUpdateApprovals"
            cmd.CommandTimeout = 30

            cmd.Parameters.Append cmd.CreateParameter("@Co", adTinyInt, adParamInput, , co)
            cmd.Parameters.Append cmd.CreateParameter("@Month", adDate, adParamInput, , wipMonth)
            cmd.Parameters.Append cmd.CreateParameter("@Dept", adVarChar, adParamInput, 100, deptList)
            cmd.Parameters.Append cmd.CreateParameter("@InitApproval", adChar, adParamInput, 1, "N")
            cmd.Parameters.Append cmd.CreateParameter("@FinalApproval", adChar, adParamInput, 1, "N")
            cmd.Parameters.Append cmd.CreateParameter("@ReadyForOps", adChar, adParamInput, 1, "Y")
            cmd.Parameters.Append cmd.CreateParameter("@AcctApproval", adChar, adParamInput, 1, "N")
            cmd.Parameters.Append cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)

            cmd.Execute
            retMsg = CStr(cmd.Parameters("@RetMsg").Value)

            MsgBox "Ready for Ops has been set." & vbCrLf & retMsg, _
                   vbInformation, "Ready for Ops"

            Set cmd = Nothing

        Case Else:
            MsgBox "Only Accounting Can Change This Setting", vbInformation
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff

    End Select

Else
    Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
    Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
    MsgBox "Select Company, Month And Division to Create new WIP Month", vbInformation
End If

Sheet17.Activate
Sheet17.Range("StartCompany").Select
GoTo 9999

errexit:
    Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
    Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
    MsgBox "There was an error in the RFOYes_Click Routine. " & Err, vbOKOnly
    ' Reset Settings flag so UI stays consistent
    Sheet2.Range("ReadyForOpsAppr1").Value = "N"

9999:
End Sub


Sub RFONo_Click()
' ACCT - Ready For Ops (No)
' NOTE: Clicking No does NOT reverse the batch state in budMoJobSumAppr.
' If someone clicked Yes (which saved to DB) and then clicks No, the DB
' still shows ReadyForOps=Y. This matches original behavior — reversal
' requires a separate workflow (handled in LCGWIPCancelBatch).
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If Len(Sheet17.Range("StartCompany").Value) <> 0 And _
   Len(Sheet17.Range("StartMonth").Value) <> 0 And _
   Len(Sheet17.Range("StartDept").Value) <> 0 Then

    Select Case Sheet2.Range("Role").Value

        Case "":
            ' No role loaded

        Case "WIPAccounting":
            Sheet2.Range("ReadyForOpsAppr1").Value = "N"
            Sheet2.Range("SendAppr").Value = "True"
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff

        Case Else:
            MsgBox "Only Accounting Can Change This Setting", vbInformation
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff

    End Select

Else
    Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
    Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
End If

Sheet17.Activate
Sheet17.Range("StartCompany").Select
GoTo 9999

errexit:
    MsgBox "There was an error in the RFONo_Click Routine. " & Err, vbOKOnly

9999:
End Sub
