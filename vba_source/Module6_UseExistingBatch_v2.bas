Attribute VB_Name = "Module6_UseExistingBatch_v2"
' =============================================================================
' UseExistingBatch — v2 (Batch-Aware, Viewpoint direct)
' REPLACES the UseExistingBatch sub in Module6.
'
' Uses Michael's existing LCGWIPBatchCheck proc in Viewpoint to check for
' existing batches — no new tables or procs needed.
'
' Behavior:
'   1. Calls LCGWIPBatchCheck (Viewpoint, no "1" suffix) to check if
'      budWIPDetail already has rows for the selected co/month.
'   2. If rows found (rcode=0): prompt user to reopen those depts.
'      - Yes: sets StartDept to the existing dept list, NoData = False
'      - No:  shows Dept picker for fresh dept selection
'   3. If no rows (rcode=1): shows Dept picker, NoData = False after selection.
'
' NOTE: budWIPDetail is not yet populated (no batch creation proc exists yet).
'   Until Sprint 2 (LylesWIPLoadBatch), LCGWIPBatchCheck will always return
'   "no rows" (rcode=1), so the Dept picker will always show for fresh starts.
'   This is correct behavior for Stage 1.
'
' LCGWIPBatchCheck (Viewpoint) parameters:
'   @Co TINYINT, @Month DATE,
'   @rcode INT OUTPUT, @DeptList VARCHAR(200) OUTPUT
'
' TO DEPLOY: In Module6, find "Public Sub UseExistingBatch(ByRef NoData As Boolean)"
'            at the very top (around line 2). Select from there down to its
'            "End Sub". Delete and replace with this sub.
' =============================================================================
Public Sub UseExistingBatch(ByRef NoData As Boolean)
    On Error GoTo errexit

    ProtectUnProtect ("Off")
    Application.EnableEvents = False

    Dim co       As Integer
    Dim wipMonth As Date

    co       = CInt(Sheet17.Range("StartCompany").Value)
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)

    ' ------------------------------------------------------------------
    ' Step 1: Check Viewpoint budWIPDetail for existing batch rows
    ' ------------------------------------------------------------------
    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    Dim existingDepts As String
    existingDepts = ""

    If Not conn Is Nothing Then
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = conn
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "LCGWIPBatchCheck"
        cmd.CommandTimeout = 30

        cmd.Parameters.Append cmd.CreateParameter("@Co", adTinyInt, adParamInput, , co)
        cmd.Parameters.Append cmd.CreateParameter("@Month", adDate, adParamInput, , wipMonth)
        cmd.Parameters.Append cmd.CreateParameter("@rcode", adInteger, adParamOutput)
        cmd.Parameters.Append cmd.CreateParameter("@DeptList", adVarChar, adParamOutput, 200)

        cmd.Execute

        Dim rcode As Integer
        rcode = CInt(cmd.Parameters("@rcode").Value)

        If rcode = 0 Then
            ' Rows exist in budWIPDetail — batches found
            If Not IsNull(cmd.Parameters("@DeptList").Value) Then
                existingDepts = CStr(cmd.Parameters("@DeptList").Value)
            End If
        End If

        Set cmd = Nothing
    End If

    ' ------------------------------------------------------------------
    ' Step 2: If existing batches found, prompt to reopen
    ' ------------------------------------------------------------------
    If existingDepts <> "" Then
        Dim monthStr As String
        monthStr = Format(wipMonth, "mmmm yyyy")
        Dim msg As String
        msg = "WIP batches already exist for " & monthStr & "." & vbCrLf & vbCrLf & _
              "Existing departments: " & existingDepts & vbCrLf & vbCrLf & _
              "Click YES to reopen these batches." & vbCrLf & _
              "Click NO to select different departments."

        If MsgBox(msg, vbYesNo + vbQuestion, "Existing WIP Batches Found") = vbYes Then
            Sheet17.Range("StartDept").Value = existingDepts
            NoData = False
            GoTo 9999
        End If
        ' If No: fall through to Dept picker below
    End If

    ' ------------------------------------------------------------------
    ' Step 3: Show Dept picker (no existing batch, or user chose fresh)
    ' ------------------------------------------------------------------
    DeptSelectionCanceled = False

    Dim Deptfrm As Dept
    Set Deptfrm = New Dept
    Deptfrm.StartUpPosition = 0
    Deptfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Deptfrm.Width)
    Deptfrm.Top  = Application.Top  + (0.5 * Application.Height) - (0.5 * Deptfrm.Height)

    Application.EnableEvents = False
    Deptfrm.Show

    If DeptSelectionCanceled Then
        NoData = True
    Else
        NoData = False
    End If

    GoTo 9999

errexit:
    MsgBox "There was an error in the UseExistingBatch Routine. " & Err.Description, vbOKOnly

9999:
    ProtectUnProtect ("On")
    Application.EnableEvents = True

End Sub
