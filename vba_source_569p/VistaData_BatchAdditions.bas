Attribute VB_Name = "VistaData_BatchAdditions"
' =============================================================================
' VistaData_BatchAdditions.bas
' Batch management functions to APPEND to VistaData.bas
'
' DEPLOY: In the VBA editor, open VistaData module. Scroll to the very bottom.
'         Paste the three functions below after the last End Function.
'
' These functions manage the udWIPBatch table in Viewpoint (10.112.11.8),
' reusing the existing mVistaConn connection — no new DB required.
' =============================================================================

' =============================================================================
' GetExistingBatches
' Returns all batches for a given company + month from udWIPBatch.
' Called by UseExistingBatch to decide whether to show "reopen" prompt.
' Returns Nothing on error.
' =============================================================================
Public Function GetExistingBatches(co As Integer, wipMonth As Date) As ADODB.Recordset
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        Set GetExistingBatches = Nothing
        Exit Function
    End If

    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "LylesWIPBatchGet"
    cmd.CommandTimeout = 30

    cmd.Parameters.Append cmd.CreateParameter("@Co", adTinyInt, adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@Month", adDate, adParamInput, , wipMonth)

    conn.CursorLocation = adUseClient

    Dim rs As New ADODB.Recordset
    rs.CursorType = adOpenStatic
    rs.Open cmd

    Set GetExistingBatches = rs
    Exit Function

ErrorHandler:
    MsgBox "Error checking existing WIP batches:" & vbCrLf & _
           Err.Description, vbCritical, "Batch Check Error"
    Set GetExistingBatches = Nothing
End Function

' =============================================================================
' CreateWIPBatch
' Creates a new batch for a single department, or returns the existing one.
' Returns the BatchId. Sets isNew = True if created, False if already existed.
' Dept must be a single CHAR(2) value — call once per department.
' =============================================================================
Public Function CreateWIPBatch(co As Integer, wipMonth As Date, dept As String, _
                                userName As String, ByRef isNew As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        CreateWIPBatch = 0
        isNew = False
        Exit Function
    End If

    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "LylesWIPBatchCreate"
    cmd.CommandTimeout = 30

    ' Input params (must match proc parameter order exactly)
    cmd.Parameters.Append cmd.CreateParameter("@Co", adTinyInt, adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@Month", adDate, adParamInput, , wipMonth)
    cmd.Parameters.Append cmd.CreateParameter("@Dept", adChar, adParamInput, 2, Left(dept & "  ", 2))
    cmd.Parameters.Append cmd.CreateParameter("@UserName", adVarChar, adParamInput, 100, userName)
    ' Output params
    cmd.Parameters.Append cmd.CreateParameter("@BatchId", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("@IsNew", adTinyInt, adParamOutput)

    cmd.Execute

    CreateWIPBatch = CLng(cmd.Parameters("@BatchId").Value)
    isNew = (CLng(cmd.Parameters("@IsNew").Value) <> 0)
    Exit Function

ErrorHandler:
    MsgBox "Error creating WIP batch for dept " & dept & ":" & vbCrLf & _
           Err.Description, vbCritical, "Batch Create Error"
    CreateWIPBatch = 0
    isNew = False
End Function

' =============================================================================
' SetBatchState
' Advances (or sets) the batch state for one department.
' newState: "Open", "ReadyForOps", "OpsApproved", or "AcctApproved"
' Call once per department — loops are handled in the caller.
' =============================================================================
Public Sub SetBatchState(co As Integer, wipMonth As Date, dept As String, _
                          newState As String, userName As String)
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then Exit Sub

    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "LylesWIPBatchSetState"
    cmd.CommandTimeout = 30

    cmd.Parameters.Append cmd.CreateParameter("@Co", adTinyInt, adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@Month", adDate, adParamInput, , wipMonth)
    cmd.Parameters.Append cmd.CreateParameter("@Dept", adChar, adParamInput, 2, Left(dept & "  ", 2))
    cmd.Parameters.Append cmd.CreateParameter("@NewState", adVarChar, adParamInput, 20, newState)
    cmd.Parameters.Append cmd.CreateParameter("@UserName", adVarChar, adParamInput, 100, userName)

    cmd.Execute
    Exit Sub

ErrorHandler:
    MsgBox "Error setting batch state for dept " & dept & " to '" & newState & "':" & vbCrLf & _
           Err.Description, vbCritical, "Batch State Error"
End Sub
