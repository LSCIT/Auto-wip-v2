Attribute VB_Name = "LylesWIPData"
Option Explicit
' =============================================================================
' LylesWIPData.bas — LylesWIP Write-Back Module
' Purpose:  Connect to LylesWIP database on P&P server and persist override,
'           approval, and batch state data entered by users in the workbook.
'           Mirrors VistaData.bas connection pattern exactly.
' Created:  March 2026 — Sprint 4 (write path build)
' Server:   PPServerName (10.103.30.11, Cloud-Apps1)
' Database: PPDBName (LylesWIP)
' Auth:     PPUsername / PPPassword from Settings sheet named ranges
' =============================================================================

' Module-level LylesWIP connection
Private mWIPConn As ADODB.Connection

' =============================================================================
' OpenWIPConnection
' Opens ADODB connection to the LylesWIP database using Settings sheet config
' =============================================================================
Public Function OpenWIPConnection() As Boolean
    On Error GoTo ErrorHandler

    ' Reuse if already open
    If Not mWIPConn Is Nothing Then
        If mWIPConn.State = adStateOpen Then
            OpenWIPConnection = True
            Exit Function
        End If
    End If

    Set mWIPConn = New ADODB.Connection

    Dim server As String
    server = CStr(Sheet2.Range("PPServerName").Value)

    If server = "" Then
        MsgBox "LylesWIP server not configured in Settings sheet (PPServerName).", vbCritical
        OpenWIPConnection = False
        Exit Function
    End If

    Dim dbName As String
    dbName = ""
    On Error Resume Next
    dbName = CStr(Sheet2.Range("WIPDBName").Value)
    On Error GoTo ErrorHandler
    If dbName = "" Then dbName = "LylesWIP"

    Dim uid As String
    Dim pwd As String
    uid = ""
    pwd = ""
    On Error Resume Next
    uid = CStr(Sheet2.Range("PPUsername").Value)
    pwd = CStr(Sheet2.Range("PPPassword").Value)
    On Error GoTo ErrorHandler

    Dim connStr As String
    If uid <> "" And pwd <> "" Then
        connStr = "Provider=MSOLEDBSQL;" & _
                  "Server=" & server & ";" & _
                  "Database=" & dbName & ";" & _
                  "UID=" & uid & ";" & _
                  "PWD=" & pwd & ";" & _
                  "Persist Security Info=False;" & _
                  "TrustServerCertificate=yes;" & _
                  "Packet Size=4096;"
    Else
        connStr = "Provider=MSOLEDBSQL;" & _
                  "Server=" & server & ";" & _
                  "Database=" & dbName & ";" & _
                  "Integrated Security=SSPI;" & _
                  "Persist Security Info=False;" & _
                  "TrustServerCertificate=yes;" & _
                  "Packet Size=4096;"
    End If

    mWIPConn.ConnectionString = connStr
    mWIPConn.CommandTimeout = 60
    mWIPConn.Open

    OpenWIPConnection = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to connect to LylesWIP server (" & server & "):" & vbCrLf & _
           Err.Description, vbCritical, "LylesWIP Connection Error"
    OpenWIPConnection = False
End Function

' =============================================================================
' CloseWIPConnection
' Cleanly closes the LylesWIP connection
' =============================================================================
Public Sub CloseWIPConnection()
    On Error Resume Next
    If Not mWIPConn Is Nothing Then
        If mWIPConn.State = adStateOpen Then
            mWIPConn.Close
        End If
        Set mWIPConn = Nothing
    End If
End Sub

' =============================================================================
' GetWIPConnection
' Returns active LylesWIP connection (opens if needed)
' =============================================================================
Public Function GetWIPConnection() As ADODB.Connection
    ' VBA does NOT short-circuit Or/And — must use nested If to avoid
    ' accessing .State on a Nothing object
    Dim needsConnection As Boolean
    needsConnection = False

    If mWIPConn Is Nothing Then
        needsConnection = True
    ElseIf mWIPConn.State <> adStateOpen Then
        needsConnection = True
    End If

    If needsConnection Then
        If Not OpenWIPConnection() Then
            Set GetWIPConnection = Nothing
            Exit Function
        End If
    End If
    Set GetWIPConnection = mWIPConn
End Function

' =============================================================================
' TestWIPConnection
' Quick connectivity test — callable from Settings sheet button
' =============================================================================
Public Sub TestWIPConnection()
    If OpenWIPConnection() Then
        Dim rs As ADODB.Recordset
        Set rs = mWIPConn.Execute("SELECT @@SERVERNAME AS ServerName, DB_NAME() AS DatabaseName, COUNT(*) AS BatchCount FROM dbo.WipBatches")
        MsgBox "Connected to LylesWIP successfully!" & vbCrLf & _
               "Server: " & rs.Fields("ServerName").Value & vbCrLf & _
               "Database: " & rs.Fields("DatabaseName").Value & vbCrLf & _
               "Batches on file: " & rs.Fields("BatchCount").Value, _
               vbInformation, "LylesWIP Connection Test"
        rs.Close
        Set rs = Nothing
        CloseWIPConnection
    End If
End Sub

' =============================================================================
' CreateBatch
' Creates a new WIP batch for the given company/month/department.
' If the batch already exists, does nothing (proc uses IF NOT EXISTS).
' Returns the batch Id, or -1 on error.
' =============================================================================
Public Function CreateBatch(co As Integer, wipMonth As Date, dept As String, createdBy As String) As Long
    CreateBatch = -1

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.LylesWIPCreateBatch"

    cmd.Parameters.Append cmd.CreateParameter("@JCCo",      adTinyInt,  adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@WipMonth",  adDBDate,   adParamInput, , wipMonth)
    cmd.Parameters.Append cmd.CreateParameter("@Department",adVarChar,  adParamInput, 10, dept)
    cmd.Parameters.Append cmd.CreateParameter("@CreatedBy", adVarChar,  adParamInput, 100, createdBy)

    Dim rs As ADODB.Recordset
    Set rs = cmd.Execute()

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            CreateBatch = CLng(rs.Fields("Id").Value)
        End If
        rs.Close
        Set rs = Nothing
    End If

    Set cmd = Nothing
    Exit Function

ErrorHandler:
    MsgBox "CreateBatch failed:" & vbCrLf & Err.Description, vbCritical, "LylesWIP Error"
    CreateBatch = -1
End Function

' =============================================================================
' GetBatchState
' Returns the BatchState string for the given company/month/department.
' Returns "" if no batch exists or on error.
' Possible values: "Open", "ReadyForOps", "OpsApproved", "AcctApproved"
' =============================================================================
Public Function GetBatchState(co As Integer, wipMonth As Date, dept As String) As String
    GetBatchState = ""

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.LylesWIPCheckBatchState"

    cmd.Parameters.Append cmd.CreateParameter("@JCCo",      adTinyInt, adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@WipMonth",  adDBDate,  adParamInput, , wipMonth)
    cmd.Parameters.Append cmd.CreateParameter("@Department",adVarChar, adParamInput, 10, dept)

    Dim rs As ADODB.Recordset
    Set rs = cmd.Execute()

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            Dim stateVal As Variant
            stateVal = rs.Fields("BatchState").Value
            If Not IsNull(stateVal) Then GetBatchState = CStr(stateVal)
        End If
        rs.Close
        Set rs = Nothing
    End If

    Set cmd = Nothing
    Exit Function

ErrorHandler:
    ' No batch exists — return ""
    GetBatchState = ""
End Function

' =============================================================================
' UpdateBatchState
' Advances the batch state machine:
'   Open → ReadyForOps → OpsApproved → AcctApproved
' Returns True on success, False on error.
' =============================================================================
Public Function UpdateBatchState(co As Integer, wipMonth As Date, dept As String, _
                                  newState As String, changedBy As String) As Boolean
    UpdateBatchState = False

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.LylesWIPUpdateBatchState"

    cmd.Parameters.Append cmd.CreateParameter("@JCCo",      adTinyInt, adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@WipMonth",  adDBDate,  adParamInput, , wipMonth)
    cmd.Parameters.Append cmd.CreateParameter("@Department",adVarChar, adParamInput, 10, dept)
    cmd.Parameters.Append cmd.CreateParameter("@NewState",  adVarChar, adParamInput, 20, newState)
    cmd.Parameters.Append cmd.CreateParameter("@ChangedBy", adVarChar, adParamInput, 100, changedBy)

    cmd.Execute
    Set cmd = Nothing

    UpdateBatchState = True
    Exit Function

ErrorHandler:
    MsgBox "UpdateBatchState failed (Co=" & co & ", " & dept & ", " & newState & "):" & _
           vbCrLf & Err.Description, vbCritical, "LylesWIP Error"
    UpdateBatchState = False
End Function

' =============================================================================
' AllBatchesApproved
' Returns True if every batch for the given Co/Month is in AcctApproved state.
' Used to gate the December year-end snapshot — only fire when ALL divisions
' have completed the full 3-stage workflow.
' =============================================================================
Public Function AllBatchesApproved(co As Integer, wipMonth As Date) As Boolean
    AllBatchesApproved = False

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    Dim rs As ADODB.Recordset
    Set rs = conn.Execute( _
        "SELECT COUNT(*) AS Total, " & _
        "SUM(CASE WHEN BatchState = 'AcctApproved' THEN 1 ELSE 0 END) AS Approved " & _
        "FROM dbo.WipBatches " & _
        "WHERE JCCo = " & co & " AND WipMonth = '" & Format(wipMonth, "yyyy-mm-dd") & "'")

    If Not rs.EOF Then
        Dim total As Long
        Dim approved As Long
        total = CLng(rs.Fields("Total").Value)
        approved = CLng(rs.Fields("Approved").Value)
        AllBatchesApproved = (total > 0 And total = approved)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    AllBatchesApproved = False
End Function

' =============================================================================
' DistributeToOps
' Saves a macro-enabled copy (.xlsm) to C:\Trusted\ for Ops PMs.
' Sets ClearFormOnOpen=False so data survives reopen.
' Called from a "Save & Distribute" button on the Start sheet.
'
' The copy retains all loaded Vista data + LylesWIP overrides.
' Ops opens it, edits yellow columns, double-clicks Done → saves to LylesWIP.
' Ops does NOT need Vista access — only P&P (LylesWIP).
' =============================================================================
Public Sub DistributeToOps()
    On Error GoTo ErrorHandler

    ' Verify batch is ReadyForOps or later
    Dim co As Integer
    Dim wipMonth As Date
    Dim dept As String
    co = CInt(Sheet17.Range("StartCompany").Value)
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)
    dept = CStr(Sheet17.Range("StartDept").Value)

    If co = 0 Or dept = "" Then
        MsgBox "Load data first (Company, Month, Division).", vbExclamation, "Distribute"
        Exit Sub
    End If

    Dim batchState As String
    batchState = GetBatchState(co, wipMonth, dept)
    If batchState <> "ReadyForOps" And batchState <> "OpsApproved" Then
        MsgBox "Batch must be Ready for Ops before distributing." & vbCrLf & _
               "Current state: " & IIf(batchState = "", "No batch", batchState), _
               vbExclamation, "Distribute"
        Exit Sub
    End If

    ' Build output filename: WIP Schedule - 15 Div51 Dec2025 20260408-1415.xlsm
    Dim monthName As String
    monthName = Format(wipMonth, "mmmyyyy")
    Dim timeStamp As String
    timeStamp = Format(Now, "yyyymmdd-HHnn")
    Dim outName As String
    outName = "WIP Schedule - " & co & " Div" & dept & " " & monthName & " " & timeStamp & ".xlsm"
    Dim outPath As String
    outPath = "C:\Trusted\" & outName

    ' Confirm with user
    If MsgBox("Save distributed copy for Ops?" & vbCrLf & vbCrLf & _
              outPath, vbYesNo + vbQuestion, "Distribute to Ops") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Set ClearFormOnOpen=False so data is preserved when Ops opens the file
    Dim origClearForm As Variant
    On Error Resume Next
    origClearForm = Sheet2.Range("ClearFormOnOpen").Value
    On Error GoTo ErrorHandler
    Sheet2.Unprotect "password"
    Sheet2.Range("ClearFormOnOpen").Value = False

    ' Remember the master file path to restore after
    Dim masterPath As String
    masterPath = ThisWorkbook.FullName

    ' SaveAs .xlsm (xlOpenXMLWorkbookMacroEnabled = 52)
    ThisWorkbook.SaveAs outPath, FileFormat:=52

    ' Restore: save back as .xltm master with ClearFormOnOpen=True
    Sheet2.Range("ClearFormOnOpen").Value = True
    ThisWorkbook.SaveAs masterPath, FileFormat:=53

    Sheet2.Protect "password"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Distributed copy saved:" & vbCrLf & outPath, vbInformation, "Distribute to Ops"
    Exit Sub

ErrorHandler:
    ' Restore ClearFormOnOpen on error
    On Error Resume Next
    Sheet2.Range("ClearFormOnOpen").Value = True
    Sheet2.Protect "password"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Distribution failed:" & vbCrLf & Err.Description, vbCritical, "Distribute Error"
End Sub

' =============================================================================
' CopyOpsToGAAP
' Copies Ops override values to GAAP override columns in WipJobData for jobs
' currently loaded on Jobs-Ops (Sheet11) — scoped to the active division.
' Only copies where GAAP overrides are currently NULL (won't overwrite).
' Called from "Copy Ops to GAAP" button on Start sheet.
' Returns number of rows copied, or -1 on error.
' =============================================================================
Public Function CopyOpsToGAAP() As Long
    CopyOpsToGAAP = -1

    Dim co As Integer
    Dim wipMonth As Date
    co = CInt(Sheet17.Range("StartCompany").Value)
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)

    If co = 0 Then
        MsgBox "Load data first.", vbExclamation, "Copy Ops to GAAP"
        Exit Function
    End If

    ' Build job list from Jobs-Ops sheet (only copy for jobs in this division)
    If NumDict Is Nothing Then
        InitializeColumnDictionaries NumDict, LetDict, 1
    End If

    Dim summaryRange As Range
    Set summaryRange = Sheet11.Range("SummaryData")
    Dim jnCol As Long
    jnCol = summaryRange.Cells(1, NumDict(Sheet11.CodeName)("COLJobNumber")).Column
    Dim lastRow As Long
    lastRow = Sheet11.Cells(Sheet11.Rows.Count, jnCol).End(xlUp).Row
    Dim totalRows As Long
    totalRows = Application.Max(summaryRange.Rows.Count, lastRow - summaryRange.Row + 1)

    Dim jobList As String
    jobList = ""
    Dim r As Long
    For r = 1 To totalRows
        Dim jobNum As String
        jobNum = CStr(summaryRange.Cells(r, NumDict(Sheet11.CodeName)("COLJobNumber")).Value)
        If jobNum <> "" Then
            If jobList <> "" Then jobList = jobList & ","
            jobList = jobList & "'" & Replace(jobNum, "'", "''") & "'"
        End If
    Next r

    If jobList = "" Then
        MsgBox "No jobs found on Jobs-Ops sheet.", vbExclamation, "Copy Ops to GAAP"
        Exit Function
    End If

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    ' Direct SQL scoped to jobs on the current sheet
    Dim sql As String
    sql = "UPDATE dbo.WipJobData " & _
          "SET GAAPRevOverride = OpsRevOverride, " & _
          "    GAAPRevPlugged = OpsRevPlugged, " & _
          "    GAAPCostOverride = OpsCostOverride, " & _
          "    GAAPCostPlugged = OpsCostPlugged, " & _
          "    UpdatedAt = GETDATE() " & _
          "WHERE JCCo = " & co & " " & _
          "  AND WipMonth = '" & Format(wipMonth, "yyyy-mm-dd") & "' " & _
          "  AND Job IN (" & jobList & ") " & _
          "  AND GAAPRevOverride IS NULL " & _
          "  AND GAAPCostOverride IS NULL " & _
          "  AND (OpsRevOverride IS NOT NULL OR OpsCostOverride IS NOT NULL)"

    conn.Execute sql

    ' Get count of affected rows
    Dim rsCount As ADODB.Recordset
    Set rsCount = conn.Execute("SELECT @@ROWCOUNT AS RowsCopied")
    Dim rowsCopied As Long
    rowsCopied = CLng(rsCount.Fields("RowsCopied").Value)
    rsCount.Close
    Set rsCount = Nothing

    MsgBox rowsCopied & " job(s) — Ops overrides copied to GAAP as starting values." & vbCrLf & _
           "Jobs with existing GAAP overrides were not changed.", _
           vbInformation, "Copy Ops to GAAP"

    CopyOpsToGAAP = rowsCopied
    Exit Function

ErrorHandler:
    MsgBox "CopyOpsToGAAP failed:" & vbCrLf & Err.Description, vbCritical, "LylesWIP Error"
    CopyOpsToGAAP = -1
End Function

' =============================================================================
' SaveJobRow
' Reads all override/flag fields from the given row on sheet sh and writes
' them to LylesWIP via dbo.LylesWIPSaveJobRow.
'
' Parameters:
'   sh       — the worksheet (Jobs-Ops = Sheet11, Jobs-GAAP = Sheet12)
'   r        — row number
'   co       — JCCo (company number)
'   wipMonth — batch month (Date)
'   userName — current user name
'
' Returns True on success.
' =============================================================================
Public Function SaveJobRow(sh As Worksheet, r As Long, co As Integer, _
                            wipMonth As Date, userName As String) As Boolean
    SaveJobRow = False

    ' Require column dictionary
    If NumDict Is Nothing Then
        InitializeColumnDictionaries NumDict, LetDict, 1
    End If

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    ' Read job number from the Job column (COLJobNumber is the dictionary key)
    Dim job As String
    job = CStr(sh.Cells(r, NumDict(sh.CodeName)("COLJobNumber")).Value)
    If job = "" Then Exit Function

    ' -------------------------------------------------------------------------
    ' Read Ops override fields.
    ' Plugged flag = Font.Bold on the COLZ hidden cell — no separate "plugged" column exists.
    ' Module6 helpers (GetOpsRevPlug, GetOpsRev, etc.) encapsulate the "Chg" buffer logic:
    '   if COLZOPsRChg="T" → user edited this session → read from COLZOPsRevNew (bold)
    '   otherwise          → read from COLZOPsRev (loaded value, bold if previously plugged)
    ' -------------------------------------------------------------------------
    Dim opsRevOverride  As Variant  ' Null if not plugged
    Dim opsCostOverride As Variant
    Dim opsRevPlugged   As Boolean
    Dim opsCostPlugged  As Boolean

    ' Guard: GetOps* helpers access Font.Bold on COLZ cells which can fail
    ' on subtotal rows, error cells, or sheets missing the column.
    On Error Resume Next
    opsRevPlugged  = (GetOpsRevPlug(r, sh) = "Y")
    opsCostPlugged = (GetOpsCostPlug(r, sh) = "Y")
    On Error GoTo ErrorHandler

    If opsRevPlugged Then
        On Error Resume Next
        opsRevOverride = ToDecimalOrNull(GetOpsRev(r, sh))
        On Error GoTo ErrorHandler
    Else
        opsRevOverride = Null
    End If

    If opsCostPlugged Then
        On Error Resume Next
        opsCostOverride = ToDecimalOrNull(GetOpsCost(r, sh))
        On Error GoTo ErrorHandler
    Else
        opsCostOverride = Null
    End If

    ' -------------------------------------------------------------------------
    ' Read GAAP override fields (same Font.Bold pattern via COLZJCOR / COLZJCOP flags)
    ' -------------------------------------------------------------------------
    Dim gaapRevOverride  As Variant
    Dim gaapCostOverride As Variant
    Dim gaapRevPlugged   As Boolean
    Dim gaapCostPlugged  As Boolean

    On Error Resume Next
    gaapRevPlugged  = (GetGAAPRevPlug(r, sh) = "Y")
    gaapCostPlugged = (GetGAAPCostPlug(r, sh) = "Y")
    On Error GoTo ErrorHandler

    If gaapRevPlugged Then
        On Error Resume Next
        gaapRevOverride = ToDecimalOrNull(GetGAAPRev(r, sh))
        On Error GoTo ErrorHandler
    Else
        gaapRevOverride = Null
    End If

    If gaapCostPlugged Then
        On Error Resume Next
        gaapCostOverride = ToDecimalOrNull(GetGAAPCost(r, sh))
        On Error GoTo ErrorHandler
    Else
        gaapCostOverride = Null
    End If

    ' -------------------------------------------------------------------------
    ' Read bonus, notes, completion date, and status flags.
    ' Bonus plugged = Font.Bold on COLZOPsBonusNew (see Module6.GetOpsBonusPlug).
    ' -------------------------------------------------------------------------
    Dim bonusProfit    As Variant
    Dim compDate       As Variant
    Dim isClosed       As Boolean
    Dim isOpsDone      As Boolean
    Dim isGAAPDone     As Boolean

    Dim opsRevNotes  As String
    Dim opsCostNotes As String
    Dim gaapRevNotes As String
    Dim gaapCostNotes As String

    ' COLZOPsBonusNew only exists on the Ops sheet.
    ' On GAAP sheet, pass Null so the stored proc preserves the existing BonusProfit.
    bonusProfit = Null
    If sh.CodeName = "Sheet11" Then
        On Error Resume Next
        If GetOpsBonusPlug(r, sh) = "Y" Then
            bonusProfit = ToDecimalOrNull(sh.Cells(r, NumDict(sh.CodeName)("COLZOPsBonusNew")).Value)
        End If
        On Error GoTo ErrorHandler
    End If
    opsRevNotes   = CStr(sh.Cells(r, NumDict(sh.CodeName)("COLZOPsRevNotes")).Value)
    opsCostNotes  = CStr(sh.Cells(r, NumDict(sh.CodeName)("COLZOPsCostNotes")).Value)
    gaapRevNotes  = CStr(sh.Cells(r, NumDict(sh.CodeName)("COLZGAAPRevNotes")).Value)
    gaapCostNotes = CStr(sh.Cells(r, NumDict(sh.CodeName)("COLZGAAPCostNotes")).Value)
    On Error GoTo ErrorHandler

    Dim compDateRaw As Variant
    compDateRaw = sh.Cells(r, NumDict(sh.CodeName)("COLCompDate")).Value
    If IsDate(compDateRaw) Then
        compDate = CDate(compDateRaw)
    Else
        compDate = Null
    End If

    ' COLClose uses "C" (not "Y"); COLDone and COLGAAPDone use "P"
    isClosed  = (sh.Cells(r, NumDict(sh.CodeName)("COLClose")).Value = "C")
    isOpsDone = (sh.Cells(r, NumDict(sh.CodeName)("COLDone")).Value = "P")

    ' COLGAAPDone only exists on the GAAP sheet; safely default to False on Ops
    On Error Resume Next
    isGAAPDone = (sh.Cells(r, NumDict(sh.CodeName)("COLGAAPDone")).Value = "P")
    On Error GoTo ErrorHandler

    ' -------------------------------------------------------------------------
    ' Execute stored proc
    ' -------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.LylesWIPSaveJobRow"

    cmd.Parameters.Append cmd.CreateParameter("@JCCo",           adTinyInt,  adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@Job",            adVarChar,  adParamInput, 50, job)
    cmd.Parameters.Append cmd.CreateParameter("@WipMonth",       adDBDate,   adParamInput, , wipMonth)
    cmd.Parameters.Append cmd.CreateParameter("@OpsRevOverride", adDecimal,  adParamInput, , opsRevOverride)
    cmd.Parameters.Item("@OpsRevOverride").NumericScale = 2
    cmd.Parameters.Item("@OpsRevOverride").Precision = 17
    cmd.Parameters.Append cmd.CreateParameter("@OpsRevPlugged",  adBoolean,  adParamInput, , opsRevPlugged)
    cmd.Parameters.Append cmd.CreateParameter("@GAAPRevOverride",adDecimal,  adParamInput, , gaapRevOverride)
    cmd.Parameters.Item("@GAAPRevOverride").NumericScale = 2
    cmd.Parameters.Item("@GAAPRevOverride").Precision = 17
    cmd.Parameters.Append cmd.CreateParameter("@GAAPRevPlugged", adBoolean,  adParamInput, , gaapRevPlugged)
    cmd.Parameters.Append cmd.CreateParameter("@OpsCostOverride",adDecimal,  adParamInput, , opsCostOverride)
    cmd.Parameters.Item("@OpsCostOverride").NumericScale = 2
    cmd.Parameters.Item("@OpsCostOverride").Precision = 17
    cmd.Parameters.Append cmd.CreateParameter("@OpsCostPlugged", adBoolean,  adParamInput, , opsCostPlugged)
    cmd.Parameters.Append cmd.CreateParameter("@GAAPCostOverride",adDecimal, adParamInput, , gaapCostOverride)
    cmd.Parameters.Item("@GAAPCostOverride").NumericScale = 2
    cmd.Parameters.Item("@GAAPCostOverride").Precision = 17
    cmd.Parameters.Append cmd.CreateParameter("@GAAPCostPlugged",adBoolean,  adParamInput, , gaapCostPlugged)
    cmd.Parameters.Append cmd.CreateParameter("@BonusProfit",    adDecimal,  adParamInput, , bonusProfit)
    cmd.Parameters.Item("@BonusProfit").NumericScale = 2
    cmd.Parameters.Item("@BonusProfit").Precision = 17
    cmd.Parameters.Append cmd.CreateParameter("@OpsRevNotes",    adVarChar,  adParamInput, 500, opsRevNotes)
    cmd.Parameters.Append cmd.CreateParameter("@GAAPRevNotes",   adVarChar,  adParamInput, 500, gaapRevNotes)
    cmd.Parameters.Append cmd.CreateParameter("@OpsCostNotes",   adVarChar,  adParamInput, 500, opsCostNotes)
    cmd.Parameters.Append cmd.CreateParameter("@GAAPCostNotes",  adVarChar,  adParamInput, 500, gaapCostNotes)
    cmd.Parameters.Append cmd.CreateParameter("@CompletionDate", adDBDate,   adParamInput, , compDate)
    cmd.Parameters.Append cmd.CreateParameter("@IsClosed",       adBoolean,  adParamInput, , isClosed)
    cmd.Parameters.Append cmd.CreateParameter("@IsOpsDone",      adBoolean,  adParamInput, , isOpsDone)
    cmd.Parameters.Append cmd.CreateParameter("@IsGAAPDone",     adBoolean,  adParamInput, , isGAAPDone)
    cmd.Parameters.Append cmd.CreateParameter("@UserName",       adVarChar,  adParamInput, 100, userName)

    cmd.Execute
    Set cmd = Nothing

    SaveJobRow = True
    Exit Function

ErrorHandler:
    MsgBox "SaveJobRow failed (Job=" & job & ", row=" & r & "):" & _
           vbCrLf & Err.Description, vbCritical, "LylesWIP Error"
    SaveJobRow = False
End Function

' =============================================================================
' SaveYearEndSnapshot
' Called at December AcctApproved. Reads all WipJobData rows for the given
' Co/Month and writes each to WipYearEndSnapshot via the stored proc.
' SnapshotYear = Year(wipMonth) (e.g., 2025 for Dec 2025).
' Returns the number of rows written, or -1 on error.
' =============================================================================
Public Function SaveYearEndSnapshot(co As Integer, wipMonth As Date) As Long
    SaveYearEndSnapshot = -1

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    ' Read all approved jobs for this Co/Month from WipJobData
    Dim rsJobs As ADODB.Recordset
    Set rsJobs = New ADODB.Recordset
    rsJobs.Open "SELECT Job, GAAPRevOverride, GAAPCostOverride, " & _
                "OpsRevOverride, OpsCostOverride, BonusProfit " & _
                "FROM dbo.WipJobData " & _
                "WHERE JCCo = " & co & " AND WipMonth = '" & Format(wipMonth, "yyyy-mm-dd") & "'", _
                conn, adOpenForwardOnly, adLockReadOnly

    Dim snapshotYear As Integer
    snapshotYear = Year(wipMonth)

    Dim rowCount As Long
    rowCount = 0

    Do While Not rsJobs.EOF
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = conn
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "dbo.LylesWIPSaveYearEndSnapshot"

        cmd.Parameters.Append cmd.CreateParameter("@JCCo", adTinyInt, adParamInput, , co)
        cmd.Parameters.Append cmd.CreateParameter("@Job", adVarChar, adParamInput, 50, _
                              CStr(rsJobs.Fields("Job").Value))
        cmd.Parameters.Append cmd.CreateParameter("@SnapshotYear", adSmallInt, adParamInput, , snapshotYear)

        cmd.Parameters.Append cmd.CreateParameter("@PriorYearGAAPRev", adDecimal, adParamInput, , _
                              rsJobs.Fields("GAAPRevOverride").Value)
        cmd.Parameters.Item("@PriorYearGAAPRev").NumericScale = 2
        cmd.Parameters.Item("@PriorYearGAAPRev").Precision = 17

        cmd.Parameters.Append cmd.CreateParameter("@PriorYearGAAPCost", adDecimal, adParamInput, , _
                              rsJobs.Fields("GAAPCostOverride").Value)
        cmd.Parameters.Item("@PriorYearGAAPCost").NumericScale = 2
        cmd.Parameters.Item("@PriorYearGAAPCost").Precision = 17

        cmd.Parameters.Append cmd.CreateParameter("@PriorYearOpsRev", adDecimal, adParamInput, , _
                              rsJobs.Fields("OpsRevOverride").Value)
        cmd.Parameters.Item("@PriorYearOpsRev").NumericScale = 2
        cmd.Parameters.Item("@PriorYearOpsRev").Precision = 17

        cmd.Parameters.Append cmd.CreateParameter("@PriorYearOpsCost", adDecimal, adParamInput, , _
                              rsJobs.Fields("OpsCostOverride").Value)
        cmd.Parameters.Item("@PriorYearOpsCost").NumericScale = 2
        cmd.Parameters.Item("@PriorYearOpsCost").Precision = 17

        cmd.Parameters.Append cmd.CreateParameter("@BonusProfit", adDecimal, adParamInput, , _
                              rsJobs.Fields("BonusProfit").Value)
        cmd.Parameters.Item("@BonusProfit").NumericScale = 2
        cmd.Parameters.Item("@BonusProfit").Precision = 17

        cmd.Execute
        Set cmd = Nothing

        rowCount = rowCount + 1
        rsJobs.MoveNext
    Loop

    rsJobs.Close
    Set rsJobs = Nothing

    SaveYearEndSnapshot = rowCount
    Exit Function

ErrorHandler:
    MsgBox "SaveYearEndSnapshot failed (Co=" & co & ", Year=" & snapshotYear & "):" & _
           vbCrLf & Err.Description, vbCritical, "LylesWIP Error"
    SaveYearEndSnapshot = -1
End Function

' =============================================================================
' GetJobOverrides
' Returns all override rows for a company/month as an open ADODB.Recordset.
' Caller is responsible for closing the recordset.
' Used by GetWIPDetailData to merge persisted overrides over Vista-calculated values.
'
' Recordset fields (from LylesWIPGetJobOverrides):
'   Job, OpsRevOverride, OpsRevPlugged, GAAPRevOverride, GAAPRevPlugged,
'   OpsCostOverride, OpsCostPlugged, GAAPCostOverride, GAAPCostPlugged,
'   BonusProfit, OpsRevNotes, GAAPRevNotes, OpsCostNotes, GAAPCostNotes,
'   CompletionDate, IsClosed, IsOpsDone, IsGAAPDone, UserName, UpdatedAt
' =============================================================================
Public Function GetJobOverrides(co As Integer, wipMonth As Date) As ADODB.Recordset
    Set GetJobOverrides = Nothing

    Dim conn As ADODB.Connection
    Set conn = GetWIPConnection()
    If conn Is Nothing Then Exit Function

    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.LylesWIPGetJobOverrides"

    cmd.Parameters.Append cmd.CreateParameter("@JCCo",    adTinyInt, adParamInput, , co)
    cmd.Parameters.Append cmd.CreateParameter("@WipMonth",adDBDate,  adParamInput, , wipMonth)

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient        ' client-side cursor so we can use rs.Find
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    ' Disconnect from connection — caller can use rs after conn closes
    Set rs.ActiveConnection = Nothing
    Set cmd = Nothing

    Set GetJobOverrides = rs
    Exit Function

ErrorHandler:
    MsgBox "GetJobOverrides failed (Co=" & co & "):" & vbCrLf & _
           Err.Description, vbCritical, "LylesWIP Error"
    Set GetJobOverrides = Nothing
End Function

' =============================================================================
' MergeOverridesOntoSheet
' Called once per sheet after GetWipDetail2 finishes writing Vista values.
' Loads saved LylesWIP overrides for co/wipMonth into a Dictionary (O(1) lookup),
' then does a single pass over SummaryData rows — applying overrides wherever
' a matching job is found.
'
' Override application rules (mirror the original load logic in GetWIPDetailData):
'   Ops sheet  (Sheet11): visible OpsRev/OpsCost overrides → COLOvrRevProj/COLOvrCostProj
'   GAAP sheet (Sheet12): visible GAAPRev/GAAPCost overrides → COLOvrRevProj/COLOvrCostProj
'   COLZ backing cells get the value with Font.Bold=True so GetOpsRevPlug etc. return "Y"
'   "Chg" buffer flags are cleared (no pending edit — this is a persisted restore)
' =============================================================================
Public Sub MergeOverridesOntoSheet(sh As Worksheet, co As Integer, wipMonth As Date)
    Dim overrides As Object
    Set overrides = BuildOverrideLookup(co, wipMonth)
    If overrides Is Nothing Then Exit Sub
    If overrides.Count = 0 Then Exit Sub      ' nothing saved yet — leave Vista values as-is

    If NumDict Is Nothing Then
        InitializeColumnDictionaries NumDict, LetDict, 1
    End If

    Dim isOps As Boolean
    isOps = (sh.CodeName <> "Sheet12")

    Dim summaryRange As Range
    Set summaryRange = sh.Range("SummaryData")

    ' SummaryData may be sized smaller than the actual data written by GetWipDetail2.
    ' Scan column A (COLJobNumber) from the range start to find the real last data row.
    Dim jnColAbs As Long
    jnColAbs = summaryRange.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Column
    Dim lastDataRow As Long
    lastDataRow = sh.Cells(sh.Rows.Count, jnColAbs).End(xlUp).Row
    Dim totalRows As Long
    totalRows = Application.Max(summaryRange.Rows.Count, lastDataRow - summaryRange.Row + 1)

    Dim r As Long
    For r = 1 To totalRows
        Dim jobNum As String
        jobNum = CStr(summaryRange.Cells(r, NumDict(sh.CodeName)("COLJobNumber")).Value)
        If jobNum = "" Then GoTo NextRow          ' subtotal or empty row
        If Right(jobNum, 1) <> "." Then jobNum = jobNum & "."  ' phase jobs: "56.1010.01" → "56.1010.01."
        If Not overrides.Exists(jobNum) Then GoTo NextRow

        Dim ov() As Variant
        ov = overrides(jobNum)
        ' ov() index map (must match BuildOverrideLookup):
        ' 0=OpsRevOverride  1=OpsRevPlugged
        ' 2=GAAPRevOverride 3=GAAPRevPlugged
        ' 4=OpsCostOverride 5=OpsCostPlugged
        ' 6=GAAPCostOverride 7=GAAPCostPlugged
        ' 8=BonusProfit
        ' 9=OpsRevNotes 10=GAAPRevNotes 11=OpsCostNotes 12=GAAPCostNotes
        ' 13=CompletionDate 14=IsClosed 15=IsOpsDone 16=IsGAAPDone
        ' 17=UserName 18=UpdatedAt

        On Error Resume Next   ' guard against missing COLZ columns on a given sheet

        ' ---- Ops Rev ----
        If CBool(ov(1)) And Not IsNull(ov(0)) Then
            Dim opsRevFrom As Variant
            If isOps Then opsRevFrom = summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")).Value
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsRev")).Value = ov(0)
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsRev")).Font.Bold = True
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsRChg")).Value = ""
            If isOps Then
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")).Value = ov(0)
                AddAuditComment summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")), _
                    opsRevFrom, ov(0), ov(17), ov(18)
            End If
        End If

        ' ---- Ops Cost ----
        If CBool(ov(5)) And Not IsNull(ov(4)) Then
            Dim opsCostFrom As Variant
            If isOps Then opsCostFrom = summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")).Value
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsCost")).Value = ov(4)
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsCost")).Font.Bold = True
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsCChg")).Value = ""
            If isOps Then
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")).Value = ov(4)
                AddAuditComment summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")), _
                    opsCostFrom, ov(4), ov(17), ov(18)
            End If
        End If

        ' ---- GAAP Rev ----
        If CBool(ov(3)) And Not IsNull(ov(2)) Then
            Dim gaapRevFrom As Variant
            If Not isOps Then gaapRevFrom = summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")).Value
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZGAAPRev")).Value = ov(2)
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZGAAPRev")).Font.Bold = True
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZJCOR")).Value = ""
            If Not isOps Then
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")).Value = ov(2)
                AddAuditComment summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")), _
                    gaapRevFrom, ov(2), ov(17), ov(18)
            End If
        End If

        ' ---- GAAP Cost ----
        If CBool(ov(7)) And Not IsNull(ov(6)) Then
            Dim gaapCostFrom As Variant
            If Not isOps Then gaapCostFrom = summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")).Value
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZGAAPCost")).Value = ov(6)
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZGAAPCost")).Font.Bold = True
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZJCOP")).Value = ""
            If Not isOps Then
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")).Value = ov(6)
                AddAuditComment summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")), _
                    gaapCostFrom, ov(6), ov(17), ov(18)
            End If
        End If

        ' ---- Bonus Profit (Ops only — GAAP tab doesn't show bonus) ----
        If isOps Then
            If Not IsNull(ov(8)) And CDbl(ov(8)) <> 0 Then
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsBonus")).Value = ov(8)
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsBonus")).Font.Bold = True
                ' Also populate COLZOPsBonusNew so GetOpsBonusPlug returns "Y"
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsBonusNew")).Value = ov(8)
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsBonusNew")).Font.Bold = True
                summaryRange.Cells(r, NumDict(sh.CodeName)("COLJTDBonusProfit")).Value = ov(8)
            End If
        End If

        ' ---- Notes ----
        If Not IsNull(ov(9))  And CStr(ov(9))  <> "" Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsRevNotes")).Value  = ov(9)
        If Not IsNull(ov(10)) And CStr(ov(10)) <> "" Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLZGAAPRevNotes")).Value = ov(10)
        If Not IsNull(ov(11)) And CStr(ov(11)) <> "" Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLZOPsCostNotes")).Value = ov(11)
        If Not IsNull(ov(12)) And CStr(ov(12)) <> "" Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLZGAAPCostNotes")).Value = ov(12)

        ' ---- Completion Date ----
        If Not IsNull(ov(13)) Then
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLCompDate")).Value = ov(13)
        End If

        ' ---- Status flags ----
        If isOps Then
            If CBool(ov(14)) Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLClose")).Value = "C"
            If CBool(ov(15)) Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLDone")).Value  = "P"
        Else
            If CBool(ov(16)) Then summaryRange.Cells(r, NumDict(sh.CodeName)("COLGAAPDone")).Value = "P"
        End If

        ' ---- UserName ----
        If Not IsNull(ov(17)) And CStr(ov(17)) <> "" Then
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLZUserName")).Value = ov(17)
        End If

        On Error GoTo 0

NextRow:
    Next r

    Set overrides = Nothing
End Sub

' =============================================================================
' BuildOverrideLookup
' Returns Scripting.Dictionary(Job → Variant array(0..17)) for all LylesWIP
' override rows for the given company/month. O(1) lookup per job during merge.
' Returns Nothing on connection failure; returns empty Dictionary if no overrides.
' =============================================================================
Private Function BuildOverrideLookup(co As Integer, wipMonth As Date) As Object
    Set BuildOverrideLookup = Nothing

    Dim ovRs As ADODB.Recordset
    Set ovRs = GetJobOverrides(co, wipMonth)
    If ovRs Is Nothing Then Exit Function

    Dim lookup As Object
    Set lookup = CreateObject("Scripting.Dictionary")
    lookup.CompareMode = 1   ' vbTextCompare — job numbers are case-insensitive

    Do While Not ovRs.EOF
        Dim jobKey As String
        jobKey = CStr(ovRs.Fields("Job").Value)

        ' ReDim inside loop guarantees a fresh array allocation each iteration
        ' (assigning Variant array to Dictionary copies by value)
        Dim ov() As Variant
        ReDim ov(18)
        ov(0)  = ovRs.Fields("OpsRevOverride").Value
        ov(1)  = ovRs.Fields("OpsRevPlugged").Value
        ov(2)  = ovRs.Fields("GAAPRevOverride").Value
        ov(3)  = ovRs.Fields("GAAPRevPlugged").Value
        ov(4)  = ovRs.Fields("OpsCostOverride").Value
        ov(5)  = ovRs.Fields("OpsCostPlugged").Value
        ov(6)  = ovRs.Fields("GAAPCostOverride").Value
        ov(7)  = ovRs.Fields("GAAPCostPlugged").Value
        ov(8)  = ovRs.Fields("BonusProfit").Value
        ov(9)  = ovRs.Fields("OpsRevNotes").Value
        ov(10) = ovRs.Fields("GAAPRevNotes").Value
        ov(11) = ovRs.Fields("OpsCostNotes").Value
        ov(12) = ovRs.Fields("GAAPCostNotes").Value
        ov(13) = ovRs.Fields("CompletionDate").Value
        ov(14) = ovRs.Fields("IsClosed").Value
        ov(15) = ovRs.Fields("IsOpsDone").Value
        ov(16) = ovRs.Fields("IsGAAPDone").Value
        ov(17) = ovRs.Fields("UserName").Value
        ov(18) = ovRs.Fields("UpdatedAt").Value

        lookup(jobKey) = ov
        ovRs.MoveNext
    Loop

    ovRs.Close
    Set ovRs = Nothing
    Set BuildOverrideLookup = lookup
End Function

' =============================================================================
' MergePriorMonthProfitsOntoSheet
' Queries LylesWIP for the prior month (wipMonth - 1) override data and writes
' PriorProjectedProfit = OpsRevOverride - OpsCostOverride into COLPriorProjProfit.
'
' Only runs on the Ops sheet (Sheet11). The GAAP sheet gets prior profit from
' bJCOR quarterly plugs already loaded by the Vista SQL query.
'
' Reuses BuildOverrideLookup with priorMonth — no extra DB proc needed.
' ov(0) = OpsRevOverride, ov(4) = OpsCostOverride (see BuildOverrideLookup)
' =============================================================================
Public Sub MergePriorMonthProfitsOntoSheet(sh As Worksheet, co As Integer, wipMonth As Date)
    If sh.CodeName = "Sheet12" Then Exit Sub

    Dim priorMonth As Date
    priorMonth = DateSerial(Year(wipMonth), Month(wipMonth) - 1, 1)

    Dim priorOverrides As Object
    Set priorOverrides = BuildOverrideLookup(co, priorMonth)
    If priorOverrides Is Nothing Then Exit Sub
    If priorOverrides.Count = 0 Then Exit Sub

    If NumDict Is Nothing Then
        InitializeColumnDictionaries NumDict, LetDict, 1
    End If

    Dim summaryRange As Range
    Set summaryRange = sh.Range("SummaryData")

    Dim jnColAbsP As Long
    jnColAbsP = summaryRange.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Column
    Dim lastDataRowP As Long
    lastDataRowP = sh.Cells(sh.Rows.Count, jnColAbsP).End(xlUp).Row
    Dim totalRowsP As Long
    totalRowsP = Application.Max(summaryRange.Rows.Count, lastDataRowP - summaryRange.Row + 1)

    Dim r As Long
    For r = 1 To totalRowsP
        Dim jobNum As String
        jobNum = CStr(summaryRange.Cells(r, NumDict(sh.CodeName)("COLJobNumber")).Value)
        If jobNum = "" Then GoTo NextPriorRow
        If Right(jobNum, 1) <> "." Then jobNum = jobNum & "."  ' phase jobs: normalize trailing dot
        If Not priorOverrides.Exists(jobNum) Then GoTo NextPriorRow

        Dim ov() As Variant
        ov = priorOverrides(jobNum)

        Dim opsRev  As Double
        Dim opsCost As Double
        Dim bonusProfit As Double
        If IsNull(ov(0)) Or IsEmpty(ov(0)) Then opsRev = 0 Else opsRev = CDbl(ov(0))
        If IsNull(ov(4)) Or IsEmpty(ov(4)) Then opsCost = 0 Else opsCost = CDbl(ov(4))
        If IsNull(ov(8)) Or IsEmpty(ov(8)) Then bonusProfit = 0 Else bonusProfit = CDbl(ov(8))

        ' Write to Z-columns (backing store for SaveJobRow / tooltips)
        ' AND to visible columns (data rows have values, not formulas).
        On Error Resume Next
        summaryRange.Cells(r, NumDict(sh.CodeName)("COLZPriorJTDOPsProfit")).Value = opsRev - opsCost
        summaryRange.Cells(r, NumDict(sh.CodeName)("COLZPriorBonusProfit")).Value = bonusProfit
        summaryRange.Cells(r, NumDict(sh.CodeName)("COLPriorProjProfit")).Value = opsRev - opsCost
        summaryRange.Cells(r, NumDict(sh.CodeName)("COLJTDPriorProfit")).Value = bonusProfit
        On Error GoTo 0

NextPriorRow:
    Next r

    Set priorOverrides = Nothing
End Sub

' =============================================================================
' MergePriorYearBonusOntoSheet
' Writes prior year-end BonusProfit (Dec of prior year) into COLAPYBonusProfit.
' Source: LylesWIP WipMonth = Dec 1 of (wipYear - 1).
' Only runs on Ops sheet (Sheet12 has no bonus column).
' =============================================================================
Public Sub MergePriorYearBonusOntoSheet(sh As Worksheet, co As Integer, wipMonth As Date)
    If sh.CodeName = "Sheet12" Then Exit Sub

    Dim priorYearMonth As Date
    priorYearMonth = DateSerial(Year(wipMonth) - 1, 12, 1)

    Dim pyOverrides As Object
    Set pyOverrides = BuildOverrideLookup(co, priorYearMonth)
    If pyOverrides Is Nothing Then Exit Sub
    If pyOverrides.Count = 0 Then Exit Sub

    If NumDict Is Nothing Then
        InitializeColumnDictionaries NumDict, LetDict, 1
    End If

    Dim summaryRange As Range
    Set summaryRange = sh.Range("SummaryData")

    Dim jnColAbsPY As Long
    jnColAbsPY = summaryRange.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Column
    Dim lastDataRowPY As Long
    lastDataRowPY = sh.Cells(sh.Rows.Count, jnColAbsPY).End(xlUp).Row
    Dim totalRowsPY As Long
    totalRowsPY = Application.Max(summaryRange.Rows.Count, lastDataRowPY - summaryRange.Row + 1)

    Dim r As Long
    For r = 1 To totalRowsPY
        Dim jobNum As String
        jobNum = CStr(summaryRange.Cells(r, NumDict(sh.CodeName)("COLJobNumber")).Value)
        If jobNum = "" Then GoTo NextPYRow
        If Right(jobNum, 1) <> "." Then jobNum = jobNum & "."  ' phase jobs: normalize trailing dot
        If Not pyOverrides.Exists(jobNum) Then GoTo NextPYRow

        Dim ov() As Variant
        ov = pyOverrides(jobNum)

        Dim bonus As Double
        If IsNull(ov(8)) Or IsEmpty(ov(8)) Then bonus = 0 Else bonus = CDbl(ov(8))

        On Error Resume Next
        If bonus <> 0 Then
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLAPYBonusProfit")).Value = bonus
        End If

        ' Correct COLAPYRev: PYCost (already in AH) + actual bonus from LylesWIP.
        ' GetWipDetail2 writes COLAPYRev = PYCost + 0 (Vista stubs PriorYrBonusProfit
        ' to 0); this backfills the actual bonus.
        ' Also correct COLAPYCalcProfit (AI): data rows have values, not formulas,
        ' so AI won't auto-calculate from AG-AH — must write explicitly.
        Dim pyCost As Double
        pyCost = CDbl(summaryRange.Cells(r, NumDict(sh.CodeName)("COLAPYCost")).Value)
        summaryRange.Cells(r, NumDict(sh.CodeName)("COLAPYRev")).Value = pyCost + bonus
        summaryRange.Cells(r, NumDict(sh.CodeName)("COLAPYCalcProfit")).Value = bonus
        On Error GoTo 0

NextPYRow:
    Next r

    Set pyOverrides = Nothing
End Sub

' =============================================================================
' MergeTrendCommentsOntoSheet
' Queries LylesWIP for the 6 months prior to wipMonth and adds tooltip comments
' on COLOvrRevProj and COLOvrCostProj showing the override trend history.
' Format matches Michael's original: "1 - 94,176,098\n2 - 93,500,000\n..." etc.
' Ops sheet uses OpsRev/OpsCost (ov(0)/ov(4)).
' GAAP sheet uses GAAPRev/GAAPCost (ov(2)/ov(6)).
' =============================================================================
Public Sub MergeTrendCommentsOntoSheet(sh As Worksheet, co As Integer, wipMonth As Date)
    ' Load 6 prior months of overrides
    Dim priorOvr(1 To 6) As Object
    Dim m As Long
    For m = 1 To 6
        Dim mDate As Date
        mDate = DateSerial(Year(wipMonth), Month(wipMonth) - m, 1)
        Set priorOvr(m) = BuildOverrideLookup(co, mDate)
        If priorOvr(m) Is Nothing Then Set priorOvr(m) = CreateObject("Scripting.Dictionary")
    Next m

    If NumDict Is Nothing Then
        InitializeColumnDictionaries NumDict, LetDict, 1
    End If

    Dim isOps As Boolean
    isOps = (sh.CodeName = "Sheet11")

    Dim summaryRange As Range
    Set summaryRange = sh.Range("SummaryData")

    Dim jnColAbsT As Long
    jnColAbsT = summaryRange.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Column
    Dim lastDataRowT As Long
    lastDataRowT = sh.Cells(sh.Rows.Count, jnColAbsT).End(xlUp).Row
    Dim totalRowsT As Long
    totalRowsT = Application.Max(summaryRange.Rows.Count, lastDataRowT - summaryRange.Row + 1)

    Dim r As Long
    For r = 1 To totalRowsT
        Dim jobNum As String
        jobNum = CStr(summaryRange.Cells(r, NumDict(sh.CodeName)("COLJobNumber")).Value)
        If jobNum = "" Then GoTo NextTrendRow
        If Right(jobNum, 1) <> "." Then jobNum = jobNum & "."

        ' Build revenue trend string
        Dim revTrend As String
        revTrend = ""
        Dim costTrend As String
        costTrend = ""
        Dim hasRevData As Boolean
        hasRevData = False
        Dim hasCostData As Boolean
        hasCostData = False
        Dim revIdx As Long
        Dim costIdx As Long
        If isOps Then revIdx = 0: costIdx = 4 Else revIdx = 2: costIdx = 6

        For m = 1 To 6
            Dim revVal As Double
            Dim costVal As Double
            revVal = 0
            costVal = 0
            If priorOvr(m).Exists(jobNum) Then
                Dim ov() As Variant
                ov = priorOvr(m)(jobNum)
                If Not IsNull(ov(revIdx)) And Not IsEmpty(ov(revIdx)) Then
                    revVal = CDbl(ov(revIdx))
                    hasRevData = True
                End If
                If Not IsNull(ov(costIdx)) And Not IsEmpty(ov(costIdx)) Then
                    costVal = CDbl(ov(costIdx))
                    hasCostData = True
                End If
            End If
            If m = 1 Then
                revTrend = m & " - " & Format(revVal, "#,##0;(#,##0)")
                costTrend = m & " - " & Format(costVal, "#,##0;(#,##0)")
            Else
                revTrend = revTrend & vbNewLine & m & " - " & Format(revVal, "#,##0;(#,##0)")
                costTrend = costTrend & vbNewLine & m & " - " & Format(costVal, "#,##0;(#,##0)")
            End If
        Next m

        ' Add tooltip comments on override columns
        On Error Resume Next
        If hasRevData Then
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj")).ClearComments
            With summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrRevProj"))
                .AddComment.Text Text:=revTrend
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 100
            End With
        End If
        If hasCostData Then
            summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj")).ClearComments
            With summaryRange.Cells(r, NumDict(sh.CodeName)("COLOvrCostProj"))
                .AddComment.Text Text:=costTrend
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 100
            End With
        End If
        On Error GoTo 0

NextTrendRow:
    Next r

    ' Clean up
    For m = 1 To 6
        Set priorOvr(m) = Nothing
    Next m
End Sub

' =============================================================================
' Private Helpers
' =============================================================================

' ToDecimalOrNull — returns Null if value is empty/zero, else returns the number.
' Mirrors the Python to_decimal() rule: zero values become NULL = no override.
Private Function ToDecimalOrNull(v As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Or v = "" Then
        ToDecimalOrNull = Null
        Exit Function
    End If
    On Error Resume Next
    Dim d As Double
    d = CDbl(v)
    If Err.Number <> 0 Or d = 0 Then
        ToDecimalOrNull = Null
    Else
        ToDecimalOrNull = d
    End If
    On Error GoTo 0
End Function

' AddAuditComment — creates an audit trail comment on a cell showing from/to/who/when.
' Clears any existing comment first. Formats the comment for readability.
Private Sub AddAuditComment(cell As Range, fromVal As Variant, toVal As Variant, _
                             userName As Variant, updatedAt As Variant)
    On Error Resume Next
    cell.ClearComments

    Dim fromStr As String
    Dim toStr As String
    If IsNull(fromVal) Or IsEmpty(fromVal) Then
        fromStr = "(Vista calculated)"
    Else
        fromStr = Format(CDbl(fromVal), "$#,##0.00")
    End If
    If IsNull(toVal) Or IsEmpty(toVal) Then
        toStr = "(cleared)"
    Else
        toStr = Format(CDbl(toVal), "$#,##0.00")
    End If

    Dim userStr As String
    If IsNull(userName) Or CStr(userName) = "" Then userStr = "system" Else userStr = CStr(userName)

    Dim dateStr As String
    If IsNull(updatedAt) Or IsEmpty(updatedAt) Then
        dateStr = ""
    Else
        dateStr = " on " & Format(CDate(updatedAt), "mm/dd/yyyy")
    End If

    Dim commentText As String
    commentText = "Changed " & fromStr & " to " & toStr & " by " & userStr & dateStr

    cell.AddComment commentText
    cell.Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
    cell.Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
    cell.Comment.Shape.TextFrame.Characters.Font.Size = 9
    cell.Comment.Shape.Height = 30
    cell.Comment.Shape.Width = 250
    On Error GoTo 0
End Sub
