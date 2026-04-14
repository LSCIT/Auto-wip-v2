' Holds pop up list values so we don't have to query the DB each time we open the pop up
Public DeptDataList() As Variant
Public CompanyDataList() As Variant
Public Form3Datalist() As Variant
Public PRange As Range
Public RBCaller As String
Public StrPos As Integer
Public uname As String
Public LastColumnCount As Long
Public Caller As String
Public BatchSelected As Boolean
Public BatchSelectionCanceled As Boolean
Public DeptSelectionCanceled As Boolean
Public DeptList As String


'Public Sub InitCols()
'    Set NumDict = Nothing
'    Set NumDict = Nothing
'    InitializeColumnDictionaries NumDict, LetDict, 1
'End Sub


Public Sub ee()
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Public Sub eo()
Application.EnableEvents = False
End Sub

Public Sub p()

If Sheet2.Range("ProtectSheet").Value = "True" Then
        Sheet11.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet12.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet13.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet14.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet15.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet16.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        
    Else
        
        Sheet11.Unprotect "password"
        Sheet12.Unprotect "password"
        Sheet13.Unprotect "password"
        Sheet14.Unprotect "password"
        Sheet15.Unprotect "password"
        Sheet16.Unprotect "password"
        
    End If



End Sub


' Get Department list
Public Sub GetDeptData(ByRef FormList As Variant)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
Application.EnableEvents = False
Application.ScreenUpdating = False

'Declare variables'
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim DataSQL As String
Dim Datars As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Set Datars = New ADODB.Recordset
Dim i As Integer
Dim intTotalElements As Integer
Dim intCount As Integer


' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If


cnn.Open

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPGetDeptData1"
cmd.CommandTimeout = 180


Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 3, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 60, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, Sheet2.Range("UserName2").Value)
cmd.Parameters.Append cmdUserName
Set cmdRole = cmd.CreateParameter("@Role", adVarChar, adParamInput, 30, Sheet2.Range("Role").Value)
cmd.Parameters.Append cmdRole
Set cmdUsePostedBatch = cmd.CreateParameter("@UsePostedBatch", adVarChar, adParamInput, 1, Sheet2.Range("UsePostedBatch1").Value)
cmd.Parameters.Append cmdUsePostedBatch
Set cmdDeptList = cmd.CreateParameter("@DeptList", adVarChar, adParamInput, 30, DeptList)
cmd.Parameters.Append cmdDeptList



cnn.CursorLocation = adUseClient

Datars.Open cmd.Execute

If Datars.EOF <> True Then
    Datars.MoveFirst
    i = 0
    Do
        ReDim Preserve FormList(i)
        FormList(i) = PadString(CStr(Datars![Department]), 10, "R") & " " & PadString(VBA.UCase(Datars![Description]), 30, "R")
        i = i + 1
        Datars.MoveNext
    Loop Until Datars.EOF
    
    cnn.Close
Else
    MsgBox ("There are no Divisions Approved For this Month.")
End If


GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the GetDeptData Routine" & Err, vbOKOnly
Application.EnableEvents = True
Application.ScreenUpdating = True

9999:

Application.EnableEvents = True
'Application.ScreenUpdating = True
End Sub


' Get Company List
Public Sub GetCompanyData(ByRef FormList() As Variant)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
Application.ScreenUpdating = False

'Declare variables'
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim DataSQL As String
Dim Datars As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Set Datars = New ADODB.Recordset
Dim i As Integer
Dim intTotalElements As Integer
Dim intCount As Integer


' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If


cnn.Open

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPGetCoList1"
cmd.CommandTimeout = 180


Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 300, Sheet2.Range("UserName2").Value)
cmd.Parameters.Append cmdUserName

cnn.CursorLocation = adUseClient

Datars.Open cmd.Execute

If Datars.EOF <> True Then
    Datars.MoveFirst
    i = 0
    Do
        ReDim Preserve FormList(i)
        FormList(i) = PadString(CStr(Datars![JCCo]), 10, "R") & " " & PadString(VBA.UCase(Datars![Name]), 30, "R")
        i = i + 1
        Datars.MoveNext
    Loop Until Datars.EOF
    
    cnn.Close
Else
    MsgBox ("There are no Companies Available.")
End If


GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the GetCompanyData Routine" & Err, vbOKOnly
Application.EnableEvents = True
Application.ScreenUpdating = True

9999: Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub





' Clear Tables
Public Sub ClearTables()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
Application.ScreenUpdating = False

'Declare variables'
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim DataSQL As String
Dim Datars As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Set Datars = New ADODB.Recordset
Dim i As Integer
Dim intTotalElements As Integer
Dim intCount As Integer


' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If


cnn.Open

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPClearTables"
cmd.CommandTimeout = 180


Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 3, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdBegMonth = cmd.CreateParameter("@BegMonth", adDate, adParamInput, 300, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdBegMonth

cmd.Execute


cnn.Close

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the ClearTables Routine" & Err, vbOKOnly
Application.EnableEvents = True
Application.ScreenUpdating = True

9999: Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

