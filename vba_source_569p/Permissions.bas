Public Sub GetSecurity(User As String)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

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
Dim RoleName As String
Dim RetMsg As String
Dim Companies As String

Sheet2.Range("Role").Value = ""
Sheet2.Range("Companies").Value = ""
Sheet2.Range("Departments").Value = ""

' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("PPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("PPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.CommandTimeout = 5
cnn.Open

' GET ROLE

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pnp.WIPSECGetRole"
cmd.CommandTimeout = 180



Set cmdUser = cmd.CreateParameter("@user", adVarChar, adParamInput, 30, User)
cmd.Parameters.Append cmdUser
Set cmdRoleName = cmd.CreateParameter("@roleName", adVarChar, adParamOutput, 30)
cmd.Parameters.Append cmdRoleName
Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 30)
cmd.Parameters.Append cmdRetMsg

cmd.Execute

RoleName = cmd.Parameters("@roleName").Value
RetMsg = cmd.Parameters("@RetMsg").Value

If RetMsg <> "" Then
    MsgBox RetMsg, vbInformation
    GoTo 9998
End If

Sheet2.Range("Role").Value = RoleName

cnn.Close
9998:
GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the GetSecurity Routine. " & Err, vbOKOnly

9999: End Sub


Public Sub GetDept(User As String, Company As Integer)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

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
Dim Departments As String

Dim RetMsg As String

Sheet2.Range("Departments").Value = ""

' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("PPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("PPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' GET DEPARTMENTS

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "pnp.WIPSECGetDepartments"
cmd.CommandTimeout = 180

Set cmdCompany = cmd.CreateParameter("@company", adInteger, adParamInput, 3, Company)
cmd.Parameters.Append cmdCompany
Set cmdUser = cmd.CreateParameter("@user", adVarChar, adParamInput, 30, User)
cmd.Parameters.Append cmdUser
Set cmdDepartments = cmd.CreateParameter("@departments", adVarChar, adParamOutput, 300)
cmd.Parameters.Append cmdDepartments
Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 30)
cmd.Parameters.Append cmdRetMsg

cmd.Execute

'Departments = cmd.Parameters("@departments").Value
RetMsg = cmd.Parameters("@RetMsg").Value

cnn.Close
GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the GetDept Routine. " & Err, vbOKOnly
9999:
'Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub


