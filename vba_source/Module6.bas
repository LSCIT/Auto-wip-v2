
Public Sub UseExistingBatch(ByRef NoData As Boolean)
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
Dim i As Integer
Dim intTotalElements As Integer
Dim DeptNameList As String
Dim rcode As Integer
Dim missingDepts As String
DeptList = ""

ProtectUnProtect ("Off")
Application.EnableEvents = False

' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPBatchCheck1"
cmd.CommandTimeout = 180



Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, Sheet2.Range("UserName2").Value)
cmd.Parameters.Append cmdUserName
Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput)
cmd.Parameters.Append cmdRcode
Set cmdDeptList = cmd.CreateParameter("@DeptList", adVarChar, adParamOutput, 200)
cmd.Parameters.Append cmdDeptList
Set cmdDeptNameList = cmd.CreateParameter("@DeptNameList", adVarChar, adParamOutput, 200)
cmd.Parameters.Append cmdDeptNameList


cmd.Execute

rcode = cmd.Parameters("@rcode").Value

If rcode = 0 Then
    ' This is a list of Dept in a batch that you have access to
    DeptList = cmd.Parameters("@DeptList").Value
    DeptNameList = cmd.Parameters("@DeptNameList").Value
Else
    If Sheet2.Range("Role").Value <> "WIPAccounting" Then
        MsgBox "There are no divisions ready for your input", vbInformation
        NoData = True
        GoTo 9999
    End If
    'NoData = True
End If


Dim Deptfrm As Dept
Set Deptfrm = New Dept
Deptfrm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
Deptfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Deptfrm.Width)
Deptfrm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Deptfrm.Height)

Application.EnableEvents = False

' If Batch Found
If rcode <> 1 Then

    If MsgBox("     Load Divisions (" & DeptList & ")" & vbCrLf & vbCrLf & "Select No to Pick Divisions from List", vbYesNo) = vbYes Then
    
        Sheet17.Range("StartDept").Value = cmd.Parameters("@DeptList").Value
        Sheet17.Range("StartDeptName").Value = cmd.Parameters("@DeptNameList").Value
    
    Else
        
        If UsePostedBatch = True Then
            CreateBatch
        Else
            Deptfrm.Show
            
            If Sheet2.Range("Role").Value = "WIPAccounting" Then
            
                missingDepts = GetMissingItems(Sheet17.Range("StartDept").Value, cmd.Parameters("@DeptList").Value)
                
                If missingDepts <> "" Then
                    CreateBatch
                End If
            
            End If
        End If
    
    End If

Else
    ' If a Posted Batch exists Set Settings "UsePostedBatch1" to "Y" and Create Batch
    If UsePostedBatch = True Then
        CreateBatch
    Else
        Deptfrm.Show
        CreateBatch
    End If

End If

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the BatchCheck Routine. " & Err, vbOKOnly

9999:
ProtectUnProtect ("On")
Application.EnableEvents = True

End Sub



Public Function UsePostedBatch() As Boolean
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
Dim i As Integer
Dim intTotalElements As Integer
'Dim DeptList As String
Dim DeptNameList As String
Dim rcode As Integer
Dim missingDepts As String

' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPPostedCheck"
cmd.CommandTimeout = 180



Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput)
cmd.Parameters.Append cmdRcode
Set cmdDeptList = cmd.CreateParameter("@DeptList", adVarChar, adParamOutput, 200)
cmd.Parameters.Append cmdDeptList
Set cmdDeptNameList = cmd.CreateParameter("@DeptNameList", adVarChar, adParamOutput, 200)
cmd.Parameters.Append cmdDeptNameList


cmd.Execute

rcode = cmd.Parameters("@rcode").Value

If rcode = 0 Then
    ' This is a list of Dept in a Poset Batch
    DeptList = cmd.Parameters("@DeptList").Value
End If

Dim Deptfrm As Dept
Set Deptfrm = New Dept
Deptfrm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
Deptfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Deptfrm.Width)
Deptfrm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Deptfrm.Height)

Application.EnableEvents = False





' If Batch Found
If rcode <> 1 Then

    If MsgBox("There are Posted Divisions for this Month" & vbCrLf & vbCrLf & "                   Load Posted Data", vbYesNo) = vbYes Then
    
        Sheet2.Range("UsePostedBatch1").Value = "Y"
        Deptfrm.Show
        UsePostedBatch = True
    
    Else
        
        Sheet2.Range("UsePostedBatch1").Value = "N"
        UsePostedBatch = False
        
    End If

Else
    UsePostedBatch = False
End If


GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the BatchCheck Routine. " & Err, vbOKOnly

9999:

End Function







Function GetMissingItems(ByVal listIn As String, ByVal listCheckAgainst As String) As String
    Dim arrIn() As String
    Dim arrCheck() As String
    Dim item As Variant
    Dim Dict As Object
    Dim missingItems As Collection
    Dim missing As String
    
    ' Split the strings into arrays
    arrIn = Split(listIn, ",")
    arrCheck = Split(listCheckAgainst, ",")

    ' Create a dictionary for fast lookup
    Set Dict = CreateObject("Scripting.Dictionary")
    Set missingItems = New Collection
    
    ' Add items from listCheckAgainst to dictionary
    For Each item In arrCheck
        If Not Dict.Exists(Trim(item)) Then
            Dict.Add Trim(item), 1
        End If
    Next item

    ' Check for missing items
    For Each item In arrIn
        If Not Dict.Exists(Trim(item)) Then
            missingItems.Add Trim(item)
        End If
    Next item

    ' Convert collection to comma-delimited string
    For Each item In missingItems
        missing = missing & item & ","
    Next item
    
    ' Remove the trailing comma if there were missing items
    If Len(missing) > 0 Then
        missing = Left(missing, Len(missing) - 1)
    End If
    
    GetMissingItems = missing
End Function


Public Sub CreateBatch()
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
Dim BatchCount As Integer
Dim BatchId As Integer
Dim frm As DataRetrievalStatus

Application.EnableEvents = False
ProtectUnProtect ("Off")

' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

Set frm = New DataRetrievalStatus
frm.Label1.Caption = "Creating / Modifying New WIP Month.." & vbCrLf & "This may take a few minutes to complete......"
frm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
frm.Show vbModeless
DoEvents

Application.EnableEvents = False

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPCreateBatch"
cmd.CommandTimeout = 180

Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 30, Sheet17.Range("StartDept").Value)
cmd.Parameters.Append cmdDept
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, Sheet2.Range("UserName2").Value)
cmd.Parameters.Append cmdUserName
Set cmdUsePostedBatch = cmd.CreateParameter("@UsePostedBatch", adVarChar, adParamInput, 30, Sheet2.Range("UsePostedBatch1").Value)
cmd.Parameters.Append cmdUsePostedBatch
Set cmdBatchId = cmd.CreateParameter("@BatchId", adInteger, adParamOutput)
cmd.Parameters.Append cmdBatchId
Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)
cmd.Parameters.Append cmdRetMsg

cmd.Execute

BatchId = cmd.Parameters("@BatchId").Value
RetMsg = cmd.Parameters("@RetMsg").Value

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPCreateBatchJV"
cmd.CommandTimeout = 180

Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth


cmd.Execute

Unload frm
cnn.Close
OpenBatch:

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the BatchCheck Routine. " & Err, vbOKOnly

9999:
ProtectUnProtect ("On")
Application.EnableEvents = True

End Sub



Public Sub ClearWIPDetail()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If MsgBox("Are you Sure you want to Clear this Batch?" & vbNewLine & vbNewLine & "           *****This Is Not Reversible*****", vbOKCancel) = vbCancel Then
    GoTo 9999
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
Dim RetMsg As String
Dim frm As DataRetrievalStatus

ProtectUnProtect ("Off")


' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPClearDetail"
cmd.CommandTimeout = 180

Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 200, Sheet17.Range("StartDept").Value)
cmd.Parameters.Append cmdDept
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth

cmd.Execute


ClearForms3

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the CancelBatch Routine. " & Err, vbOKOnly

9999:
ProtectUnProtect ("On")

End Sub

Public Function GetGAAPRev(row As Long, sh As Worksheet) As Double

If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOR")).Value = "T" Then
    GetGAAPRev = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRevNew")).Value
Else
    GetGAAPRev = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRev")).Value
End If

End Function

Public Function GetGAAPRevPlug(row As Long, sh As Worksheet) As String

If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOR")).Value = "T" Then
   
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRevNew")).Font.Bold = True Then
        GetGAAPRevPlug = "Y"
    Else
        GetGAAPRevPlug = "N"
    End If
    
Else
    
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRev")).Font.Bold = True Then
        GetGAAPRevPlug = "Y"
    Else
        GetGAAPRevPlug = "N"
    End If
End If

End Function


Public Function GetGAAPCost(row As Long, sh As Worksheet) As Double

If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOP")).Value = "T" Then
    GetGAAPCost = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCostNew")).Value
Else
    GetGAAPCost = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCost")).Value
End If

End Function

Public Function GetGAAPCostPlug(row As Long, sh As Worksheet) As String

If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOP")).Value = "T" Then
   
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCostNew")).Font.Bold = True Then
        GetGAAPCostPlug = "Y"
    Else
        GetGAAPCostPlug = "N"
    End If
    
Else
    
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCost")).Font.Bold = True Then
        GetGAAPCostPlug = "Y"
    Else
        GetGAAPCostPlug = "N"
    End If
End If

End Function





Public Function GetOpsCost(row As Long, sh As Worksheet) As Double

If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCChg")).Value = "T" Then
    GetOpsCost = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCostNew")).Value
Else
    GetOpsCost = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCost")).Value
End If

End Function

Public Function GetOpsCostPlug(row As Long, sh As Worksheet) As String

If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCChg")).Value = "T" Then
   
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCostNew")).Font.Bold = True Then
        GetOpsCostPlug = "Y"
    Else
        GetOpsCostPlug = "N"
    End If
    
Else
    
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCost")).Font.Bold = True Then
        GetOpsCostPlug = "Y"
    Else
        GetOpsCostPlug = "N"
    End If
End If

End Function


Public Function GetOpsRev(row As Long, sh As Worksheet) As Double

If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRChg")).Value = "T" Then
    GetOpsRev = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRevNew")).Value
Else
    GetOpsRev = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRev")).Value
End If

End Function

Public Function GetOpsRevPlug(row As Long, sh As Worksheet) As String

If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRChg")).Value = "T" Then
   
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRevNew")).Font.Bold = True Then
        GetOpsRevPlug = "Y"
    Else
        GetOpsRevPlug = "N"
    End If
    
Else
    
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRev")).Font.Bold = True Then
        GetOpsRevPlug = "Y"
    Else
        GetOpsRevPlug = "N"
    End If
End If

End Function


Public Function GetOpsBonusPlug(row As Long, sh As Worksheet) As String

If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsBonusNew")).Font.Bold = True Then
    GetOpsBonusPlug = "Y"
Else
    GetOpsBonusPlug = "N"
End If
    

End Function



Function HexToByteArray(hexString As String) As Byte()
    Dim byteArray() As Byte
    Dim i As Long
    Dim length As Long
    
    ' Remove any spaces from the hex string
    hexString = Replace(hexString, " ", "")
    
    ' Ensure the string length is even (since each byte is represented by two characters)
    If Len(hexString) Mod 2 <> 0 Then
        Err.Raise 5, "HexToByteArray", "Hex string must have an even number of characters."
        Exit Function
    End If
    
    length = Len(hexString)
    ReDim byteArray(length \ 2 - 1) ' Each byte needs two chars
    
    ' Iterate over the string two characters at a time
    For i = 0 To length - 1 Step 2
        ' Convert two hex characters to a byte
        byteArray(i \ 2) = CLng("&H" & Mid(hexString, i + 1, 2))
    Next i
    
    HexToByteArray = byteArray
End Function


Public Sub UpdateAllRows(sh As Worksheet, Complete As String)

Dim cell As Range
Dim rng As Range
Dim ColOffset As Integer
Dim cp As String
If Complete = "Y" Then
    cp = "P"
Else
    cp = ""
End If


If sh.CodeName = "Sheet11" Then ' Ops

    ColOffset = NumDict(sh.CodeName)("COLJobNumber") - NumDict(sh.CodeName)("COLDone")


    Set rng = Sheet11.Range("Done")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
        
        If cp = "" Then
            cell.ClearComments
        End If
    Next cell
    
'    Set rng = Sheet12.Range("Done")
'    ColOffset = NumDict("Sheet12")("COLJobNumber") - NumDict("Sheet12")("COLDone")
'
'    For Each cell In rng
'        If cell.Offset(0, ColOffset).Value <> "" Then
'            cell.Value = cp
'        End If
'    Next cell



Else ' GAAP

    ColOffset = NumDict(sh.CodeName)("COLJobNumber") - NumDict(sh.CodeName)("COLGAAPDone")

    Set rng = Sheet12.Range("DoneGAAP")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
        
        If cp = "" Then
            cell.ClearComments
        End If
        
    Next cell

End If




For Each cell In sh.Range("MyJobNos")

If cell.Value <> "" Then

    Call UpdateRow(cell.row, sh)

End If

Next cell





End Sub




Public Sub UpdateRow(row As Long, sh As Worksheet)

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
Dim intTotalElements As Integer
Dim RetMsg As String
Dim frm As DataRetrievalStatus

Set frm = New DataRetrievalStatus
frm.Label1.Caption = "Updating Database......"
frm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
frm.Show vbModeless
DoEvents
'

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
cmd.CommandTimeout = 180

Dim Done As String

If sh.CodeName = "Sheet11" Then
    ' Ops
    
    cmd.CommandText = "LCGWIPUpdateRowOps"
    Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
    cmd.Parameters.Append cmdCo
    Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
    cmd.Parameters.Append cmdMonth
    Set cmdKeyId = cmd.CreateParameter("@KeyId", adInteger, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZBatchSeq")).Value)
    cmd.Parameters.Append cmdKeyId
    Set cmdDept = cmd.CreateParameter("@cmdDept", adVarChar, adParamInput, 10, Left(sh.Cells(row, NumDict(sh.CodeName)("COLJobNumber")).Value, 2))
    cmd.Parameters.Append cmdDept
    Set cmdCompletionDate = cmd.CreateParameter("@CompletionDate", adDate, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLCompDate")).Value)
    cmd.Parameters.Append cmdCompletionDate
    Set cmdOpsRev = cmd.CreateParameter("@OpsRev", adDouble, adParamInput, 30, GetOpsRev(row, sh))
    cmd.Parameters.Append cmdOpsRev
    Set cmdOpsCost = cmd.CreateParameter("@OpsCost", adDouble, adParamInput, 30, GetOpsCost(row, sh))
    cmd.Parameters.Append cmdOpsCost
    Set cmdEstimator = cmd.CreateParameter("@Estimator", adVarChar, adParamInput, 30, "")
    cmd.Parameters.Append cmdEstimator
    Set cmdPM = cmd.CreateParameter("@PM", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLPrjMngr")).Value)
    cmd.Parameters.Append cmdPM
    Set cmdProjectExec = cmd.CreateParameter("@ProjectExec", adVarChar, adParamInput, 30, "")
    cmd.Parameters.Append cmdProjectExec
    Set cmdOpsRevNotes = cmd.CreateParameter("@OpsRevNotes", adVarChar, adParamInput, 8000, sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRevNotes")).Value)
    cmd.Parameters.Append cmdOpsRevNotes
    Set cmdOpsCostNotes = cmd.CreateParameter("@OpsCostNotes", adVarChar, adParamInput, 8000, sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCostNotes")).Value)
    cmd.Parameters.Append cmdOpsCostNotes
    Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, Sheet2.Range("UserName2"))
    cmd.Parameters.Append cmdUserName
    Set cmdClose = cmd.CreateParameter("@Close", adVarChar, adParamInput, 10, sh.Cells(row, NumDict(sh.CodeName)("COLClose")).Value)
    cmd.Parameters.Append cmdClose
    Set cmdCompleted = cmd.CreateParameter("@Completed", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value)
    cmd.Parameters.Append cmdCompleted
    Set cmdOpsRevPlugged = cmd.CreateParameter("@OpsRevPlugged", adVarChar, adParamInput, 30, GetOpsRevPlug(row, sh))
    cmd.Parameters.Append cmdOpsRevPlugged
    Set cmdOpsCostPlugged = cmd.CreateParameter("@OpsCostPlugged", adVarChar, adParamInput, 30, GetOpsCostPlug(row, sh))
    cmd.Parameters.Append cmdOpsCostPlugged
    Set cmdBonusProfit = cmd.CreateParameter("@BonusProfit", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJTDBonusProfit")).Value)
    cmd.Parameters.Append cmdBonusProfit
    Set cmdBonusProfitPlugged = cmd.CreateParameter("@BonusProfitPlugged", adVarChar, adParamInput, 1, GetOpsBonusPlug(row, sh))
    cmd.Parameters.Append cmdBonusProfitPlugged
    Set cmdBonusProfitNotes = cmd.CreateParameter("@BonusProfitNotes", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJTDBonusProfitNotes")).Value)
    cmd.Parameters.Append cmdBonusProfitNotes
    Set cmdRowVersion = cmd.CreateParameter("@RowVersion", adVarBinary, adParamInput, 8, HexToByteArray(sh.Cells(row, NumDict(sh.CodeName)("COLZRowVersion")).Value))
    cmd.Parameters.Append cmdRowVersion
    Set cmdCurRowVersion = cmd.CreateParameter("@CurRowVersionOut", adVarChar, adParamOutput, 16)
    cmd.Parameters.Append cmdCurRowVersion
    Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput, 30)
    cmd.Parameters.Append cmdRcode
    
    Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)
    cmd.Parameters.Append cmdRetMsg


Else

    cmd.CommandText = "LCGWIPUpdateRowGAAP"
    Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
    cmd.Parameters.Append cmdCo
    Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
    cmd.Parameters.Append cmdMonth
    Set cmdKeyId = cmd.CreateParameter("@KeyId", adInteger, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZBatchSeq")).Value)
    cmd.Parameters.Append cmdKeyId
    Set cmdDept = cmd.CreateParameter("@cmdDept", adVarChar, adParamInput, 10, Left(sh.Cells(row, NumDict(sh.CodeName)("COLJobNumber")).Value, 2))
    cmd.Parameters.Append cmdDept
    Set cmdGAAPRev = cmd.CreateParameter("@GAAPRev", adDouble, adParamInput, 30, GetGAAPRev(row, sh))
    cmd.Parameters.Append cmdGAAPRev
    Set cmdGAAPOtherRevAmount = cmd.CreateParameter("@GAAPOtherRevAmount", adDouble, adParamInput, 30, 0)
    cmd.Parameters.Append cmdGAAPOtherRevAmount
    Set cmdGAAPRevNotes = cmd.CreateParameter("@GAAPRevNotes", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRevNotes")).Value)
    cmd.Parameters.Append cmdGAAPRevNotes
    Set cmdGAAPRevPlugged = cmd.CreateParameter("@GAAPRevPlugged", adVarChar, adParamInput, 30, GetGAAPRevPlug(row, sh))
    cmd.Parameters.Append cmdGAAPRevPlugged
    Set cmdGAAPCost = cmd.CreateParameter("@GAAPCost", adDouble, adParamInput, 30, GetGAAPCost(row, sh))
    cmd.Parameters.Append cmdGAAPCost
    Set cmdGAAPOtherCostAmount = cmd.CreateParameter("@GAAPOtherCostAmount", adDouble, adParamInput, 30, 0)
    cmd.Parameters.Append cmdGAAPOtherCostAmount
    Set cmdGAAPCostNotes = cmd.CreateParameter("@GAAPCostNotes", adVarChar, adParamInput, 200, sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCostNotes")).Value)
    cmd.Parameters.Append cmdGAAPCostNotes
    Set cmdGAAPCostPlugged = cmd.CreateParameter("@GAAPCostPlugged", adVarChar, adParamInput, 30, GetGAAPCostPlug(row, sh))
    cmd.Parameters.Append cmdGAAPCostPlugged
    Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, Sheet2.Range("UserName2"))
    cmd.Parameters.Append cmdUserName
    Set cmdCompletedGAAP = cmd.CreateParameter("@CompletedGAAP", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLGAAPDone")).Value)
    cmd.Parameters.Append cmdCompletedGAAP
    Set cmdRowVersion = cmd.CreateParameter("@RowVersion", adVarBinary, adParamInput, 30, HexToByteArray(sh.Cells(row, NumDict(sh.CodeName)("COLZRowVersion")).Value))
    cmd.Parameters.Append cmdRowVersion
    Set cmdCurRowVersion = cmd.CreateParameter("@CurRowVersionOut", adVarChar, adParamOutput, 16)
    cmd.Parameters.Append cmdCurRowVersion
    Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput, 30)
    cmd.Parameters.Append cmdRcode
    
    Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)
    cmd.Parameters.Append cmdRetMsg

End If

cmd.Execute


RetMsg = cmd.Parameters("@RetMsg").Value

If cmd.Parameters("@rcode").Value = 1 Then
    Unload frm
    
    If sh.CodeName = "Sheet11" Then
    
        If sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = "P" Then
            sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = ""
        Else
            sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = "P"
        End If
        
    Else
    
        If sh.Cells(row, NumDict(sh.CodeName)("COLGAAPDone")).Value = "P" Then
            sh.Cells(row, NumDict(sh.CodeName)("COLGAAPDone")).Value = ""
        Else
            sh.Cells(row, NumDict(sh.CodeName)("COLGAAPDone")).Value = "P"
        End If
    
    
    End If
    
    MsgBox RetMsg, vbInformation
    
    GoTo 8888
Else
    Sheet11.Unprotect "password"
    Sheet11.Cells(row, NumDict(Sheet11.CodeName)("COLZRowVersion")).Value = cmd.Parameters("@CurRowVersionOut")
    If Sheet2.Range("ProtectSheet") = True Then
        Sheet11.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    End If
    Sheet12.Unprotect "password"
    Sheet12.Cells(row, NumDict(Sheet12.CodeName)("COLZRowVersion")).Value = cmd.Parameters("@CurRowVersionOut")
    If Sheet2.Range("ProtectSheet") = True Then
        Sheet12.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    End If


End If

Unload frm

8888:
cnn.Close

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the UpdateRow Routine. " & Err, vbOKOnly

9999:

End Sub







Public Sub UpdateRowJV(row As Long, sh As Worksheet)

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
'Dim i As Integer
Dim intTotalElements As Integer
Dim RetMsg As String
Dim frm As DataRetrievalStatus

Set frm = New DataRetrievalStatus
frm.Label1.Caption = "Updating Database......"
frm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
frm.Show vbModeless
DoEvents
'

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
Dim Completed As String


If sh.CodeName = "Sheet15" Then

    If sh.Cells(row, NumDict(sh.CodeName)("COLDoneGAAP")).Value = "P" Then
        Completed = "Y"
    Else
        Completed = "N"
    End If


    cmd.CommandText = "LCGWIPUpdateRowJVGAAP"
    cmd.CommandTimeout = 180
    
    Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
    cmd.Parameters.Append cmdCo
    Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
    cmd.Parameters.Append cmdMonth
    Set cmdBatchSeq = cmd.CreateParameter("@BatchSeq", adInteger, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZBatchSeq")).Value)
    cmd.Parameters.Append cmdBatchSeq
    Set cmdGAAPContractAmt = cmd.CreateParameter("@GAAPContractAmt", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVCurContAmt")).Value)
    cmd.Parameters.Append cmdGAAPContractAmt
    Set cmdGAAPProjectedFinalProfit = cmd.CreateParameter("@GAAPProjectedFinalProfit", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVProjFinalProfit")).Value)
    cmd.Parameters.Append cmdGAAPProjectedFinalProfit
    Set cmdGAAPEarnedRev = cmd.CreateParameter("@GAAPEarnedRev", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVJTDEarnedRev")).Value)
    cmd.Parameters.Append cmdGAAPEarnedRev
    Set cmdGAAPJTDCost = cmd.CreateParameter("@GAAPJTDCost", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVJTDCost")).Value)
    cmd.Parameters.Append cmdGAAPJTDCost
    Set cmdGAAPJTDBillings = cmd.CreateParameter("@GAAPJTDBillings", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVBILLBillings")).Value)
    cmd.Parameters.Append cmdGAAPJTDBillings
    Set cmdCompleted = cmd.CreateParameter("@Completed", adVarChar, adParamInput, 30, Completed)
    cmd.Parameters.Append cmdCompleted
    Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZUserName")).Value)
    cmd.Parameters.Append cmdUserName
    Set cmdRowVersion = cmd.CreateParameter("@RowVersion", adVarBinary, adParamInput, 30, HexToByteArray(sh.Cells(row, NumDict(sh.CodeName)("COLZRowVersion")).Value))
    cmd.Parameters.Append cmdRowVersion
    Set cmdCurRowVersion = cmd.CreateParameter("@CurRowVersionOut", adVarChar, adParamOutput, 16)
    cmd.Parameters.Append cmdCurRowVersion
    Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput, 30)
    cmd.Parameters.Append cmdRcode
    Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)
    cmd.Parameters.Append cmdRetMsg
    
Else

    If sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = "P" Then
        Completed = "Y"
    Else
        Completed = "N"
    End If
    
    cmd.CommandText = "LCGWIPUpdateRowJVOps"
    cmd.CommandTimeout = 180
    
    Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
    cmd.Parameters.Append cmdCo
    Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
    cmd.Parameters.Append cmdMonth
    Set cmdBatchSeq = cmd.CreateParameter("@BatchSeq", adInteger, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZBatchSeq")).Value)
    cmd.Parameters.Append cmdBatchSeq
    Set cmdOpsContractAmt = cmd.CreateParameter("@OpsContractAmt", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVCurContAmt")).Value)
    cmd.Parameters.Append cmdOpsContractAmt
    Set cmdOpsEarnedRev = cmd.CreateParameter("@OpsEarnedRev", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVJTDEarnedRev")).Value)
    cmd.Parameters.Append cmdOpsEarnedRev
    Set cmdOpsJTDCost = cmd.CreateParameter("@OpsJTDCost", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVJTDCost")).Value)
    cmd.Parameters.Append cmdOpsJTDCost
    Set cmdOpsProjectedRevenue = cmd.CreateParameter("@OpsProjectedRevenue", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVOvrRevProj")).Value)
    cmd.Parameters.Append cmdOpsProjectedRevenue
    Set cmdOpsProjectedCost = cmd.CreateParameter("@OpsProjectedCost", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVOvrCostProj")).Value)
    cmd.Parameters.Append cmdOpsProjectedCost
    Set cmdOpsJTDBillings = cmd.CreateParameter("@OpsJTDBillings", adDouble, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLJVBILLBillings")).Value)
    cmd.Parameters.Append cmdOpsJTDBillings
    Set cmdCompleted = cmd.CreateParameter("@Completed", adVarChar, adParamInput, 30, Completed)
    cmd.Parameters.Append cmdCompleted
    Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, sh.Cells(row, NumDict(sh.CodeName)("COLZUserName")).Value)
    cmd.Parameters.Append cmdUserName
    Set cmdRowVersion = cmd.CreateParameter("@RowVersion", adVarBinary, adParamInput, 30, HexToByteArray(sh.Cells(row, NumDict(sh.CodeName)("COLZRowVersion")).Value))
    cmd.Parameters.Append cmdRowVersion
    Set cmdCurRowVersion = cmd.CreateParameter("@CurRowVersionOut", adVarChar, adParamOutput, 16)
    cmd.Parameters.Append cmdCurRowVersion
    Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput, 30)
    cmd.Parameters.Append cmdRcode
    Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)
    cmd.Parameters.Append cmdRetMsg
    
End If

cmd.Execute

RetMsg = cmd.Parameters("@RetMsg").Value

If cmd.Parameters("@rcode").Value = 1 Then
    Unload frm
    
    If sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = "P" Then
        sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = ""
    Else
        sh.Cells(row, NumDict(sh.CodeName)("COLDone")).Value = "P"
    End If
    
    MsgBox RetMsg, vbInformation
    
    GoTo 8888
Else
    sh.Cells(row, NumDict(sh.CodeName)("COLZRowVersion")).Value = cmd.Parameters("@CurRowVersionOut")
End If

Unload frm


8888:
cnn.Close

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the UpdateRow Routine. " & Err, vbOKOnly

9999:


End Sub

Public Sub UpdateApprovals()
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
Dim RetMsg As String
Dim frm As DataRetrievalStatus
Dim ReadyForOps As String
Dim InitAppr As String
Dim FinalAppr As String
Dim AcctAppr As String

ReadyForOps = Sheet2.Range("ReadyForOpsAppr1").Value
InitAppr = Sheet2.Range("InitAppr").Value
FinalAppr = Sheet2.Range("FinalAppr").Value
AcctAppr = Sheet2.Range("AcctAppr").Value


Application.EnableEvents = False
ProtectUnProtect ("Off")

Sheet2.Range("Sent").Value = True

' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPUpdateApprovals1"
cmd.CommandTimeout = 180



Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 300, Sheet17.Range("StartDept").Value)
cmd.Parameters.Append cmdDept
Set cmdInitAppr = cmd.CreateParameter("@InitApproval", adVarChar, adParamInput, 1, Sheet2.Range("InitAppr").Value)
cmd.Parameters.Append cmdInitAppr
Set cmdFinalAppr = cmd.CreateParameter("@FinalApproval", adVarChar, adParamInput, 1, Sheet2.Range("FinalAppr").Value)
cmd.Parameters.Append cmdFinalAppr
Set cmdReadyForOps = cmd.CreateParameter("@ReadyForOps", adVarChar, adParamInput, 1, Sheet2.Range("ReadyForOpsAppr1").Value)
cmd.Parameters.Append cmdReadyForOps
Set cmdAcctAppr = cmd.CreateParameter("@AcctApproval", adVarChar, adParamInput, 1, Sheet2.Range("AcctAppr").Value)
cmd.Parameters.Append cmdAcctAppr
Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 512)
cmd.Parameters.Append cmdRetMsg

cmd.Execute

RetMsg = cmd.Parameters("@RetMsg").Value

MsgBox RetMsg, vbInformation

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the CancelBatch Routine. " & Err, vbOKOnly

9999:
ProtectUnProtect ("On")
Application.EnableEvents = True

End Sub



Public Function CompleteCheck(CompType As String, Caller As String) As Boolean
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
Dim RetMsg As String
Dim frm As DataRetrievalStatus
Dim ReadyForOps As String
Dim InitAppr As String
Dim FinalAppr As String
Dim AcctAppr As String

ReadyForOps = Sheet2.Range("ReadyForOpsAppr1").Value
InitAppr = Sheet2.Range("InitAppr").Value
FinalAppr = Sheet2.Range("FinalAppr").Value
AcctAppr = Sheet2.Range("AcctAppr").Value


' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPCompleteCheck"
cmd.CommandTimeout = 180



Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 300, Sheet17.Range("StartDept").Value)
cmd.Parameters.Append cmdDept
Set cmdCompType = cmd.CreateParameter("@CompType", adVarChar, adParamInput, 1, CompType)
cmd.Parameters.Append cmdCompType
Set cmdRcode = cmd.CreateParameter("@rcode", adInteger, adParamOutput, 512)
cmd.Parameters.Append cmdRcode
Set cmdRetMsg = cmd.CreateParameter("@RetMsg", adVarChar, adParamOutput, 8000)
cmd.Parameters.Append cmdRetMsg

cmd.Execute


If cmd.Parameters("@rcode").Value <> 0 Then
    RetMsg = cmd.Parameters("@RetMsg").Value
    If Caller = "" Then
        MsgBox RetMsg, vbInformation
    End If
    
    CompleteCheck = False
    Exit Function
End If

CompleteCheck = True


GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the CompleteCheck Routine. " & Err, vbOKOnly

9999:

End Function


Public Function MarkAllComplete(CodeName As String, Complete As String, Field As String)
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
Dim RetMsg As String
Dim frm As DataRetrievalStatus



' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPMarkAllComplete"
cmd.CommandTimeout = 180


Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 300, Sheet17.Range("StartDept").Value)
cmd.Parameters.Append cmdDept
Set cmdComplete = cmd.CreateParameter("@Complete", adVarChar, adParamInput, 1, Complete)
cmd.Parameters.Append cmdComplete
Set cmdCompleteField = cmd.CreateParameter("@CompleteField", adVarChar, adParamInput, 30, Field)
cmd.Parameters.Append cmdCompleteField


cmd.Execute

Dim rng As Range
Dim cell As Range
Dim ColOffset As Integer
Dim cp As String
If Complete = "Y" Then
    cp = "P"
Else
    cp = ""
End If


If Field = "Done" Then

    ColOffset = NumDict(CodeName)("COLJobNumber") - NumDict(CodeName)("COLDone")


    Set rng = Sheet11.Range("Done")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
    Next cell
    
    Set rng = Sheet12.Range("Done")
    ColOffset = NumDict("Sheet12")("COLJobNumber") - NumDict(CodeName)("COLDone")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
    Next cell



Else

    ColOffset = NumDict(CodeName)("COLJobNumber") - NumDict(CodeName)("COLGAAPDone")

    Set rng = Sheet12.Range("DoneGAAP")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
    Next cell

End If





cnn.Close
GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
MsgBox "There was an error in the CancelBatch Routine. " & Err, vbOKOnly

9999:

End Function


Public Sub GLCheck()
If Sheet2.Range("ErrorCtl").Value = True Then
On Error GoTo errexit
End If

Dim frm As DataRetrievalStatus
Set frm = New DataRetrievalStatus
frm.Label1.Caption = "Getting Security Info......"
frm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
frm.Show vbModeless
DoEvents

Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim DataSQL As String
Dim Datars As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Dim i As Integer
Dim intTotalElements As Integer
Dim intCount As Integer

'Open Connection'
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' GET DEPARTMENTS

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPGLCheck"
cmd.CommandTimeout = 60

Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 3, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamOutput)
cmd.Parameters.Append cmdMonth

cmd.Execute

cnn.Close

Sheet2.Range("LastClosedMth").Value = cmd.Parameters("@Month").Value

If Sheet2.Range("LastClosedMth").Value >= Sheet17.Range("StartMonth").Value Then
    With Sheet17.Range("StartMonth").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheet17.Range("StartMonth").Offset(0, 1).Value = "Closed!"
Else
    With Sheet17.Range("StartMonth").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheet17.Range("StartMonth").Offset(0, 1).Value = ""
End If


GoTo 9999
errexit:
If Datars.State <> 0 Then
    Datars.Close
End If
Unload frm
MsgBox "There was an error in the GLCheck Routine. " & Err, vbOKOnly
Application.EnableEvents = True
9999:

Unload frm

End Sub


Public Sub ApprCheck2()
If Sheet2.Range("ErrorCtl").Value = True Then
On Error GoTo errexit
End If

Dim Co As String
Dim Mo As String
Dim Dept As String
Dim ApprByAccounting As String
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim Datars As ADODB.Recordset
Dim objStrSQL
Set cnn = New ADODB.Connection
Set cmd = New ADODB.Command
Set Datars = New ADODB.Recordset


Co = CStr(Sheet17.Range("StartCompany").Value)
Mo = CStr(Sheet17.Range("StartMonth").Value)
Dept = CStr(Sheet17.Range("StartDept").Value)

'Open Connection'
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPApprCheck"
cmd.CommandTimeout = 180

Set cmdCo = cmd.CreateParameter("@Co", adTinyInt, adParamInput, 20, Co)
cmd.Parameters.Append cmdCo
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, CDate(Mo))
cmd.Parameters.Append cmdMonth
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 30, Dept)
cmd.Parameters.Append cmdDept

cnn.CursorLocation = adUseClient

Datars.Open cmd.Execute

If Datars.EOF <> True Then
    Datars.MoveLast
    Datars.MoveFirst
    
    T = Datars.RecordCount
    Dim i As Integer
       
    For i = 1 To T
    
        Sheet2.Range("ReadyForOpsAppr1").Value = Datars.Fields("ReadyForOpsAppr1").Value
        If Datars.Fields("ReadyForOpsAppr1").Value = "Y" Then
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOn
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
        Else
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff

        End If
        
        Sheet2.Range("FinalAppr").Value = Datars.Fields("FinalAppr").Value
        If Datars.Fields("FinalAppr").Value = "Y" Then
            Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOn
            Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOff
        Else
            Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff

        End If
        
        Sheet2.Range("AcctAppr").Value = Datars.Fields("AcctAppr").Value
        If Datars.Fields("AcctAppr").Value = "Y" Then
            Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOn
            Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOff
        Else
            Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff

        End If

    Next i

End If

Datars.Close
cnn.Close
GoTo 9999
errexit:
If Datars.State <> 0 Then
    Datars.Close
End If
cnn.Close
MsgBox "There was an error in the ApprCheck2 Routine. " & Err, vbOKOnly
9999:

End Sub


Public Sub SetTotals(r As Integer, StartRow As Integer, EndRow As Integer, Range As String, sh As Worksheet)
Dim Col As String
Dim colNum As String
Col = LetDict(sh.CodeName)(Range)
colNum = NumDict(sh.CodeName)(Range)

sh.Range("SummaryData").Cells(r, CInt(colNum)).Formula = "=Sum(" & Col & CStr(StartRow) & ":" & Col & CStr(EndRow) & ")"
sh.Range("SummaryData").Cells(r, CInt(colNum)).Font.Bold = True

End Sub

Public Sub SetDTotals(r As Integer, StartRow As Integer, EndRow As Integer, Range As String, sh As Worksheet)
Dim Col As String
Dim colNum As String

Col = LetDict(sh.CodeName)(Range)
colNum = NumDict(sh.CodeName)(Range)

sh.Range("SummaryData").Cells(r, CInt(colNum)).Formula = "=Sum(" & Col & CStr(StartRow) & ":" & Col & CStr(EndRow) & ")/2"
sh.Range("SummaryData").Cells(r, CInt(colNum)).Font.Bold = True

End Sub



Public Sub SetGrandTotals(r As Integer, OpenRowTotal As Integer, EndRow As Integer, Range As String, sh As Worksheet)
Dim Col As String
Dim colNum As String
Dim ContStat As String
Dim Div As String
Col = LetDict(sh.CodeName)(Range)
colNum = NumDict(sh.CodeName)(Range)
ContStat = LetDict(sh.CodeName)("COLZContractStatus")

If Sheet2.Range("SortOption").Value = 1 Then
    ' Open Job Totals
    sh.Range("SummaryData").Cells(r, CInt(colNum)).Formula = "=SUMIF(" & ContStat & "8:" & ContStat & EndRow & ",1, " & Col & "8" & ":" & Col & EndRow & ")"
    sh.Range("SummaryData").Cells(r, CInt(colNum)).Font.Bold = True
    
    ' Closed Job Totals
    sh.Range("SummaryData").Cells(r + 1, CInt(colNum)).Formula = "=SUMIF(" & ContStat & "8:" & ContStat & EndRow & ",2, " & Col & "8" & ":" & Col & EndRow & ")"
    sh.Range("SummaryData").Cells(r + 1, CInt(colNum)).Font.Bold = True
    Div = "4"
Else
    Div = "2"
End If


' Grand Total
sh.Range("SummaryData").Cells(r + 2, CInt(colNum)).Formula = "=Sum(" & Col & "8" & ":" & Col & EndRow + 1 & ")/" & Div
sh.Range("SummaryData").Cells(r + 2, CInt(colNum)).Font.Bold = True


End Sub


'   pass the column letter and returns the number of the column
Function ColLetToNum(InputLetter As String) As Integer
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If


Dim OutputNumber As Integer
Dim Leng As Integer
Dim i As Integer
On Error GoTo errexit
Leng = Len(InputLetter)
OutputNumber = 0

For i = 1 To Leng
   OutputNumber = (Asc(VBA.UCase(VBA.Mid(InputLetter, i, 1))) - 64) + OutputNumber * 26
Next i

ColLetToNum = OutputNumber  'Output the corresponding number
GoTo 9999
errexit:
MsgBox "There was an error in the ColLetToNum Routine. " & Err, vbOKOnly
9999:
End Function


Public Sub ProtectUnProtect(OnOff As String)

Application.ScreenUpdating = False

If OnOff = "On" Then

    If Sheet2.Range("ProtectSheet").Value = "True" Then
        Sheet11.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet12.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet13.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet14.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet15.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    
    Else
        Sheet11.Unprotect "password"
        Sheet12.Unprotect "password"
        Sheet13.Unprotect "password"
        Sheet14.Unprotect "password"
        Sheet15.Unprotect "password"
        
    End If

Else
    
    Sheet11.Unprotect "password"
    Sheet12.Unprotect "password"
    Sheet13.Unprotect "password"
    Sheet14.Unprotect "password"
    Sheet15.Unprotect "password"

End If

9999:

'Application.ScreenUpdating = True


End Sub