'Option Explicit


Public Function PadString(str As String, Padto As Integer, Side As String) As String
On Error GoTo errexit
Dim PadStr As String
Dim i As Integer
PadStr = str

For i = 1 To Padto - Len(str)

    If Side = "L" Then
        PadStr = " " & PadStr
    End If

    If Side = "R" Then
        PadStr = PadStr & " "
    End If

Next

PadString = PadStr
GoTo 9999
errexit:
MsgBox ("There was an Error in the PadString Routine")
9999:
End Function


Public Sub OpenForm()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If



Application.ScreenUpdating = False
ProtectUnProtect ("Off")
Dim frm As BatchSelection

Dim ws As Excel.Worksheet


If FindIntersect(Range(ActiveCell.Address), Range("StartDept")) = "Yes" Then
    On Error GoTo 9999
    Dim Deptfrm As Dept
    Set Deptfrm = New Dept
    Deptfrm.StartUpPosition = 0
    ' Calculate the position where the UserForm should appear to be centered in Excel
    Deptfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Deptfrm.Width)
    Deptfrm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Deptfrm.Height)
    Deptfrm.Show
    'Dept.Show
    'BatchCheck
Else
    If FindIntersect(Range(ActiveCell.Address), Range("StartCompany")) = "Yes" Then
        Dim Companyfrm As Company
        Set Companyfrm = New Company
        Companyfrm.StartUpPosition = 0
        ' Calculate the position where the UserForm should appear to be centered in Excel
        Companyfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Companyfrm.Width)
        Companyfrm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Companyfrm.Height)
        Companyfrm.Show
        'Company.Show
    End If
   

End If


GoTo 9999
errexit:
MsgBox "There was an error in the OpenForm Routine. " & Err, vbOKOnly
9999:
ProtectUnProtect ("On")
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub


'   finds if ranges intersect with one another
Public Function FindIntersect(Target As Range, rng1 As Range, Optional rng2 As Range, Optional rng3 As Range, Optional rng4 As Range, Optional rng5 As Range, Optional rng6 As Range, Optional rng7 As Range)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
If rng2 Is Nothing Then Set rng2 = rng1
If rng3 Is Nothing Then Set rng3 = rng1
If rng4 Is Nothing Then Set rng4 = rng1
If rng5 Is Nothing Then Set rng5 = rng1
If rng6 Is Nothing Then Set rng6 = rng1
If rng7 Is Nothing Then Set rng7 = rng1

If Application.Intersect(Target, Union(rng1, rng2, rng3, rng4, rng5, rng6, rng7)) Is Nothing Then
    FindIntersect = "No"
Else
    FindIntersect = "Yes"
End If

GoTo 9999
errexit:
MsgBox ("There was an Error in the FindIntersect Routine")
9999:
End Function


Public Function GetCoName(Co As Integer) As String

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
cmd.CommandText = "LCGWIPGetCoName1"
cmd.CommandTimeout = 180


Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 3, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdUserName = cmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, Sheet2.Range("UserName2").Value)
cmd.Parameters.Append cmdUserName
Set cmdRcode = cmd.CreateParameter("@Rcode", adInteger, adParamOutput, 3)
cmd.Parameters.Append cmdRcode
Set cmdName = cmd.CreateParameter("@Name", adVarChar, adParamOutput, 60)
cmd.Parameters.Append cmdName

cmd.Execute

If cmd.Parameters("@Rcode").Value = 0 Then

    GetCoName = cmd.Parameters("@Name").Value
    
Else

    MsgBox "Check Company Security or your Connection to the Server", vbInformation
    GetCoName = ""

End If

cnn.Close


Unload frm
    

GoTo 999
errexit:
MsgBox "There was an error in the GetCoName Function", vbInformation
999:
End Function



Function IsInCommaDelimitedList(checkValue As Variant, listCell As Range) As Boolean
    Dim listArray() As String
    Dim item As Variant
    Dim i As Integer
    
    ' Convert the comma-delimited string into an array
    listArray = Split(listCell.Value, ",")
    
    ' Trim each element to remove leading/trailing spaces
    For i = LBound(listArray) To UBound(listArray)
        listArray(i) = Trim(listArray(i))
    Next i
    
    ' Check if the value is in the array
    For Each item In listArray
        If item = CStr(checkValue) Then ' Ensure comparison is string to string
            IsInCommaDelimitedList = True
            Exit Function
        End If
    Next item
    
    ' If the loop completes without finding the value, return False
    IsInCommaDelimitedList = False
End Function


