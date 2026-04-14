'Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If



' returns full name of user from local computer
Public Function GetUserAccountInfo(MyDomain As String, username As String) As String
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

 
Dim objWMI As Object
Dim accounts As Object
Dim account As Object
 
Set objWMI = GetWMIService
  
Set accounts = objWMI.ExecQuery("Select * from Win32_UserAccount Where Domain = 'C01' And Name = '" & username & "'")

For Each account In accounts

GetUserAccountInfo = account.FullName

Next
GoTo 9999
errexit:
MsgBox "There was an error in the GetUserAccountInfo Routine. " & Err, vbOKOnly
9999:
End Function
 
Function GetWMIService() As Object
' http://msdn.microsoft.com/en-us/library/aa394586(VS.85).aspx
Dim strComputer As String
 
  strComputer = "."
 
  Set GetWMIService = GetObject("winmgmts:" _
                              & "{impersonationLevel=impersonate}!\\" _
                              & strComputer & "\root\cimv2")
End Function
