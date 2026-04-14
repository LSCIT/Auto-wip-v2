Private Sub GetVPData_Click()
On Error GoTo errexit
 

password = TBPassword.Value

If password = Sheet2.Range("YrEndPW").Value Then
    Application.EnableEvents = False

    Sheet7.Range("SendYrEndData").Value = "Yes"

    Application.EnableEvents = True


End If



Unload YREndLogin
 
GoTo 9999
errexit:
MsgBox ("There was an Error in the GetVPData_Click Routine")
9999:
 
End Sub




Private Sub UserForm_Terminate()
Application.EnableEvents = False
ProtectUnProtect ("Off")
If TBPassword.Value = Sheet2.Range("YrEndPW").Value Then

    Sheet7.Range("SendYrEndData").Value = "Yes"
    
    Else
    
    Sheet7.Range("SendYrEndData").Value = "No"

    End If
Application.EnableEvents = True
ProtectUnProtect ("On")
Unload YREndLogin

End Sub