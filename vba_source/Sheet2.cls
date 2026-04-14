Private Sub Worksheet_Change(ByVal Target As Range)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
Application.EnableEvents = False


If FindIntersect(Range(Target.Address), Me.Range("ProtectSheet")) = "Yes" Then

    If Target.Value = True Then
        ProtectUnProtect ("On")
    Else
        ProtectUnProtect ("Off")
    End If
    
End If
Application.EnableEvents = True

errexit:
End Sub