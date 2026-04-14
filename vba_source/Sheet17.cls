
Private Sub Worksheet_Change(ByVal Target As Range)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
Application.EnableEvents = False


' Company
If FindIntersect(Range(Target.Address), Me.Range("StartCompany")) = "Yes" Then
    
    Application.EnableEvents = False
    ' This checks list of accessable Companies returned during GetSecurity Check on Workbook open
    Sheet17.Range("StartCoName").Value = GetCoName(Target.Value)
    If Sheet17.Range("StartCoName").Value = "" Then
        Target.Value = ""
    End If
    Application.EnableEvents = True
    
    GoTo 99

End If

' Month
If FindIntersect(Range(Target.Address), Me.Range("StartMonth")) = "Yes" Then
    
    If IsValidDateRegex(Target.Value) = True Then
        

        Application.ScreenUpdating = False
        
        Me.Unprotect "password"
        Me.Range("StartDept").Value = ""
        
        If IsDate(Target.Value) Then
            Target.Value = DateSerial(Year(Target.Value), Month(Target.Value), 1)
        End If
        
        ClearForms3
        ' Check status of the GL Month
        GLCheck
        
        Dim NoData As Boolean
        NoData = False
        
        
        
        ' Check if WIPDetail has data
        UseExistingBatch NoData
            
        If NoData = False Then
            GetWipDetail2 Sheet11
            GetWipDetail2 Sheet12
            GetudWIPJVSub Sheet14
            GetudWIPJVSub Sheet15
        End If
        
        Sheet11.Activate
        Sheet11.Range("A2").Select
        Sheet12.Activate
        Sheet12.Range("A2").Select
        Sheet17.Activate
        Sheet17.Range("StartMonth").Select
        
        GoTo 99
        
    Else
        

        MsgBox "Enter a valid date. ", vbOKOnly
        On Error Resume Next
        Application.Undo
        On Error GoTo 0
        Target.Select
        GoTo 99
        
    
    End If

End If


GoTo 9999
errexit:
MsgBox "There was an error in the Worksheet_Change Routine. " & Err, vbOKOnly
9999:
If Sheet2.Range("ProtectSheet").Value = "True" Then
    Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If
Application.ScreenUpdating = True

99:
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub


Function IsValidDateRegex(inputValue As Variant) As Boolean
    On Error GoTo ErrHandler
    
    ' Check for empty or invalid inputs
    If IsEmpty(inputValue) Or IsNull(inputValue) Or Trim(inputValue) = "" Then
        IsValidDateRegex = False
        Exit Function
    End If
    
    ' Convert input to string
    Dim strInput As String
    strInput = CStr(inputValue)
    
    ' Set up regex for #/#/####, #/#/##, ##/##/####, ##-##-####, #-#-####, #-#-##
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Pattern = "^(0?[1-9]|1[0-2])[-/](0?[1-9]|[12][0-9]|3[01])[-/]([0-9]{2}|[0-9]{4})$"
        .IgnoreCase = True
        .Global = False
    End With
    
    ' Check if input matches the pattern
    If Not regEx.Test(strInput) Then
        IsValidDateRegex = False
        Set regEx = Nothing
        Exit Function
    End If
    
    ' Validate as a true date using IsDate
    If IsDate(strInput) Then
        Dim testDate As Date
        testDate = CDate(strInput)
        ' Ensure year is within Excel's valid range (1900–9999)
        If Year(testDate) >= 1900 And Year(testDate) <= 9999 Then
            IsValidDateRegex = True
        Else
            IsValidDateRegex = False
        End If
    Else
        IsValidDateRegex = False
    End If
    
    Set regEx = Nothing
    Exit Function

ErrHandler:
    IsValidDateRegex = False
    Set regEx = Nothing
End Function
