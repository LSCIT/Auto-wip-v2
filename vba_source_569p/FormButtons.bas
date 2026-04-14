Attribute VB_Name = "FormButtons"
Sub RFOYes_Click()
' ACCT - Ready For Ops (Yes)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If Len(Sheet17.Range("StartCompany").Value) <> 0 And Len(Sheet17.Range("StartMonth").Value) <> 0 And Len(Sheet17.Range("StartDept").Value) <> 0 Then

    Select Case Sheet2.Range("Role").Value
    
    
        Case "":
    
    
        Case "WIPAccounting":
        
                Sheet2.Range("ReadyForOpsAppr1").Value = "Y"
                Sheet2.Range("SendAppr").Value = "True"
                Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
                Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOn
                UpdateApprovals
                
        Case Else:
        
            MsgBox "Only Accounting Can Change This Setting", vbInformation
            ' Set Ready For Ops to No
            Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
        
    End Select

Else
    
    Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
    Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
    MsgBox "Select Company, Month And Division to Create new WIP Month", vbInformation
End If

Sheet17.Activate
Sheet17.Range("StartCompany").Select

GoTo 9999
errexit:
Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
MsgBox "There was an error in the RFOYes_Click Routine. " & Err, vbOKOnly

9999:

End Sub


Sub RFONo_Click()
' ACCT - Ready For Ops (No)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

' Once Ready for Ops is set, it cannot be manually undone.
' Only CompleteCheck failure in the Yes handler can reset it.
If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then
    MsgBox "Ready for Ops has been set and cannot be changed.", vbInformation, "Ready for Ops"
    Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
    Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOn
    GoTo 9999
End If

If Sheet2.Range("Role").Value = "WIPAccounting" Then

    If Sheet17.Range("StartCompany").Value <> "" And Sheet17.Range("StartMonth").Value <> "" And Sheet17.Range("StartDept").Value <> "" Then
    
        If MsgBox("This will delete all WIP Information for the Company, Month and Divisions Selected.  Are you sure?", vbYesNo) = vbYes Then
        
            Select Case Sheet2.Range("Role").Value
            
                Case "":
            
            
                Case "WIPAccounting":
                    Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
                    Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
                    Sheet2.Range("ReadyForOpsAppr1").Value = "N"
                    Sheet2.Range("FinalAppr").Value = "N"
                    Sheet2.Range("SendAppr").Value = "True"
                    UpdateApprovals
                    ClearWIPDetailTable
                
                
                Case Else:
                
                    MsgBox "Only Accounting Can Change This Setting", vbInformation
                    If Sheet2.Range("ReadyForOpsAppr1").Value = "N" Then
                        Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
                        Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
                    
                    Else
                        Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
                        Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOn
                    
                    End If
                    
            End Select
        
        Else
            If Sheet2.Range("ReadyForOpsAppr1").Value = "N" Then
                Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
                Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
            
            Else
                Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
                Sheet17.Shapes("RFO-es").ControlFormat.Value = xlOn
            
            End If
        
        End If
    
    Else
        MsgBox "Select Company and Month to Clear Existing Data", vbInformation
    End If


Else

    MsgBox "Only Accounting Can Change this setting", vbInformation
    
    If Sheet2.Range("ReadyForOpsAppr1").Value = "N" Then
        Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOn
        Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOff
    
    Else
        Sheet17.Shapes("RFO-No").ControlFormat.Value = xlOff
        Sheet17.Shapes("RFO-Yes").ControlFormat.Value = xlOn
    
    End If



End If

GoTo 9999
errexit:
MsgBox "There was an error in the RFONo_Click Routine. " & Err, vbOKOnly
9999:
End Sub


Sub OFANo_Click()
' Final Approval (No)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

' Once Ops Final Approval is set, it cannot be manually undone.
' Only CompleteCheck failure in the Yes handler can reset it.
If Sheet2.Range("FinalAppr").Value = "Y" Then
    MsgBox "Ops Final Approval has been set and cannot be changed.", vbInformation, "Ops Final Approval"
    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOff
    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOn
    GoTo 9999
End If

If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then
    
    Select Case Sheet2.Range("Role").Value
    
    
        Case "":
        
        Case "WIPLevel2":
    
            ' Set Final Approval to No
            Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
            Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
            Sheet2.Range("FinalAppr").Value = "N"
            Sheet2.Range("SendAppr").Value = "True"
            UpdateApprovals
    
        Case "WIPAccounting":
        
            ' Set Final Approval to No
            Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
            Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
            Sheet2.Range("FinalAppr").Value = "N"
            Sheet2.Range("SendAppr").Value = "True"
            UpdateApprovals
    
    
        Case Else:
        
            MsgBox "Only Final Approvers Can Change This Setting", vbInformation
            Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOn
            Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOff

    
    End Select
    
Else

    MsgBox "Period Not Ready for Ops Input", vbOKOnly, "Contact Accounting"
    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOn
    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOf
    Sheet2.Range("FinalAppr").Value = "N"

End If

GoTo 9999
errexit:
MsgBox "There was an error in the OFANo_Click Routine. " & Err, vbOKOnly
9999:
Sheet17.Activate

End Sub

Sub OFAYes_Click()
' Final Approval (Yes)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then

    'If Sheet2.Range("InitAppr").Value = "Y" Then
    
        Select Case Sheet2.Range("Role").Value
        
            Case "":
        
            Case "WIPLevel2":
        
                If CompleteCheck("O", "") = True Then
                    ' Set Final Approval to Yes
                    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOff
                    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOn
                    Sheet2.Range("FinalAppr").Value = "Y"
                    UpdateApprovals
                
                Else
                    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
                    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff

                End If
                
            Case "WIPAccounting":
        
                If CompleteCheck("O", "") = True Then
                    ' Set Final Approval to Yes
                    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOff
                    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOn
                    Sheet2.Range("FinalAppr").Value = "Y"
                    UpdateApprovals
                Else
                    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
                    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
                End If
        
        
            Case Else:
            
                MsgBox "Only Final Approvers Can Change This Setting", vbInformation
                Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
                Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
    
        
        End Select
        
    
    
    
Else

    MsgBox "Period Not Ready for Ops Input", vbInformation, "Contact Accounting"
    Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
    Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
    Sheet2.Range("FinalAppr").Value = "N"

End If

GoTo 9999
errexit:
Sheet17.Shapes("OFA-No").ControlFormat.Value = xlOn
Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff
MsgBox "There was an error in the OFAYes_Click Routine. " & Err, vbOKOnly
9999:

Sheet17.Activate

End Sub


Sub AFANo_Click()
' ACCT - Final Approval (No)
' Once Accounting Final Approval is set, batch is immutable.
If Sheet2.Range("AcctAppr").Value = "Y" Then
    MsgBox "Accounting Final Approval has been set. The batch is locked.", vbInformation, "Accounting Final Approval"
    Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOff
    Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOn
    GoTo 9999
End If

If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'ProtectUnProtect ("Off")
Sheet7.Activate

Select Case Sheet2.Range("Role").Value


    Case "":


    Case "WIPAccounting":

        Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
        Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
        Sheet2.Range("AcctAppr").Value = "N"
        Sheet2.Range("SendAppr").Value = "True"
        UpdateApprovals

    Case Else:
    
        MsgBox "Only Accounting Can Change This Setting", vbInformation

        Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
        Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
        

End Select

GoTo 9999
errexit:
MsgBox "There was an error in the AFANo_Click Routine. " & Err, vbOKOnly
9999:

Sheet17.Activate

'ProtectUnProtect ("On")
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub


Sub AFAYes_Click()
' ACCT - Final Approval (Yes)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Select Case Sheet2.Range("Role").Value
    Case "":
    Case "WIPAccounting":

        If Sheet2.Range("FinalAppr").Value = "Y" Then
            If CompleteCheck("G", "") = True Then
                Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOff
                Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOn
                Sheet2.Range("AcctAppr").Value = "Y"
                Sheet2.Range("SendAppr").Value = "True"
                UpdateApprovals
            Else
                ' Reset radio button — CompleteCheck failed
                Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
                Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
            End If
            
        Else
        
            MsgBox "Need Ops Approval", vbInformation, "Waiting On Ops Final Approval"
            Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
            Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
        
        End If

    Case Else:
    
        MsgBox "Only Accounting Can Change This Setting", vbInformation

        Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
        Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
        Sheet2.Range("AcctAppr").Value = "N"
        Sheet2.Range("SendAppr").Value = "True"

End Select



GoTo 9999
errexit:
Sheet17.Shapes("AFA-No").ControlFormat.Value = xlOn
Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff
MsgBox "There was an error in the AFAYes_Click Routine. " & Err, vbOKOnly
9999:

Sheet17.Activate
End Sub


' ========================================
' SaveDistributeClick — Start sheet button
' Saves a .xlsm copy to C:\Trusted\ for Ops PMs
' ========================================
Sub SaveDistributeClick()
    LylesWIPData.DistributeToOps
End Sub


' ========================================
' PushToVistaClick — Start sheet button
' Pushes approved GAAP/OPS overrides to Vista override tables.
' Requires ALL three approvals: RFO=Yes, OFA=Yes, AFA=Yes.
' Only available on GAAP quarter months (Mar/Jun/Sep/Dec).
'
' Currently DISABLED for demo/testing — shows confirmation that
' guards passed but does not execute the write.
' To enable: uncomment the WriteBackToVista call below.
' ========================================
Sub PushToVistaClick()
    On Error GoTo errexit

    ' Require company/month/dept
    If Len(Sheet17.Range("StartCompany").Value) = 0 Or _
       Len(Sheet17.Range("StartMonth").Value) = 0 Or _
       Len(Sheet17.Range("StartDept").Value) = 0 Then
        MsgBox "Select Company, Month and Division first.", vbInformation
        Exit Sub
    End If

    ' Role check — only Accounting can push to Vista
    If Sheet2.Range("Role").Value <> "WIPAccounting" Then
        MsgBox "Only Accounting can push data to Vista.", vbInformation
        Exit Sub
    End If

    ' All three approvals must be set
    If Sheet2.Range("ReadyForOpsAppr1").Value <> "Y" Then
        MsgBox "Ready for Ops must be set before pushing to Vista.", _
               vbInformation, "Not Ready"
        Exit Sub
    End If

    If Sheet2.Range("FinalAppr").Value <> "Y" Then
        MsgBox "Ops Final Approval must be set before pushing to Vista.", _
               vbInformation, "Ops Not Approved"
        Exit Sub
    End If

    If Sheet2.Range("AcctAppr").Value <> "Y" Then
        MsgBox "Accounting Final Approval must be set before pushing to Vista.", _
               vbInformation, "Not Yet Approved"
        Exit Sub
    End If

    ' Quarterly check
    Dim wipMonth As Date
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)
    If Month(wipMonth) Mod 3 <> 0 Then
        MsgBox "Vista write-back is only available on GAAP quarter months " & _
               "(March, June, September, December).", vbInformation, "Not a GAAP Quarter"
        Exit Sub
    End If

    ' === DEMO/TEST MODE: Show message but do NOT execute write ===
    MsgBox "All guards passed. Push to Vista is disabled during testing." & vbCrLf & vbCrLf & _
           "Company: " & Sheet17.Range("StartCompany").Value & vbCrLf & _
           "Month: " & Format(wipMonth, "mmmm yyyy") & vbCrLf & _
           "Division: " & Sheet17.Range("StartDept").Value & vbCrLf & vbCrLf & _
           "When enabled, this will write approved GAAP/OPS overrides " & _
           "to the Vista override tables.", _
           vbInformation, "Push to Vista (Test Mode)"
    Exit Sub
    ' === END DEMO/TEST MODE ===

    ' --- Production code (uncomment when ready to go live) ---
    'Dim co As Integer
    'co = CInt(Sheet17.Range("StartCompany").Value)
    'Dim dept As String
    'dept = CStr(Sheet17.Range("StartDept").Value)
    'Dim userName As String
    'userName = CStr(Sheet2.Range("UserName2").Value)
    'If userName = "" Then userName = Environ("UserName")
    'Dim result As String
    'result = LylesWIPData.WriteBackToVista(co, wipMonth, dept, userName)
    'If result <> "" Then
    '    MsgBox result, vbInformation, "Push to Vista Complete"
    'End If

    Exit Sub

errexit:
    MsgBox "There was an error in the PushToVistaClick Routine: " & _
           Err.Description, vbOKOnly, "Push to Vista Error"
End Sub


Sub SetJVRolePermissions(sh As Worksheet)


If Sheet2.Range("Role").Value <> "WIPAccounting" Then

    sh.Range("JTDBill").Locked = True
    'Revenue Color
    With sh.Range("JTDBill").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With sh.Range("JTDBillHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    
    sh.Range("OPCCA").Locked = True
    'Revenue Color
    With sh.Range("OPCCA").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With sh.Range("OPCCAHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

    sh.Range("JVJTDEarnedRev").Locked = True
    'Revenue Color
    With sh.Range("JVJTDEarnedRev").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With sh.Range("JVJTDEarnedRevHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    
    sh.Range("JVJTDC").Locked = True
    ' Cost Color
    With sh.Range("JVJTDC").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With sh.Range("JVJTDCHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    
    sh.Range("JVJTDD").Locked = True
    ' Distribution Color
    With sh.Range("JVJTDD").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With sh.Range("JVJTDDHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    

Else
    
    sh.Range("JTDBill").Locked = False
    ' Yellow
    With sh.Range("JTDBill").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With sh.Range("JTDBillHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    sh.Range("OPCCA").Locked = False
    ' Yellow
    With sh.Range("OPCCA").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With sh.Range("OPCCAHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    sh.Range("JVJTDEarnedRev").Locked = False
    ' Yellow
    With sh.Range("JVJTDEarnedRev").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With sh.Range("JVJTDEarnedRevHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    sh.Range("JVJTDC").Locked = False
    ' Yellow
    With sh.Range("JVJTDC").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With sh.Range("JVJTDCHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    sh.Range("JVJTDD").Locked = False
    ' Yellow
    With sh.Range("JVJTDD").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With sh.Range("JVJTDDHeading").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End If

If Sheet2.Range("ProtectSheet").Value = True Then
    sh.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If


End Sub




