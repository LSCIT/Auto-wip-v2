Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
If Sheet2.Range("Role").Value = "WipViewOnly" Then
    MsgBox "You do not have permission to edit.", vbOKOnly, "View Olny Permission"
    GoTo 9999:
End If
    
If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If
' JTD Show Hide
If FindIntersect(Range(Target.Address), Me.Range("JTDHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("JTDPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLJVOJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJVOJTDProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLJVOJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJVOJTDProfit")).Hidden = False
        'Me.Columns(NumDict(Me.CodeName)("COLJTDBonusProfitNotes")).Hidden = True
        Me.Range("JTDHeading").Select
        Me.Range("JTDPlus").Value = "-"
        Me.Range("JTDPlus").Offset(1, 0).Value = ""
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLJVOJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJVOJTDProfit")).Hidden = True
        Me.Range("JTDPlus").Select
        Me.Range("JTDPlus").Value = "+"
        Me.Range("JTDPlus").Offset(1, 0).Value = "J T D"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If


' APY Show Hide
If FindIntersect(Range(Target.Address), Me.Range("APYHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("APYPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLJVAPYRev") & ":" & LetDict(Me.CodeName)("COLJVAPYCalcProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLJVAPYRev") & ":" & LetDict(Me.CodeName)("COLJVAPYCalcProfit")).Hidden = False
        Me.Range("APYHeading").Select
        Me.Range("APYPlus").Value = "-"
        Me.Range("APYPlus").Offset(1, 0).Value = ""
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLJVAPYRev") & ":" & LetDict(Me.CodeName)("COLJVAPYCalcProfit")).Hidden = True
        Me.Range("APYPlus").Select
        Me.Range("APYPlus").Value = "+"
        Me.Range("APYPlus").Offset(1, 0).Value = "A P Y"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If


' CY Show Hide
If FindIntersect(Range(Target.Address), Me.Range("CYHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("CYPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLJVCYRev") & ":" & LetDict(Me.CodeName)("COLJVCYCalcProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLJVCYRev") & ":" & LetDict(Me.CodeName)("COLJVCYCalcProfit")).Hidden = False
        Me.Range("CYHeading").Select
        Me.Range("CYPlus").Value = "-"
        Me.Range("CYPlus").Offset(1, 0).Value = ""
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLJVCYRev") & ":" & LetDict(Me.CodeName)("COLJVCYCalcProfit")).Hidden = True
        Me.Range("CYPlus").Select
        Me.Range("CYPlus").Value = "+"
        Me.Range("CYPlus").Offset(1, 0).Value = "C Y"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If

' B Show Hide
If FindIntersect(Range(Target.Address), Me.Range("BHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("BPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLJVBRev") & ":" & LetDict(Me.CodeName)("COLJVBProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLJVBRev") & ":" & LetDict(Me.CodeName)("COLJVBProfit")).Hidden = False
        Me.Range("BHeading").Select
        Me.Range("BPlus").Value = "-"
        Me.Range("BPlus").Offset(1, 0).Value = ""
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLJVBRev") & ":" & LetDict(Me.CodeName)("COLJVBProfit")).Hidden = True
        Me.Range("BPlus").Select
        Me.Range("BPlus").Value = "+"
        Me.Range("BPlus").Offset(1, 0).Value = "B L"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If



' ALL Show Hide
If FindIntersect(Range(Target.Address), Me.Range("ALLPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Range("ALLPlus").Value = "+" Then
        Me.Columns(LetDict(Me.CodeName)("COLJVOJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJVOJTDProfit")).Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLJVAPYRev") & ":" & LetDict(Me.CodeName)("COLJVAPYCalcProfit")).Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLJVCYRev") & ":" & LetDict(Me.CodeName)("COLJVCYCalcProfit")).Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLJVBRev") & ":" & LetDict(Me.CodeName)("COLJVBProfit")).Hidden = False
        
        Me.Range("JTDPlus").Value = "-"
        Me.Range("ALLPlus").Value = "-"
        Me.Range("APYPlus").Value = "-"
        Me.Range("APYPlus").Value = "-"
        Me.Range("CYPlus").Value = "-"
        Me.Range("BPlus").Value = "-"
        Me.Range("JTDPlus").Offset(1, 0).Value = ""
        Me.Range("APYPlus").Offset(1, 0).Value = ""
        Me.Range("CYPlus").Offset(1, 0).Value = ""
        Me.Range("BPlus").Offset(1, 0).Value = ""
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLJVOJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJVOJTDProfit")).Hidden = True
        Me.Columns(LetDict(Me.CodeName)("COLJVAPYRev") & ":" & LetDict(Me.CodeName)("COLJVAPYCalcProfit")).Hidden = True
        Me.Columns(LetDict(Me.CodeName)("COLJVCYRev") & ":" & LetDict(Me.CodeName)("COLJVCYCalcProfit")).Hidden = True
        Me.Columns(LetDict(Me.CodeName)("COLJVBRev") & ":" & LetDict(Me.CodeName)("COLJVBProfit")).Hidden = True
        
        Me.Range("ALLPlus").Value = "+"
        Me.Range("JTDPlus").Value = "+"
        Me.Range("APYPlus").Value = "+"
        Me.Range("CYPlus").Value = "+"
        Me.Range("BPlus").Value = "+"
        Me.Range("JTDPlus").Offset(1, 0).Value = "J T D"
        Me.Range("APYPlus").Offset(1, 0).Value = "A P Y"
        Me.Range("CYPlus").Offset(1, 0).Value = "C Y"
        Me.Range("BPlus").Offset(1, 0).Value = "B L"
    
    
    End If
    Me.Range("AllPlus").Select
    Application.EnableEvents = True

GoTo 99
End If








If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then
    
    'ProtectUnProtect ("Off")
    Me.Unprotect "password"
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
        
    
    ' Done
    If FindIntersect(Range(Target.Address), Me.Range("Done")) = "Yes" Then
       If Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff Then
            Application.EnableEvents = False
            Application.ScreenUpdating = Fales
            Application.Calculation = xlCalculationManual
            
            If Target.Value = "P" Then
                Target.Value = ""
                Sheet15.Range(LetDict("Sheet15")("COLDone") & Target.row).Value = ""
                Target.ClearComments
            Else
                Target.Value = "P"
                Sheet15.Range(LetDict("Sheet15")("COLDone") & Target.row).Value = "P"
            End If
            
            Me.Range(LetDict(Me.CodeName)("COLJVZudChg") & Target.row).Value = "T"
            Me.Range(LetDict(Me.CodeName)("COLZUserName") & Target.row).Value = Sheet2.Range("UserName2").Value
            
            UpdateRowJV Target.row, Me
            
            
            
            ' Set Border to Thin
            With Target.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Target.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Target.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Target.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        
            
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            
        Else
    
            MsgBox "Cannot Update - Has been Final Approved For Ops", vbInformation
    
        End If
    GoTo 99
        
    End If
    
    ' JV JTD Earned Revenue
    If FindIntersect(Range(Target.Address), Range("JVJTDEarnedRevenue")) = "Yes" Then
        
        Sheet2.Range("SendJV").Value = "True"
        
        If Left(Target.Formula, 1) <> "=" Then
        
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            Target.Formula = "=IF(" & LetDict(Me.CodeName)("COLJVOvrCostProj") & Target.row & "<>0,IF((" & LetDict(Me.CodeName)("COLJVJTDCost") & Target.row _
            & "/" & LetDict(Me.CodeName)("COLJVOvrCostProj") & Target.row & ")>0.1,(" & LetDict(Me.CodeName)("COLJVJTDCost") & Target.row & "/" & _
            LetDict(Me.CodeName)("COLJVOvrCostProj") & Target.row & ")*" & LetDict(Me.CodeName)("COLJVOvrRevProj") & Target.row & ",0),0)"
            Me.Range(LetDict(Me.CodeName)("COLJVZudChg") & Target.row).Value = "T"
            Me.Range(LetDict(Me.CodeName)("COLJVZJTDER") & Target.row).Value = "R"

        Else

            If Me.Range(LetDict(Me.CodeName)("COLJVZJTDER2") & Target.row).Value <> 0 Then
                Target.Value = Me.Range(LetDict(Me.CodeName)("COLJVZJTDER2") & Target.row).Value
                Me.Range(LetDict(Me.CodeName)("COLJVZudChg") & Target.row).Value = ""
                Me.Range(LetDict(Me.CodeName)("COLJVZJTDER") & Target.row).Value = "T"
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If

        End If
            
        Target.Select
        Cancel = True
      
    End If
    

    
    If Sheet2.Range("ProtectSheet").Value = True Then
        Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
    End If

    Application.EnableEvents = True

Else
    MsgBox "Period Is Not Ready For Ops Input!", vbInformation
End If

99:
9999:

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
If Sheet2.Range("Role").Value = "WipViewOnly" Then
    MsgBox "You do not have permission to edit.", vbOKOnly, "View Olny Permission"
    Application.EnableEvents = False
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    Application.EnableEvents = True
    GoTo 9999:
End If

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If

Me.Unprotect "password"

If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then

    ' Marked Ops Done
    If Me.Cells(Target.row, NumDict(Me.CodeName)("COLDone")).Value = "P" Then
        On Error Resume Next
        Application.Undo
        On Error GoTo 0
        MsgBox "Uncheck the Done Checkbox to make changes", vbInformation
        GoTo 99
    End If
    
    
    
    ' Final Approval Yes Button
    If Sheet7.Shapes("Option Button 12").ControlFormat.Value = xlOff Then
    
        If Target.row >= 9 And Target.row <= 40 And Target.Column < 47 Then
                
            If Me.Range("A" & Target.row).Value <> "" Then
                
                ' ACCT - Final Approval (No)
                If Sheet7.Shapes("Option Button 12").ControlFormat.Value = xlOff Then
                    
                    Sheet2.Unprotect "password"
                    Sheet2.Range("SendJV").Value = "True"
                            
                End If
                
                ' JV JTD Earned Revenue
                If Target.Column = COLJVGJTDEarnedRev Then
                        Me.Range(COLLETJVZudChg & Target.row).Value = "T"
                        Me.Range(COLLETJVZGJTDER & Target.row).Value = "T"
                    
                    With Target.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                End If
                
                ' Ops JTD Earned Revenue
                If Target.Column = 17 Then
                        Me.Range(COLLETJVZudChg & Target.row).Value = "T"
                        Me.Range(COLLETJVZOJTDER & Target.row).Value = "T"
                    
                    With Target.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                End If
                
                ' Ops JTD Cost
                If Target.Column = 18 Then
                        Me.Range(COLLETJVZudChg & Target.row).Value = "T"
                        Me.Range(COLLETJVZOJTDCost & Target.row).Value = "T"
                    
                    With Target.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                End If
                
                               
                
                ' Other Columns
                If Target.Column = COLJVCurContAmt Or Target.Column = COLJVOvrRevProj Or Target.Column = COLJVOvrCostProj _
                Or Target.Column = COLJVCompDate Or Target.Column = COLJVOJTDBonusProfit Or Target.Column = COLJVGDist _
                Or Target.Column = COLJVBILLBillings Then
                        Me.Range(COLLETJVZudChg & Target.row).Value = "T"
                End If
                
                
                
            
           End If
            
        End If
     
    
        
    Else
    
        MsgBox "Month has Been Final Approved", vbInformation, "Contact Accounting"
        Application.EnableEvents = False
        On Error Resume Next
        Application.Undo
        On Error GoTo 0
        Application.EnableEvents = True
    
    
    End If
    
    
Else
    Application.EnableEvents = False
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    MsgBox "Period Is Not Ready For Ops Input", vbInformation
    Application.EnableEvents = True

End If


    
GoTo 10000
errexit:
MsgBox ("There was an automation error")

10000:
If Sheet2.Range("ProtectSheet").Value = "True" Then
    Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If

99:
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

9999:


End Sub

