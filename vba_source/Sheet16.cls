Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then
    
    'ProtectUnProtect ("Off")
    Me.Unprotect "password"
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    
    ' JV JTD Earned Revenue
    If FindIntersect(Range(Target.Address), Range("JVJTDEarnedRev")) = "Yes" Then
        
        Sheet2.Range("SendJV").Value = "True"
        
        If Left(Target.Formula, 1) <> "=" Then
        
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            Target.Formula = "=IF(" & COLLETJVOvrCostProj & Target.row & "<>0,IF((" & COLLETJVGJTDCost & Target.row _
            & "/" & COLLETJVOvrCostProj & Target.row & ")>0.1,(" & COLJVGJTDCost & Target.row & "/" & COLJVOvrCostProj & Target.row & ")*" & COLJVOvrRevProj & Target.row & ",0),0)"
            Me.Range(COLLETJVZudChg & Target.row).Value = "T"
            Me.Range(COLLETJVZGJTDER & Target.row).Value = "R"
            'Me.Range("AR" & Target.row).Value = 0
        
        Else
        
            If Me.Range(COLLETJVZGJTDER2 & Target.row).Value <> 0 Then
                Target.Value = Me.Range(COLLETJVZGJTDER2 & Target.row).Value
                Me.Range(COLLETJVZudChg & Target.row).Value = ""
                Me.Range(COLLETJVZGJTDER & Target.row).Value = "T"
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        
        End If
        
      
    End If
    
    ' Ops JTD Earned Revenue
    If FindIntersect(Range(Target.Address), Range("OpsJTDEarnedRev")) = "Yes" Then
        
        If Left(Target.Formula, 1) <> "=" Then
        
            Sheet2.Range("SendJV").Value = "True"
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            '
            Target.Formula = "=IF(AND(" & COLLETJVOJTDPctComp & Target.row & "<>""," & COLLETJVOvrRevProj & Target.row & "<> "")," & COLLETJVOJTDPctComp & Target.row & "*" & COLLETJVOvrRevProj & Target.row & ",0)"
            Me.Range(COLLETJVZudChg & Target.row).Value = "T"
            Me.Range(COLLETJVZOJTDER & Target.row).Value = "T"
            
        Else
            
            If Me.Range(COLLETJVZOJTDER2 & Target.row).Value <> 0 Then
                Target.Value = Me.Range(COLLETJVZOJTDER2 & Target.row).Value
                Me.Range(COLLETJVZudChg & Target.row).Value = ""
                Me.Range(COLLETJVZOJTDER & Target.row).Value = "T"
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
          
        End If
    End If
    
    ' Ops JTD Cost
    If FindIntersect(Range(Target.Address), Range("OpsJTDCost")) = "Yes" Then
        
        If Left(Target.Formula, 1) <> "=" Then
        
            Sheet2.Range("SendJV").Value = "True"
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            '
            Target.Formula = "=" & COLLETJVGJTDCost & Target.row
            Me.Range(COLLETJVZudChg & Target.row).Value = "T"
            Me.Range(COLLETJVZOJTDCost & Target.row).Value = "T"
        
        Else
            
            If Me.Range(COLLETJVZOJTDCost2 & Target.row).Value <> 0 Then
                Target.Value = Me.Range("BE" & Target.row).Value
                Me.Range(COLLETJVZudChg & Target.row).Value = ""
                Me.Range(COLLETJVZOJTDCost & Target.row).Value = "T"
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
          
        End If
        
    End If
    
    
    Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
    Application.EnableEvents = True

Else
    MsgBox "Period Is Not Ready For Ops Input!", vbInformation
End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
'On Error GoTo errExit
    Me.Unprotect "password"
    Application.EnableEvents = False



If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then

    
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


    
GoTo 999
errexit:
MsgBox ("There was an automation error")

999:
If Sheet2.Range("ProtectSheet").Value = "True" Then
    Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If
Application.EnableEvents = True
Application.ScreenUpdating = True




End Sub

