

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If Sheet2.Range("Role").Value = "WipViewOnly" Then
    MsgBox "You do not have permission to edit.", vbOKOnly, "View Olny Permission"
    GoTo 9999:
End If

Cancel = True
Me.Unprotect "password"
Application.ScreenUpdating = False
Application.EnableEvents = False

If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If

' JTD Show Hide
If FindIntersect(Range(Target.Address), Me.Range("JTDHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("JTDPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJTDChgProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJTDChgProfit")).Hidden = False
        Me.Columns(NumDict(Me.CodeName)("COLJTDBonusProfitNotes")).Hidden = True
        Me.Range("JTDHeading").Select
        Me.Range("JTDPlus").Value = "-"
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLJTDPctComp") & ":" & LetDict(Me.CodeName)("COLJTDChgProfit")).Hidden = True
        Me.Range("JTDPlus").Select
        Me.Range("JTDPlus").Value = "+"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If


' APY Show Hide
If FindIntersect(Range(Target.Address), Me.Range("APYHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("APYPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLAPYPctComp") & ":" & LetDict(Me.CodeName)("COLAPYBonusProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLAPYPctComp") & ":" & LetDict(Me.CodeName)("COLAPYBonusProfit")).Hidden = False
        Me.Range("APYHeading").Select
        Me.Range("APYPlus").Value = "-"
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLAPYPctComp") & ":" & LetDict(Me.CodeName)("COLAPYBonusProfit")).Hidden = True
        Me.Range("APYPlus").Select
        Me.Range("APYPlus").Value = "+"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If


' CY Show Hide
If FindIntersect(Range(Target.Address), Me.Range("CYHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("CYPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLCYRev") & ":" & LetDict(Me.CodeName)("COLCYBonusProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLCYRev") & ":" & LetDict(Me.CodeName)("COLCYBonusProfit")).Hidden = False
        Me.Range("CYHeading").Select
        Me.Range("CYPlus").Value = "-"
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLCYRev") & ":" & LetDict(Me.CodeName)("COLCYBonusProfit")).Hidden = True
        Me.Range("CYPlus").Select
        Me.Range("CYPlus").Value = "+"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If

' B Show Hide
If FindIntersect(Range(Target.Address), Me.Range("BHeading")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("BPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Columns(LetDict(Me.CodeName)("COLBRev") & ":" & LetDict(Me.CodeName)("COLBProfit")).Hidden = True Then
    
        Me.Columns(LetDict(Me.CodeName)("COLBRev") & ":" & LetDict(Me.CodeName)("COLBProfit")).Hidden = False
        Me.Range("BHeading").Select
        Me.Range("BPlus").Value = "-"
        
    Else
    
        Me.Columns(LetDict(Me.CodeName)("COLBRev") & ":" & LetDict(Me.CodeName)("COLBProfit")).Hidden = True
        Me.Range("BPlus").Select
        Me.Range("BPlus").Value = "+"
    
    End If
    
    Application.EnableEvents = True

GoTo 99
End If



' ALL Show Hide
If FindIntersect(Range(Target.Address), Me.Range("ALLPlus")) = "Yes" Then
    Application.EnableEvents = False
    
    If Me.Range("ALLPlus").Value = "+" Then
        Me.Columns(LetDict(Me.CodeName)("COLJTDPriorProfit") & ":" & LetDict(Me.CodeName)("COLBILLBillExRev")).Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLOvrRevProjNotes")).Hidden = True
        Me.Columns(LetDict(Me.CodeName)("COLOvrCostProjNotes")).Hidden = True
        Me.Columns(LetDict(Me.CodeName)("COLJTDBonusProfitNotes")).Hidden = True
        Me.Range("ALLPlus").Value = "-"
    Else
    
        Columns(LetDict(Me.CodeName)("COLJTDPriorProfit") & ":" & LetDict(Me.CodeName)("COLJTDChgProfit")).EntireColumn.Hidden = True
        Columns(LetDict(Me.CodeName)("COLAPYPctComp") & ":" & LetDict(Me.CodeName)("COLAPYBonusProfit")).EntireColumn.Hidden = True
        Columns(LetDict(Me.CodeName)("COLCYRev") & ":" & LetDict(Me.CodeName)("COLCYBonusProfit")).EntireColumn.Hidden = True
        Columns(LetDict(Me.CodeName)("COLBRev") & ":" & LetDict(Me.CodeName)("COLBProfit")).EntireColumn.Hidden = True
        Columns(LetDict(Me.CodeName)("COLBILLBillings") & ":" & LetDict(Me.CodeName)("COLBILLBillExRev")).EntireColumn.Hidden = True
        Me.Range("ALLPlus").Value = "+"
    
    
    End If
        Me.Range("JTDHeading").Select
    Application.EnableEvents = True

GoTo 99
End If

' Bonus Profit Notes
If FindIntersect(Range(Target.Address), Me.Range("BonusProfitNotesShow")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("BonusProfitNotes")) = "Yes" Then
    Application.EnableEvents = False
    If Me.Columns(LetDict(Me.CodeName)("COLJTDBonusProfitNotes")).Hidden = False Then
    
        Me.Columns(LetDict(Me.CodeName)("COLJTDBonusProfitNotes")).EntireColumn.Hidden = True
        Me.Range("BonusProfitNotes").WrapText = False
        Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotesShow")).Select
        If Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotes")).Value <> "" Then
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotesShow).Value")) = "+"
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotesShow")).Font.Bold = True
        End If
        
    Else
        Me.Columns(LetDict(Me.CodeName)("COLJTDBonusProfitNotes")).EntireColumn.Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLJTDBonusProfitNotes")).ColumnWidth = 50
        Me.Cells(Target.row, LetDict(Me.CodeName)("COLJTDBonusProfitNotes")).Select
        If Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotes")).Value <> "" Then
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotesShow")).Value = "+"
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotesShow")).Font.Bold = True
        Else
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotesShow")).Value = ""
        End If
        
        With Me.Cells(Target.row, NumDict(Me.CodeName)("COLJTDBonusProfitNotes"))
            .VerticalAlignment = xlTop
            .WrapText = True
        End With
        
    End If
    Application.EnableEvents = True

GoTo 99
End If


' Revenue Notes
If FindIntersect(Range(Target.Address), Me.Range("NotesRevShow")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("NotesRev")) = "Yes" Then
    Application.EnableEvents = False
    If Me.Columns(LetDict(Me.CodeName)("COLOvrRevProjNotes")).Hidden = False Then
    
        Me.Columns(LetDict(Me.CodeName)("COLOvrRevProjNotes")).EntireColumn.Hidden = True
        Me.Range("NotesRev").WrapText = False
        Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes2")).Select
        If Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes")).Value <> "" Then
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes2")).Value = "+"
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes2")).Font.Bold = True
        End If
        
    Else
        Me.Columns(LetDict(Me.CodeName)("COLOvrRevProjNotes")).EntireColumn.Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLOvrRevProjNotes")).ColumnWidth = 50
        Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes")).Select
        If Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes")).Value <> "" Then
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes2")).Value = "+"
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes2")).Font.Bold = True
        Else
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes2")).Value = ""
        End If
        
        With Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrRevProjNotes"))
            .VerticalAlignment = xlTop
            .WrapText = True
        End With
        
    End If
    Application.EnableEvents = True

GoTo 99
End If


' Cost Notes
If FindIntersect(Range(Target.Address), Me.Range("NotesCostShow")) = "Yes" Or FindIntersect(Range(Target.Address), Me.Range("NotesCost")) = "Yes" Then
    Application.EnableEvents = False
    If Me.Columns(LetDict(Me.CodeName)("COLOvrCostProjNotes")).Hidden = False Then
    
        Me.Columns(LetDict(Me.CodeName)("COLOvrCostProjNotes")).EntireColumn.Hidden = True
        Me.Range("NotesCost").WrapText = False
        Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes2")).Select
        If Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes")).Value <> "" Then
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes2")).Value = "+"
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes2")).Font.Bold = True
        End If
        
    Else
        Me.Columns(LetDict(Me.CodeName)("COLOvrCostProjNotes")).EntireColumn.Hidden = False
        Me.Columns(LetDict(Me.CodeName)("COLOvrCostProjNotes")).ColumnWidth = 50
        Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes")).Select
        If Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes")).Value <> "" Then
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes2")).Value = "+"
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes2")).Font.Bold = True
        Else
            Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes2")).Value = ""
        End If
        With Me.Cells(Target.row, NumDict(Me.CodeName)("COLOvrCostProjNotes"))
            .VerticalAlignment = xlTop
            .WrapText = True
        End With
        
    End If
    Application.EnableEvents = True

GoTo 99
End If

' Closed Month
If Sheet2.Range("LastClosedMth").Value <= Sheet17.Range("StartMonth").Value And Sheet2.Range("CheckClosedMonth").Value = "Y" _
    Or Sheet2.Range("LastClosedMth").Value = "" Then
    MsgBox "Cannot Update Closed Month", vbInformation
    GoTo 99
End If



' Not Ready For Ops
If Sheet2.Range("ReadyForOpsAppr1").Value <> "Y" Then
    MsgBox "Period Is Not Ready For Ops Input!", vbInformation
    GoTo 99
End If
    

' Done Heading
If FindIntersect(Range(Target.Address), Me.Range("DoneHeading")) = "Yes" Then
    If Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff Then
        Application.EnableEvents = False
        Application.ScreenUpdating = Fales
        Application.Calculation = xlCalculationManual
        
        If Sheet2.Range("InitAppr").Value <> "Y" And Sheet2.Range("FinalAppr").Value <> "Y" Then
        
            If MsgBox("    This will toggle all Contracts Done Or Not " & vbCrLf & "       Done that are shown on this page." & vbCrLf & vbCrLf & _
            "                           Continue?", vbYesNo) = vbYes Then
                If Sheet2.Range("CompleteAll").Value = True Then
                    Call UpdateAllRows(Me, "N")
                    Sheet2.Range("CompleteAll").Value = False
                Else
                
                    Call UpdateAllRows(Me, "Y")
                    Sheet2.Range("CompleteAll").Value = True
                End If
                
                
                
            
            End If
            
        Else
        
            MsgBox "You Need to Un-Approve Before you can take this action", vbInformation
        
        End If
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    Else
    
        MsgBox "Cannot Update - Has been Final Approved For Ops", vbInformation
    
    End If
Sheet11.Activate
GoTo 99
End If
    
    
    
    
' Done
If FindIntersect(Range(Target.Address), Me.Range("Done")) = "Yes" Then
   If Sheet17.Shapes("OFA-Yes").ControlFormat.Value = xlOff Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        If Target.Value = "P" Then
            Target.Value = ""
            'Sheet12.Range(LetDict("Sheet12")("COLDone") & Target.row).Value = ""
            Target.ClearComments
        Else
            Target.Value = "P"
            'Sheet12.Range(LetDict("Sheet12")("COLDone") & Target.row).Value = "P"
        End If
        
        Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
        Me.Range(LetDict(Me.CodeName)("COLZUserName") & Target.row).Value = Sheet2.Range("UserName2").Value
        
        UpdateRow Target.row, Me
        
        
        
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
    
' Close
If FindIntersect(Range(Target.Address), Me.Range("Close")) = "Yes" Then
    ' C9: Block when OpsApproved or AcctApproved
    If Sheet2.Range("FinalAppr").Value = "Y" Then
        MsgBox "Cannot Update - Has been Final Approved For Ops", vbInformation
        GoTo 99
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = Fales
    Application.Calculation = xlCalculationManual
    
    If Target.Value = "C" Then
        Target.Value = ""
        Target.ClearComments
    Else
        Target.Value = "C"
        With Target
            .AddComment.Text Text:="Closed By = " & CStr(Sheet2.Range("UserName2").Value)
            .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
            .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
            .Comment.Shape.TextFrame.Characters.Font.Size = 10
            .Comment.Shape.Height = 25
            .Comment.Shape.Width = 125
        End With
    End If
    
    Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
    Me.Range(LetDict(Me.CodeName)("COLZUserName") & Target.row).Value = Sheet2.Range("UserName2").Value
    
    UpdateRow Target.row, Me
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
GoTo 99
End If



' Revenue
If FindIntersect(Range(Target.Address), Me.Range("ExecOvrRevAdj")) = "Yes" Then
    ' C9: Block when OpsApproved or AcctApproved
    If Sheet2.Range("FinalAppr").Value = "Y" Then
        MsgBox "Cannot Update - Has been Final Approved For Ops", vbInformation
        GoTo 99
    End If

    ' GAAP View On
    If Me.Shapes("Option Button 23133").ControlFormat.Value = xlOn And Sheet2.Range("Role").Value = "WIPAccounting" Then
        
        If Left(Target.Formula, 1) = "=" Or Target.Value = Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNew") & Target.row).Value Then
            
            If Target.Value = Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNew") & Target.row).Value Or Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNew") & Target.row).Value = "" Then
                Target.Value = Me.Range(LetDict(Me.CodeName)("COLZGAAPRev") & Target.row).Value
                
                Me.Range(LetDict(Me.CodeName)("COLZJCOR") & Target.row).Value = ""
                If Me.Range(LetDict(Me.CodeName)("COLZGAAPRev") & Target.row).Font.Bold = True Then
                    ' Orange
                    With Target.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                Else
                
                    If Target.Value = Target.Offset(0, -2).Value Then
                        ' Yellow
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    
                    Else
                        
                        If Me.Range(LetDict(Me.CodeName)("COLZJCOR") & Target.row).Value = "T" Then
                            ' Orange
                            With Target.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 49407
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        Else
                            ' Light Red
                            With Target.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorAccent2
                                .TintAndShade = 0.399975585192419
                                .PatternTintAndShade = 0
                            End With
                        
                        
                        End If
                    
                    End If
                
                End If
            
            Else
                Target.Value = Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNew") & Target.row).Value
            
                ' Orange
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            
            End If
        
        Else
            Target.Formula = "=" & LetDict(Me.CodeName)("COLPMProjRev") & Target.row
            Me.Range(LetDict(Me.CodeName)("COLZJCOR") & Target.row).Value = "T"
            
            Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNew") & Target.row).Value = Me.Range(LetDict(Me.CodeName)("COLOvrRevProj") & Target.row).Value
            Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNew") & Target.row).Font.Bold = False
            
            ' Yellow
            With Target.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        
        End If
    
    Else
    ' OPs View
        If Left(Target.Formula, 1) = "=" Or Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Value Then
            
            If Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Value Or Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Value = "" Then
                Target.Value = Me.Range(COLZOPsRev & Target.row).Value
                Me.Range(LetDict(Me.CodeName)("COLZOPsRChg") & Target.row).Value = ""
                
                If Me.Range(LetDict(Me.CodeName)("COLZOPsRev") & Target.row).Font.Bold = True Then
                    ' Orange
                    With Target.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                Else
                
                    If Target.Value = Target.Offset(0, -2).Value Then
                        ' Yellow
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    
                    Else
                        
                        If Me.Range(LetDict(Me.CodeName)("COLZOPsRChg") & Target.row).Value = "T" Then
                            ' Orange
                            With Target.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 49407
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        Else
                            ' Light Red
                            With Target.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorAccent2
                                .TintAndShade = 0.399975585192419
                                .PatternTintAndShade = 0
                            End With
                        
                        
                        End If
                    
                    End If
                
                End If
            
            Else
                Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Value
            
                ' Orange
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            
            End If
        
        Else
            Target.Formula = "=" & LetDict(Me.CodeName)("COLPMProjRev") & Target.row
            Me.Range(LetDict(Me.CodeName)("COLZOPsRChg") & Target.row).Value = "T"
            Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
            Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Value = Me.Range(COLOvrRevProj & Target.row).Value
            Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Font.Bold = False
            ' Yellow
            With Target.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        
        End If
    
    End If
  
End If


' Exec Overide Cost
If FindIntersect(Range(Target.Address), Me.Range("ExecOvrCostAdj")) = "Yes" Then
    ' C9: Block when OpsApproved or AcctApproved
    If Sheet2.Range("FinalAppr").Value = "Y" Then
        MsgBox "Cannot Update - Has been Final Approved For Ops", vbInformation
        GoTo 99
    End If

    ' OPs View
    If Left(Target.Formula, 1) = "=" Or Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Value Then
        
        ' If
        If Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Value Or _
            Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Value = "" Then
            
            Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsCost") & Target.row).Value
            
            If Me.Range(LetDict(Me.CodeName)("COLZOPsCost") & Target.row).Font.Bold = True Then
                ' Orange
                With Target.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                
            Else
            
                If Target.Value = Target.Offset(0, -1).Value Then
                    ' Yellow
                    With Target.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                
                Else
                    
                    If Me.Range(LetDict(Me.CodeName)("COLZOPsCChg") & Target.row).Value = "T" Then
                        ' Orange
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 49407
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    Else
                        ' Light Red
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent2
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                    
                    
                    End If
                
                End If
            
            End If
        
        Else
            Target.Value = Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Value
                        
            ' Orange
            With Target.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 49407
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        End If
    
    Else
        Target.Formula = "=" & LetDict(Me.CodeName)("COLPMProjCost") & Target.row
        Me.Range(LetDict(Me.CodeName)("COLZOPsCChg") & Target.row).Value = "T"
        Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
        Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Value = Me.Range(LetDict(Me.CodeName)("COLOvrCostProj") & Target.row).Value
        Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Font.Bold = False
        
        ' Yellow
        With Target.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    
    End If
    
  
End If

99:
Application.EnableEvents = True
Application.ScreenUpdating = True
If Sheet2.Range("ProtectSheet").Value = True Then
    Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If


GoTo 9999
errexit:
MsgBox "There was an error in the Worksheet_BeforeDoubleClick Routine. " & Err, vbOKOnly
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


If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If


Dim CurrentColumnCount As Long

' Get the current number of columns
CurrentColumnCount = Me.Columns.count

' Check if the number of columns has increased
If CurrentColumnCount > LastColumnCount Or IsEmpty(LastColumnCount) = True Then
    LastColumnCount = CurrentColumnCount

End If

' C9: Block direct typing when OpsApproved or AcctApproved
If Sheet2.Range("FinalAppr").Value = "Y" Then
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    MsgBox "Cannot Update - Batch has been Ops Approved. Un-approve first.", vbInformation
    GoTo 99
End If

' Marked Ops Done
If Me.Cells(Target.row, NumDict(Me.CodeName)("COLDone")).Value = "P" And Sheet2.Range("GAAPView").Value = "N" Then
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    MsgBox "Uncheck the Done Checkbox to make changes", vbInformation
    GoTo 99
End If





' Done, GAAP Done, Close
If FindIntersect(Range(Target.Address), Me.Range("Done")) = "Yes" _
    Or FindIntersect(Range(Target.Address), Me.Range("Close")) = "Yes" Then
    
    Application.EnableEvents = False

    On Error Resume Next
    Application.Undo
    On Error GoTo 0

    MsgBox "Double Click to Change Value", vbInformation
    
    Application.EnableEvents = True

    GoTo 99
End If



If Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then

    Dim row As Integer
    
    ' Final Approval Yes Button
    If Sheet17.Shapes("AFA-Yes").ControlFormat.Value = xlOff Then
    
        If FindIntersect(Target, Me.Range("SummaryDataInput")) = "Yes" Then
        
            If Me.Range(LetDict(Me.CodeName)("COLJobNumber") & Target.row).Value <> "" Then
                
                If Me.Range(LetDict(Me.CodeName)("COLZContractStatus") & Target.row).Value = 1 Or Sheet2.Range("EditClosedJobs") = True Then
                    
                    If Sheet2.Range("ProtectSheet").Value = "True" Then
                        Me.Unprotect "password"
                    End If
                    
                    ' Revenue
                    If Target.Column = NumDict(Me.CodeName)("COLOvrRevProj") Then
                        
                       
                        If JTDProfitCheck(Me, CInt(NumDict(Me.CodeName)("COLOvrRevProj")), Target.Value, Target.row) = False Then
                            MsgBox "Your Entry has caused JTD Profit to Exceed Profit at Completion", vbInformation
                            On Error Resume Next
                            Application.Undo
                            On Error GoTo 0
                            GoTo 99
                        End If
                    
                    
     
                        Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
                        Me.Range(LetDict(Me.CodeName)("COLZOPsRChg") & Target.row).Value = "T"
                        Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Value = Target.Value
                        Me.Range(LetDict(Me.CodeName)("COLZOPsRevNew") & Target.row).Font.Bold = True
    
                        ' Orange
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 49407
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        GoTo FormatDoneField
                        
                    End If
                    
                    ' Cost
                    If Target.Column = NumDict(Me.CodeName)("COLOvrCostProj") Then
                    
                        If JTDProfitCheck(Me, CInt(NumDict(Me.CodeName)("COLOvrCostProj")), Target.Value, Target.row) = False Then
                            MsgBox "Your Entry has caused JTD Profit to Exceed Profit at Completion", vbInformation
                            On Error Resume Next
                            Application.Undo
                            On Error GoTo 0
                            GoTo 99
                        End If
    
                        If Target.Value < Me.Range(LetDict(Me.CodeName)("COLJTDCost") & Target.row).Value Then
                            MsgBox "Your Entry has Caused % Complete over 100%", vbInformation
                            On Error Resume Next
                            Application.Undo
                            On Error GoTo 0
                            GoTo 99
                        End If
    
                        
                        
                        
                        Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
                        Me.Range(LetDict(Me.CodeName)("COLZOPsCChg") & Target.row).Value = "T"
                        Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Value = Target.Value
                        Me.Range(LetDict(Me.CodeName)("COLZOPsCostNew") & Target.row).Font.Bold = True
                            
                      
                        ' Orange
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 49407
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        GoTo FormatDoneField
                        
                    End If
                    
                    ' Revenue Notes
                    If Target.Column = NumDict(Me.CodeName)("COLOvrRevProjNotes") Then
                        ' GAAP View On
                        If Me.Shapes("Option Button 23133").ControlFormat.Value = xlOn Then
                            Me.Range(LetDict(Me.CodeName)("COLZJCOR") & Target.row).Value = "T"
                            Me.Range(LetDict(Me.CodeName)("COLZGAAPRevNotes") & Target.row).Value = Target.Value
                        Else
                            Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
                            Me.Range(LetDict(Me.CodeName)("COLZOPsRevNotes") & Target.row).Value = Target.Value
                        End If
                        
                        Target.Offset(0, 1).Value = "+"
                        Target.Offset(0, 1).Font.Bold = True
                        Target.WrapText = True
                        GoTo FormatDoneField
                        
                    End If
                    
                    
                    ' Cost Notes
                    If Target.Column = NumDict(Me.CodeName)("COLOvrCostProjNotes") Then
                        
                        Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
                        Me.Range(LetDict(Me.CodeName)("COLZOPsCostNotes") & Target.row).Value = Target.Value
                        
                        Target.Offset(0, 1).Value = "+"
                        Target.Offset(0, 1).Font.Bold = True
                        Target.WrapText = True
                        GoTo FormatDoneField
                        
                    End If
                    
                    ' Bonus Profit
                    If Target.Column = NumDict(Me.CodeName)("COLJTDBonusProfit") Then
                                            
                        If Target.Value > Me.Range(LetDict(Me.CodeName)("COLProjProfit") & Target.row).Value Then
                            MsgBox "Profit Not Allowed to be More Than Profit @ Completion", vbInformation
                            On Error Resume Next
                            Application.Undo
                            On Error GoTo 0
                            GoTo 99
                        End If
                        
                        Me.Range(LetDict(Me.CodeName)("COLZOPsBonusNew") & Target.row).Value = Target.Value
                        Me.Range(LetDict(Me.CodeName)("COLZOPsBonusNew") & Target.row).Font.Bold = True
                        ' Orange
                        With Target.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 49407
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        
                        GoTo FormatDoneField
                    End If
                    
                    
                    ' Bonus Profit Notes
                    If Target.Column = NumDict(Me.CodeName)("COLJTDBonusProfitNotes") Then
                                            
                        Target.Offset(0, 1).Value = "+"
                        Target.Offset(0, 1).Font.Bold = True
                        Target.WrapText = True
                        Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
                        
                    End If
                    
                    
                    ' Estimator, PM, Completion Date, COLJTDBonusProfit, COLJTDBonusProfitNotes
                    If Target.Column = NumDict(Me.CodeName)("COLEstimator") Or Target.Column = NumDict(Me.CodeName)("COLPrjMngr") Or Target.Column = NumDict(Me.CodeName)("COLPrjExec") Or Target.Column = NumDict(Me.CodeName)("COLCompDate") Or _
                        Target.Column = NumDict(Me.CodeName)("COLNotes") Or Target.Column = NumDict(Me.CodeName)("COLJTDBonusProfit") Or Target.Column = NumDict(Me.CodeName)("COLJTDBonusProfitNotes") Then
                            Me.Range(LetDict(Me.CodeName)("COLZudChg") & Target.row).Value = "T"
                            ' Add Border to Done field to indicate data needs to be sent back to Vista
                            GoTo FormatDoneField
                        
                    End If
                        
                
                Else

                    On Error Resume Next
                    Application.Undo
                    On Error GoTo 0

                    MsgBox ("You cannot edit closed Jobs")
                    
                End If
                
            Else

                On Error Resume Next
                Application.Undo
                On Error GoTo 0
                MsgBox ("You cannot edit Non-Job Rows")
            
            End If
            
            Me.Range(LetDict(Me.CodeName)("COLZUserName") & Target.row).Value = Sheet2.Range("UserName2").Value
            
        End If
        
        row = Target.row
        
        
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

GoTo 9998
FormatDoneField:
' Add Border to Done field to indicate data needs to be sent back to Vista
Dim Col As String

Col = LetDict(Me.CodeName)("COLDone")

With Me.Range(Col & Target.row).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Me.Range(Col & Target.row).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Me.Range(Col & Target.row).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Me.Range(Col & Target.row).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With

9998:

    
GoTo 10000
errexit:
MsgBox "There was an error in the Worksheet_Change Routine. " & Err, vbOKOnly
10000:

Application.ScreenUpdating = True

99:
Application.EnableEvents = True
Application.ScreenUpdating = True

9999:

If Sheet2.Range("ProtectSheet").Value = "True" Then
    Me.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
End If

End Sub


