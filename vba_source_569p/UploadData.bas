



Public Sub SendToVista()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.ScreenUpdating = False


If IsEmpty(COLJobNo) = True Or COLJobNo = 0 Then
    SetColumns
End If


If Sheet2.Range("LastClosedMth").Value <= Sheet7.Range("Month").Value And Sheet2.Range("CheckClosedMonth").Value = "Y" Then
    MsgBox "Cannot Update Closed Month", vbInformation
    GoTo 99
End If


ProtectUnProtect ("Off")

'Application.ScreenUpdating = False
Sheet7.Activate

8888:


SendToDatabase.Show vbModeless
DoEvents
       
Call CopyWIPDetail

GoTo 9999
errexit:
MsgBox ("There was an Error in the SendToVista Routine")
9999:

ProtectUnProtect ("On")

Unload SendToDatabase

Application.ScreenUpdating = True
Application.EnableEvents = True

99:

End Sub



Public Sub CopyWIPDetail()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If IsEmpty(COLJobNo) = True Or COLJobNo = 0 Then
    SetColumns
End If


Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim MyRow As Range
ProtectUnProtect ("Off")
Sheet8.Range("A4:Z2000").ClearContents


Dim prg As Integer
prg = 4

For Each MyRow In Sheet7.Range("SummaryData").Rows
        
    With MyRow
        
        If MyRow.Cells(1, COLZudChg).Value = "T" Or MyRow.Cells(1, COLZJCOR).Value = "T" Or MyRow.Cells(1, COLZJCOP).Value = "T" Then
        
            If Len(MyRow.Cells(1, COLJobNo).Value) > 0 Then
            
                ' WIPDetail
                Sheet8.Range("A" & CStr(prg)).Value = CStr(Sheet7.Range("Company").Value)                               ' Co
                Sheet8.Range("B" & CStr(prg)).Value = VBA.Left(CStr(MyRow.Cells(1, COLJobNo).Value), 2)                 ' Dept
                Sheet8.Range("C" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLJobNo).Value)                              ' Contract
                Sheet8.Range("D" & CStr(prg)).Value = CStr(Sheet7.Range("Month").Value)                                 ' Month
                If CStr(MyRow.Cells(1, COLZJCOR).Value) = "T" Then
                    Sheet8.Range("E" & CStr(prg)).Value = MyRow.Cells(1, COLZGAAPRevNew).Value                          ' GAAPRev
                    
                    If MyRow.Cells(1, COLZGAAPRevNew).Font.Bold = True Then
                        Sheet8.Range("H" & CStr(prg)).Value = "Y"
                    Else
                        Sheet8.Range("H" & CStr(prg)).Value = "N"
                    End If
                    
                Else
                    Sheet8.Range("E" & CStr(prg)).Value = MyRow.Cells(1, COLZGAAPRev).Value
                    
                    If MyRow.Cells(1, COLZGAAPRev).Font.Bold = True Then
                        Sheet8.Range("H" & CStr(prg)).Value = "Y"                                                       ' GAAP Rev Plugged
                    Else
                        Sheet8.Range("H" & CStr(prg)).Value = "N"
                    End If
                End If
                Sheet8.Range("F" & CStr(prg)).Value = 0                                                                 ' GAAP Other Rev
                Sheet8.Range("G" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLZGAAPRevNotes).Value)                      ' GAAP RevNotes
                If CStr(MyRow.Cells(1, COLZJCOP).Value) = "T" Then
                    Sheet8.Range("I" & CStr(prg)).Value = MyRow.Cells(1, COLZGAAPCostNew).Value                         ' GAAP Cost
                    Sheet8.Range("L" & CStr(prg)).Value = "Y"                                                           ' GAAP Cost Plugged
                Else
                    If MyRow.Cells(1, COLOvrCostProj).Value > 0 Then
                        Sheet8.Range("I" & CStr(prg)).Value = MyRow.Cells(1, COLOvrCostProj).Value
                    Else
                        Sheet8.Range("I" & CStr(prg)).Value = 0
                    End If
                    
                    If MyRow.Cells(1, COLZGAAPCost).Font.Bold = True Then
                        Sheet8.Range("L" & CStr(prg)).Value = "Y"
                    Else
                        Sheet8.Range("L" & CStr(prg)).Value = "N"
                    End If
                    
                End If
                
                Sheet8.Range("J" & CStr(prg)).Value = 0                                                                 ' GAAP Other Cost
                Sheet8.Range("K" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLZGAAPCostNotes).Value)                     ' GAAP Cost Notes
                
                If CStr(MyRow.Cells(1, COLCompDate).Value) <> "" Then
                    Sheet8.Range("M" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLCompDate).Value)                       ' Comp Date
                End If
                
                If CStr(MyRow.Cells(1, COLZOPsRChg).Value) = "T" And MyRow.Cells(1, COLZOPsRevNew).Value <> 0 Then      'Op's Rev
                    Sheet8.Range("N" & CStr(prg)).Value = MyRow.Cells(1, COLZOPsRevNew).Value
                Else
                    Sheet8.Range("N" & CStr(prg)).Value = MyRow.Cells(1, COLZOPsRev).Value
                End If
                
                If CStr(MyRow.Cells(1, COLZOPsCChg).Value) = "T" Or MyRow.Cells(1, COLZOPsCost).Font.Bold = True Then   ' Op's Cost
                    Sheet8.Range("O" & CStr(prg)).Value = MyRow.Cells(1, COLZOPSCostNew).Value
                Else
                    Sheet8.Range("O" & CStr(prg)).Value = MyRow.Cells(1, COLZOPsCost).Value
                End If
                
                Sheet8.Range("P" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLEstimator).Value)                          ' Estimator
                Sheet8.Range("Q" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLPrjMngr).Value)                            ' PM
                Sheet8.Range("R" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLZOPsRevNotes).Value)                       ' Op's Rev Notes
                Sheet8.Range("S" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLZOPsCostNotes).Value)                      ' Op's Cost Notes
                Sheet8.Range("T" & CStr(prg)).Value = CStr(MyRow.Cells(1, COLZUserName).Value)                          ' UserName
                
                If CStr(MyRow.Cells(1, COLDone).Value) = "P" Then                                                       ' Completed
                    Sheet8.Range("U" & CStr(prg)).Value = "Y"
                Else
                    Sheet8.Range("U" & CStr(prg)).Value = "N"
                End If
                
                If (MyRow.Cells(1, COLZOPsRev).Font.Bold = True Or MyRow.Cells(1, COLZOPsRevNew).Value <> 0) And _
                    MyRow.Cells(1, COLZOPsRevNew).Value <> MyRow.Cells(1, COLPMProjRev).Value Then                      ' Op's Rev Plug
                        Sheet8.Range("V" & CStr(prg)).Value = "Y"
                Else
                        Sheet8.Range("V" & CStr(prg)).Value = "N"
                End If
                
                
                
                
                If (MyRow.Cells(1, COLZOPsCost).Font.Bold = True Or MyRow.Cells(1, COLZOPSCostNew).Value <> 0) And _
                    MyRow.Cells(1, COLZOPSCostNew).Value <> MyRow.Cells(1, COLPMProjCost).Value Then                   ' Op's Cost Plug
                        Sheet8.Range("W" & CStr(prg)).Value = "Y"
                Else
                        Sheet8.Range("W" & CStr(prg)).Value = "N"
                End If
                
                Sheet8.Range("X" & CStr(prg)).Value = MyRow.Cells(1, COLJTDBonusProfit).Value                           ' BonusProfit
                Sheet8.Range("Y" & CStr(prg)).Value = MyRow.Cells(1, COLJTDBonusProfitNotes).Value                      ' BonusProfitNotes
                Sheet8.Range("Z" & CStr(prg)).Value = MyRow.Cells(1, COLZBatchSeq).Value                                ' BatchSeq
                
                HasUDData = True
                
                prg = prg + 1
            
            End If
            
        End If

123:

    End With

Next MyRow


Dim pauseTime As Double
pauseTime = 2

If HasUDData = True Then
    SendToDatabase.Label1 = "Sending WIP Details"
    DoEvents
    'Call UploadData("WIPDetail", "budWIPDetail", "LCGRCSUpdatebudWIPDetail")
    Application.Wait (Now + TimeValue("0:00:" & pauseTime))
End If


GoTo 9999
errexit:
MsgBox ("There was an Error in the CopyWIPDetail Routine")
9999:

Sheet2.Range("Sent").Value = True

ProtectUnProtect ("On")
Application.ScreenUpdating = True
End Sub


Function GUID$(Optional lowercase As Boolean, Optional parens As Boolean)
    Dim k&, h$
    GUID = Space(36)
    For k = 1 To Len(GUID)
        Randomize
        Select Case k
            Case 9, 14, 19, 24: h = "-"
            Case 15:            h = "4"
            Case 20:            h = Hex(Rnd * 3 + 8)
            Case Else:          h = Hex(Rnd * 15)
        End Select
        Mid$(GUID, k, 1) = h
    Next
    If lowercase Then GUID = LCase$(GUID)
    If parens Then GUID = "{" & GUID & "}"
    GUID = "##RCS" & Replace(GUID, "-", "")
End Function


Public Function CreateColumnList(TabName As String, Def As Boolean)
Application.ScreenUpdating = False

Dim myTab As Excel.Worksheet
Dim myRecord As Range
Dim myField As Range
Dim sOut As String

Dim i As Integer
i = 1
sOut = " ("


For Each myTab In ThisWorkbook.worksheets
    If myTab.Name = TabName Then
         For Each myRecord In myTab.Range("A1:A1")
                With myRecord
                    For Each myField In myTab.Range(.Cells(1), _
                        myTab.Cells(.row, myTab.Columns.count).End(xlToLeft))
                        If Def = True Then
                            sOut = sOut & "[" & Replace(myField.Value, "*", "") & "] " & Replace(myField.Offset(2, 0).Value, "*", "") & ", "
                        Else
                            sOut = sOut & "[" & Replace(myField.Value, "*", "") & "],"
                        End If
                            
                        i = i + 1
                    Next myField
                
                
                End With
            
        Next myRecord
    End If
Next
    
CreateColumnList = Left(sOut, Len(sOut) - 1) & ")"

Application.ScreenUpdating = True
End Function


Public Function CreateParamList(TabName As String)
Application.ScreenUpdating = False

Dim myTab As Excel.Worksheet
Dim myRecord As Range
Dim myField As Range
Dim sOut As String

Dim i As Integer
i = 1
sOut = " VALUES ("


For Each myTab In ThisWorkbook.worksheets
    If myTab.Name = TabName Then
         For Each myRecord In myTab.Range("A1:A1")
                With myRecord
                    For Each myField In myTab.Range(.Cells(1), _
                        myTab.Cells(.row, myTab.Columns.count).End(xlToLeft))
                        sOut = sOut & "?,"
                        i = i + 1
                    Next myField
                
                
                End With
            
        Next myRecord
    End If
Next
    

CreateParamList = Left(sOut, Len(sOut) - 1) & ")"


Application.ScreenUpdating = True
End Function

Public Function CountCols(TabName As String) As Integer
Application.ScreenUpdating = False

Dim myTab As Excel.Worksheet
Dim myRecord As Range
Dim myField As Range

Dim i As Integer
i = 1

For Each myTab In ThisWorkbook.worksheets
    If myTab.Name = TabName Then
         For Each myRecord In myTab.Range("A1:A1")
                With myRecord
                    For Each myField In myTab.Range(.Cells(1), _
                        myTab.Cells(.row, myTab.Columns.count).End(xlToLeft))
                        
                        i = i + 1
                    Next myField
                
                End With
            
        Next myRecord
    End If
Next

CountCols = i - 1

Application.ScreenUpdating = True
End Function



