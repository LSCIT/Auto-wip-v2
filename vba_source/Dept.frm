
Option Explicit
Private MyRange As Variant
Private ActRange As Range

'   if a key is pressed while the form is open set focus on the textbox
Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

TextBox1.SetFocus

GoTo 9999
errexit:
MsgBox ("There was an Error in the ListBox1_KeyPress Routine")
9999:
End Sub

'   close form
Private Sub CloseForm_Click()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Erase DeptDataList()
DeptSelectionCanceled = True
Unload Me

GoTo 9999
errexit:
MsgBox ("There was an Error in the CloseForm_Click Routine")
9999:

End Sub


Private Sub UserForm_Activate()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Set PRange = ActiveCell

GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_Activate Routine")
9999:

End Sub


Private Sub UserForm_Initialize()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If


If Sheet17.Range("StartCompany").Value <> "" Then
    
    Call GetDeptData(DeptDataList())
    
    If DoesArrayHaveItems(DeptDataList) = True Then
    
        With Me.ListBox1
            .Clear
            .List = DeptDataList
        End With
    Else
        Unload Me
    
    End If

Else

    MsgBox ("Please select a Company!")
    
End If


GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_Initialize Routine")
9999:
End Sub

Function IsArrayAllocated(Arr As Variant) As Boolean
        
        
        IsArrayAllocated = IsArray(Arr) And _
                           Not IsError(LBound(Arr, 1)) And _
                           LBound(Arr, 1) <= UBound(Arr, 1)
End Function
Function DoesArrayHaveItems(Arr As Variant) As Boolean
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

DoesArrayHaveItems = LBound(Arr, 1) <= UBound(Arr, 1)
GoTo 101
100:
DoesArrayHaveItems = False

GoTo 9999
errexit:
MsgBox ("There was an Error in the DoesArrayHaveItems Routine")
9999:


101:
End Function


Private Sub TextBox1_Change()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Me.ListBox1.List = Filter(DeptDataList, VBA.UCase(TextBox1.Value), True, vbTextCompare)
Me.Repaint

GoTo 9999
errexit:
MsgBox ("There was an Error in the TextBox1_Change Routine")
9999:
End Sub



Private Sub PhaseOK_Click()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Call SetDataVal

GoTo 9999
errexit:
MsgBox ("There was an Error in the PhaseOK_Click Routine")
9999:

Erase DeptDataList()
Unload Me

End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Call PhaseOK_Click

GoTo 9999
errexit:
MsgBox ("There was an Error in the ListBox1_DblClick Routine")
9999:

End Sub


Private Sub SetDataVal()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
Application.ScreenUpdating = False

ProtectUnProtect ("Off")

Dim i As Integer
Dim DeptNo As String
DeptNo = ""
Dim DeptName As String
DeptName = ""
Dim Seperator As String
Seperator = ""
Dim si As Integer
si = 0

On Error GoTo errexit
For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then
        If si > 0 Then
        
            Seperator = ","
        
        End If
    
        DeptNo = DeptNo & Seperator & RTrim(VBA.Left(ListBox1.Column(0, i), 10))
        DeptName = DeptName & Seperator & " " & VBA.Trim(VBA.Mid(ListBox1.Column(0, i), 11, 30))
        si = si + 1
    End If
Next

Sheet17.Range("StartDept").Value = DeptNo
Sheet17.Range("StartDeptName").Value = DeptName

GoTo 9999
errexit:
MsgBox ("There was an Error in the SetDataVal Routine")
9999:
Sheet17.Range("StartMonth").Select
ProtectUnProtect ("On")

Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub



Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

TextBox1.SetFocus

GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_KeyPress Routine")
9999:
Application.EnableEvents = True

End Sub



'clear selections on list box prior to running the search function
Private Sub btnClearAll_Click()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
 
 Me.ListBox1.MultiSelect = fmMultiSelectSingle
 Me.ListBox1.ListIndex = -1
 Me.ListBox1.MultiSelect = fmMultiSelectExtended

GoTo 9999
errexit:
MsgBox ("There was an Error in the btnClearAll_Click Routine")
9999:
End Sub

'search function on list box (note .column (0,i) and .column(1,i) fields indicate that both columns A and B are being searched)
Private Sub CommandButton1_Click()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If


Dim strSearch As String
Dim i As Long
Dim iContinueSearch As Integer
Dim Filecount As Variant

btnClearAll_Click
TextBox1.Value = ""

'strSearch = VBA.UCase("*" & TextBox2.Value & "*")

With Me.ListBox1

If .ListCount > 0 Then
   For i = 0 To .ListCount - 2
       If .Column(0, i) Like strSearch Then
           .Selected(i) = True
           .ListIndex = i + (.Height / 15)
           iContinueSearch = MsgBox("Find Next?", vbQuestion + vbYesNo)
           If iContinueSearch = vbYes Then
              .Selected(i) = False
           End If
           If iContinueSearch = vbNo Then
               Exit For
           End If
       End If
   Next i
   
End If
End With
 
GoTo 9999
errexit:
MsgBox ("There was an Error in the CommandButton1_Click Routine")
9999:

End Sub



