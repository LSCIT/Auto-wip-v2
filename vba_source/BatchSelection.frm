


Option Explicit
'Private words As Variant
Private MyRange As Variant
Private ActRange As Range

'   if a key is pressed while the form is open set focus on the textbox
Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo errexit

TextBox1.SetFocus

GoTo 9999
errexit:
MsgBox ("There was an Error in the ListBox1_KeyPress Routine")
9999:
End Sub

'   close form
Private Sub CloseForm_Click()
On Error GoTo errexit
BatchSelected = False
BatchSelectionCanceled = True
Unload Me

GoTo 9999
errexit:
MsgBox ("There was an Error in the CloseForm_Click Routine")
9999:

End Sub


Private Sub UserForm_Activate()
Set PRange = ActiveCell

End Sub


Private Sub UserForm_Initialize()
On Error GoTo errexit

Dim intTotalElements As Integer

On Error GoTo 50

'Check if Form 1 list is already in the array
intTotalElements = UBound(Form3Datalist) - LBound(Form3Datalist) + 1
50:

'If not poll the database and fill the array
'If intTotalElements = 0 Then

'Call GetForm3Data(Form3Datalist())

'End If

With Me.ListBox1
    .Clear
    .List = Form3Datalist
End With

GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_Initialize Routine")
9999:
End Sub


Private Sub TextBox1_Change()
On Error GoTo errexit
Me.ListBox1.List = Filter(Form3Datalist, VBA.UCase(TextBox1.Value), True, vbTextCompare)
Me.Repaint

GoTo 9999
errexit:
MsgBox ("There was an Error in the TextBox1_Change Routine")
9999:

End Sub



Private Sub PhaseOK_Click()
On Error GoTo errexit
Call SetDataVal
Unload BatchSelection
GoTo 9999
errexit:
MsgBox ("There was an Error in the PhaseOK_Click Routine")
9999:

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'On Error GoTo errExit

Call SetDataVal
Unload Me
GoTo 9999
errexit:
MsgBox ("There was an Error in the ListBox1_DblClick Routine")
9999:

End Sub

Private Sub SetDataVal()
'On Error GoTo errExit

Application.EnableEvents = False
Application.ScreenUpdating = False

ProtectUnProtect ("Off")

Dim i As Integer

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then
        Sheet7.Range("BatchId").Value = VBA.Left(ListBox1.Column(0, i), 3)
        Sheet7.Range("DEPT").Value = LTrim(RTrim(VBA.Mid(ListBox1.Column(0, i), 4, 30)))
    End If
Next

GoTo 9999
errexit:
MsgBox ("There was an Error in the SetDataVal Routine")
9999:

'Call GetDept(Sheet2.Range("UserName2").Value, Sheet7.Range("Company").Value)


Sheet7.Range("Month").Select
ProtectUnProtect ("On")
BatchSelected = True

Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub



Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo errexit

TextBox1.SetFocus

GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_KeyPress Routine")
9999:

End Sub



'clear selections on list box prior to running the search function
Private Sub btnClearAll_Click()
On Error GoTo errexit

Me.ListBox1.MultiSelect = fmMultiSelectSingle
Me.ListBox1.ListIndex = -1
Me.ListBox1.MultiSelect = fmMultiSelectExtended

GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_KeyPress Routine")
9999:

 End Sub

'search function on list box (note .column (0,i) and .column(1,i) fields indicate that both columns A and B are being searched)
Private Sub CommandButton1_Click()
On Error GoTo errexit

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



