

Option Explicit
'Private words As Variant
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
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Dim intTotalElements As Integer

On Error GoTo 50

'Check if Form 1 list is already in the array
intTotalElements = UBound(CompanyDataList) - LBound(CompanyDataList) + 1
50:

'If not poll the database and fill the array
'If intTotalElements = 0 Then

Call GetCompanyData(CompanyDataList())

'End If

With Me.ListBox1
    .Clear
    .List = CompanyDataList
End With

GoTo 9999
errexit:
MsgBox ("There was an Error in the UserForm_Initialize Routine")
9999:
End Sub


Private Sub TextBox1_Change()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
Me.ListBox1.List = Filter(CompanyDataList, VBA.UCase(TextBox1.Value), True, vbTextCompare)
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
Unload Me
GoTo 9999
errexit:
MsgBox ("There was an Error in the PhaseOK_Click Routine")
9999:

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Call SetDataVal
Unload Me
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

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then
        Sheet17.Range("StartCompany").Value = VBA.Left(ListBox1.Column(0, i), 3)
        Sheet17.Range("StartCoName").Value = VBA.Trim(VBA.Mid(ListBox1.Column(0, i), 3 + 1, 30))
    End If
Next

GoTo 9999
errexit:
MsgBox "There was an Error in the SetDataVal Routine. " & Err.Number & ": " & Err.Description
9999:
Sheet17.Range("StartDept").ClearContents
Sheet17.Range("StartDept").Offset(0, 1).ClearContents
Sheet2.Range("Departments").Value = ""

Call GetDept(Sheet2.Range("UserName2").Value, Sheet17.Range("StartCompany").Value)

Application.EnableEvents = False
With Sheet17.Range("StartMonth").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 5287936
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Sheet17.Range("StartMonth").Offset(0, 1).Value = ""
Sheet17.Range("StartMonth").ClearContents
'Sheet17.Range("BatchId").ClearContents

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
MsgBox ("There was an Error in the UserForm_KeyPress Routine")
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



