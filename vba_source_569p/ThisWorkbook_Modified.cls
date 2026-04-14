'   Created by Michael Roberts 05/06/2014
'   Modified: March 2026 — Phase 1 Vista Direct Connection
'   Changes:
'     - Bug C1 FIX: Removed hardcoded uname = "cjordan" — now uses actual Windows login
'     - Bug C4 FIX: Removed Workstation ID=MROBERTS from connection strings (in Permissions.bas)
'     - Phase 1: Security check uses Environ("USERNAME") for actual logged-in user


Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim c As Range
For Each c In Sheet11.Range("Done").Cells

    If c.Borders(xlEdgeTop).Weight = xlMedium Then

        If MsgBox("There are Un-Saved Op's Changes on the form." & vbCrLf & "                             Continue?", vbYesNo) = vbNo Then
            Cancel = True
            GoTo 99

        Else
            GoTo 99
        End If

    End If
Next c

For Each c In Sheet12.Range("DoneGAAP").Cells

    If c.Borders(xlEdgeTop).Weight = xlMedium Then

        If MsgBox("There are Un-Saved GAAP Changes on the form." & vbCrLf & "                             Continue?", vbYesNo) = vbNo Then
            Cancel = True
            GoTo 99

        Else
            GoTo 99
        End If

    End If
Next c


' Clear Batch prompt removed — batches are immutable after approval.
' Historical data stays in LylesWIP for prior-month lookups and year-end snapshots.



99:
' Clean up connections on close
Call CloseVistaConnection
Call LylesWIPData.CloseWIPConnection
ThisWorkbook.Saved = True



'Cancel = True
End Sub


Public Sub Workbook_Open()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Dim MyName As String
Dim MyDomain As String
Application.EnableEvents = False
Application.ScreenUpdating = False


' Domain for users system
Dim WSHnet, username, UserDomain, objUser
Set WSHnet = CreateObject("WScript.Network")

' ========================================
' BUG C1 FIX: Use actual Windows login
' BEFORE: uname = "cjordan" (hardcoded)
' AFTER:  uname = actual logged-in user
' ========================================
uname = LCase(CStr(WSHnet.username))
' uname = "cjordan"  ' REMOVED — was hardcoded override


BatchSelected = False

' Set Column Number and Letters in the Dictionary
If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If

'SetColumns
SetColWidth Sheet11
SetColWidth Sheet12

Application.OnKey "{F4}", "OpenForm"

ProtectUnProtect ("Off")

If Sheet2.Range("ClearFormOnOpen") = True Then

    Sheet2.Range("UserName2").Value = ""
    Sheet2.Range("Role").Value = ""
    Sheet2.Range("Companies").Value = ""
    Sheet2.Range("Departments").Value = ""

    ' Open splash form while looking up users Security and Role
    Set frm = New GetSecurityForm
    frm.StartUpPosition = 0
    ' Calculate the position where the UserForm should appear to be centered in Excel
    frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
    frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
    frm.Show vbModeless
    DoEvents

    Caller = "Workbook_Open"

    ' Clear all data entry forms
    ClearForms3

    ' Clear comments and Bold Format
    Call ClearFormOpsVsGAAP

    ' Gets User Role from the P&P Database
    Call GetSecurity(uname)

End If

Sheet2.Range("UserName2").Value = uname



If Sheet2.Range("ProtectSheet").Value = True Then
    ProtectUnProtect ("On")
End If
GoTo 9999
errexit:
MsgBox "There was an error in the Workbook_Open Routine. " & Err, vbOKOnly
9999:
Caller = ""
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True
Sheet17.Range("StartCompany").Select


Unload frm

Sheet2.Visible = xlSheetVeryHidden
'Sheet2.Visible = xlSheetVisible


End Sub
