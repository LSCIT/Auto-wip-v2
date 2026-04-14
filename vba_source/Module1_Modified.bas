Attribute VB_Name = "Module1"
' Module1 — Global variables + dropdown data retrieval
' Modified: March 2026 — Phase 1 Vista Direct Connection
' Changes:
'   - GetCompanyData: Now queries Vista (bHQCO) instead of WipDb (LCGWIPGetCoList1)
'   - GetDeptData: Now queries Vista (bJCDM) instead of WipDb (LCGWIPGetDeptData1)
'   - ClearTables: Commented out — WipDb batch tables not used in Phase 1
'   - Removed Workstation ID=MROBERTS from all connection strings (Bug C4)

' Holds pop up list values so we don't have to query the DB each time we open the pop up
Public DeptDataList() As Variant
Public CompanyDataList() As Variant
Public Form3Datalist() As Variant
Public PRange As Range
Public RBCaller As String
Public StrPos As Integer
Public uname As String
Public LastColumnCount As Long
Public Caller As String
Public BatchSelected As Boolean
Public BatchSelectionCanceled As Boolean
Public DeptSelectionCanceled As Boolean
Public DeptList As String


Public Sub ee()
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Public Sub eo()
Application.EnableEvents = False
End Sub

Public Sub p()

If Sheet2.Range("ProtectSheet").Value = "True" Then
        Sheet11.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet12.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet13.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet14.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet15.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet16.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True

    Else

        Sheet11.Unprotect "password"
        Sheet12.Unprotect "password"
        Sheet13.Unprotect "password"
        Sheet14.Unprotect "password"
        Sheet15.Unprotect "password"
        Sheet16.Unprotect "password"

    End If



End Sub


' Get Department list — NOW USES VISTA DIRECT
Public Sub GetDeptData(ByRef FormList As Variant)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If
Application.EnableEvents = False
Application.ScreenUpdating = False

Dim Datars As ADODB.Recordset
Dim i As Integer
Dim co As Integer

co = CInt(Sheet17.Range("StartCompany").Value)

' Query Vista directly for department list
Set Datars = GetVistaDepartmentList(co)

If Datars Is Nothing Then
    MsgBox "Could not retrieve departments from Vista.", vbCritical
    GoTo 9999
End If

If Datars.EOF <> True Then
    i = 0
    Do
        ReDim Preserve FormList(i)
        FormList(i) = PadString(CStr(Datars!Department), 10, "R") & " " & PadString(VBA.UCase(Datars!Description), 30, "R")
        i = i + 1
        Datars.MoveNext
    Loop Until Datars.EOF

    Datars.Close
    Set Datars = Nothing
Else
    Datars.Close
    Set Datars = Nothing
    MsgBox ("There are no Divisions for this Company.")
End If


GoTo 9999
errexit:
MsgBox "There was an error in the GetDeptData Routine. " & Err.Description, vbOKOnly
Application.EnableEvents = True
Application.ScreenUpdating = True

9999:

Application.EnableEvents = True
'Application.ScreenUpdating = True
End Sub


' Get Company List — NOW USES VISTA DIRECT
Public Sub GetCompanyData(ByRef FormList() As Variant)
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Application.EnableEvents = False
Application.ScreenUpdating = False

Dim Datars As ADODB.Recordset
Dim i As Integer

' Query Vista directly for company list
Set Datars = GetVistaCompanyList()

If Datars Is Nothing Then
    MsgBox "Could not retrieve companies from Vista.", vbCritical
    GoTo 9999
End If

If Datars.EOF <> True Then
    i = 0
    Do
        ReDim Preserve FormList(i)
        FormList(i) = PadString(CStr(Datars!JCCo & ""), 10, "R") & " " & PadString(VBA.UCase(Datars!Name & ""), 30, "R")
        i = i + 1
        Datars.MoveNext
    Loop Until Datars.EOF

    Datars.Close
    Set Datars = Nothing
Else
    Datars.Close
    Set Datars = Nothing
    MsgBox ("There are no Companies Available.")
End If


GoTo 9999
errexit:
MsgBox "There was an error in the GetCompanyData Routine. " & Err.Description, vbOKOnly
Application.EnableEvents = True
Application.ScreenUpdating = True

9999: Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub


' Clear Tables — Phase 1: No-op (WipDb batch tables not used)
' Original called LCGWIPClearTables on WipDb to clear batch data.
' In Phase 1 we read directly from Vista, no batches to clear.
Public Sub ClearTables()
    ' Phase 1: No WipDb batch management — nothing to clear
End Sub

