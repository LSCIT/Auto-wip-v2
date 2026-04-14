' ========================================
' UseExistingBatch — Phase 1 Replacement
' REPLACES the UseExistingBatch sub in Module6 (around line 2)
'
' Original: Connected to WipDb, called LCGWIPBatchCheck1 to check for
'           existing WIP batches. Managed batch creation, posted batch
'           loading, and department selection.
'
' Phase 1:  No WipDb batches. Just shows the Dept picker form.
'           Sets NoData = False so data loading proceeds after dept selection.
'
' TO DEPLOY: In Module6, find "Public Sub UseExistingBatch(ByRef NoData As Boolean)"
'            (at the very top, around line 2).
'            Select from there down to its "End Sub" (around line 137).
'            Delete that entire block and paste this replacement.
' ========================================
Public Sub UseExistingBatch(ByRef NoData As Boolean)
    On Error GoTo errexit

    ProtectUnProtect ("Off")
    Application.EnableEvents = False

    ' Phase 1: No batch management — just show dept picker
    Dim Deptfrm As Dept
    Set Deptfrm = New Dept
    Deptfrm.StartUpPosition = 0
    Deptfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Deptfrm.Width)
    Deptfrm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Deptfrm.Height)

    Application.EnableEvents = False
    Deptfrm.Show

    ' If user cancelled the dept picker, don't load data
    If DeptSelectionCanceled Then
        NoData = True
    Else
        NoData = False
    End If

    GoTo 9999
errexit:
    MsgBox "There was an error in the UseExistingBatch Routine. " & Err.Description, vbOKOnly

9999:
    ProtectUnProtect ("On")
    Application.EnableEvents = True

End Sub
