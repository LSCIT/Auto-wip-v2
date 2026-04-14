Attribute VB_Name = "BatchValidate"
' =============================================================================
' BatchValidate.bas
' Loops all 24 company/division combinations for Dec 2025, loads each one,
' and saves a copy to OUTPUT_PATH for systematic validation.
'
' Usage:
'   1. Update OUTPUT_PATH below to a folder on this machine (e.g. C:\WIP\validate-d3\)
'   2. Run BatchValidateAll from the VBA editor (Alt+F8 → BatchValidateAll)
'   3. Let it run unattended — progress shows in Excel title bar
'   4. Copy the 24 saved files to Mac: vm/validate-d3/
' =============================================================================

Private Const OUTPUT_PATH As String = "C:\Trusted\validate-d3\"
Private Const WIP_MONTH   As String = "12/1/2025"

Public Sub BatchValidateAll()
    Dim co       As Integer
    Dim dept     As Integer
    Dim outFile  As String
    Dim i        As Integer
    Dim errCount As Integer

    ' Verify output folder exists
    If Dir(OUTPUT_PATH, vbDirectory) = "" Then
        MsgBox "Output folder not found: " & OUTPUT_PATH & vbCrLf & _
               "Create the folder first, then re-run.", vbExclamation
        Exit Sub
    End If

    ' 24 combinations: Array(JCCo, Dept)
    ' WML (15): Div50 added — Nicole has 7 jobs there (50.0541–50.0551)
    ' APC (12): Div20 added — Nicole has 10 jobs there (20.252–20.2574)
    Dim combos As Variant
    combos = Array( _
        Array(15, 50), Array(15, 51), Array(15, 52), Array(15, 53), Array(15, 54), _
        Array(15, 55), Array(15, 56), Array(15, 57), Array(15, 58), _
        Array(16, 70), Array(16, 71), Array(16, 72), Array(16, 73), Array(16, 74), _
        Array(16, 75), Array(16, 76), Array(16, 77), Array(16, 78), _
        Array(12, 20), Array(12, 21), _
        Array(13, 31), Array(13, 32), Array(13, 33), Array(13, 35) _
    )

    Dim total As Integer
    total = UBound(combos) + 1

    For i = 0 To UBound(combos)
        co   = combos(i)(0)
        dept = combos(i)(1)

        Application.Caption = "WIP Batch " & (i + 1) & "/" & total & _
                               "  —  Co" & co & " Div" & dept & " ..."
        DoEvents

        ' --- Full reset (same as Clear Workbook button) for clean state ---
        Call ResetWorkbook

        ' --- Set Start sheet values without firing Worksheet_Change ---
        Application.EnableEvents = False
        Sheet17.Range("StartCompany").Value = co
        Sheet17.Range("StartMonth").Value   = WIP_MONTH
        Sheet17.Range("StartDept").Value    = dept
        Application.EnableEvents = True

        ' --- Load from Vista + merge LylesWIP overrides ---
        On Error Resume Next
        Err.Clear
        GetWipDetail2 Sheet11
        If Err.Number <> 0 Then
            errCount = errCount + 1
            Debug.Print "LOAD ERROR Co" & co & " Div" & dept & " Sheet11: " & Err.Description
            Err.Clear
        End If
        GetWipDetail2 Sheet12
        If Err.Number <> 0 Then
            errCount = errCount + 1
            Debug.Print "LOAD ERROR Co" & co & " Div" & dept & " Sheet12: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' --- Save copy ---
        outFile = OUTPUT_PATH & co & "-" & dept & ".xltm"
        On Error Resume Next
        ThisWorkbook.SaveCopyAs outFile
        If Err.Number <> 0 Then
            errCount = errCount + 1
            Debug.Print "SAVE ERROR Co" & co & " Div" & dept & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        DoEvents
    Next i

    ' Restore Excel state
    Application.Caption = "Microsoft Excel"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If errCount = 0 Then
        MsgBox "Done! " & total & " files saved to:" & vbCrLf & OUTPUT_PATH & vbCrLf & _
               "WML Div50 + APC Div20 now included.", vbInformation
    Else
        MsgBox "Done with " & errCount & " errors (see Immediate window)." & vbCrLf & _
               total & " files attempted → " & OUTPUT_PATH, vbExclamation
    End If
End Sub
