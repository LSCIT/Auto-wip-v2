Attribute VB_Name = "Module6"
' Module6 — Phase 1 Complete Replacement
' Modified: March 2026 — Vista Direct Connection, Phase 1 guards
' Changes vs original:
'   - UseExistingBatch: WipDb batch logic removed; shows Dept picker only
'   - UsePostedBatch: Guarded (no WipDb batch check needed in Phase 1)
'   - CreateBatch: Phase 1 guard (Exit Sub)
'   - ClearWIPDetail: Phase 1 guard (Exit Sub)
'   - UpdateRow: Phase 1 guard (Exit Sub)
'   - UpdateRowJV: Phase 1 guard (Exit Sub)
'   - UpdateApprovals: Phase 1 guard (Exit Sub)
'   - CompleteCheck: Always returns True (no WipDb check in Phase 1)
'   - MarkAllComplete: Phase 1 guard (Exit Function, sheet-only update kept)
'   - GLCheck: Queries Vista bGLCO.LastMthSubClsd instead of WipDb LCGWIPGLCheck
'   - ApprCheck2: Phase 1 guard (Exit Sub)
'   - GetGAAPRev/Cost, GetOpsRev/Cost, GetOpsBonusPlug, HexToByteArray,
'     UpdateAllRows, SetTotals, SetDTotals, SetGrandTotals, ColLetToNum,
'     ProtectUnProtect, GetMissingItems: unchanged from original

' ========================================
' UseExistingBatch — C6 Implementation
' Original: Connected to WipDb, called LCGWIPBatchCheck1
' C6: Shows Dept picker, then reads LylesWIP batch state and maps it onto
'     the four Settings approval flags so all downstream button/lock logic
'     works correctly without any other code changes.
'
' State → flag mapping (matches what FormButtons.bas sets):
'   Open         → all "N"
'   ReadyForOps  → ReadyForOpsAppr1="Y"
'   OpsApproved  → ReadyForOpsAppr1="Y", FinalAppr="Y"
'   AcctApproved → ReadyForOpsAppr1="Y", FinalAppr="Y", AcctAppr="Y"
'
' Note: InitAppr is a legacy flag (commented out in FormButtons OFAYes_Click).
' It is reset to "N" here for cleanliness but never set to "Y".
'
' Batch creation (INSERT into WipBatches) happens automatically in GetWipDetail2
' via LylesWIPData.CreateBatch — not here.
' ========================================
Public Sub UseExistingBatch(ByRef NoData As Boolean)
    On Error GoTo errexit

    ProtectUnProtect ("Off")
    Application.EnableEvents = False

    ' Show dept picker
    Dim Deptfrm As Dept
    Set Deptfrm = New Dept
    Deptfrm.StartUpPosition = 0
    Deptfrm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Deptfrm.Width)
    Deptfrm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Deptfrm.Height)

    Application.EnableEvents = False
    Deptfrm.Show

    If DeptSelectionCanceled Then
        NoData = True
        GoTo 9999
    End If

    NoData = False

    ' C6: Read LylesWIP batch state and set approval flags on Settings sheet.
    ' Dept picker has now set Sheet17.Range("StartDept"); co and month already set.
    Dim co       As Integer
    Dim wipMonth As Date
    Dim dept     As String
    co       = CInt(Sheet17.Range("StartCompany").Value)
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)
    dept     = CStr(Sheet17.Range("StartDept").Value)

    ' Reset all four flags to N before applying state
    Sheet2.Range("ReadyForOpsAppr1").Value = "N"
    Sheet2.Range("InitAppr").Value         = "N"   ' legacy — always stays "N"
    Sheet2.Range("FinalAppr").Value        = "N"
    Sheet2.Range("AcctAppr").Value         = "N"

    ' GetBatchState returns "" if no batch exists yet — defaults to Open (all N)
    Dim batchState As String
    batchState = LylesWIPData.GetBatchState(co, wipMonth, dept)

    Select Case batchState
        Case "ReadyForOps"
            Sheet2.Range("ReadyForOpsAppr1").Value = "Y"
        Case "OpsApproved"
            Sheet2.Range("ReadyForOpsAppr1").Value = "Y"
            Sheet2.Range("FinalAppr").Value        = "Y"   ' enables AFAYes guard
        Case "AcctApproved"
            Sheet2.Range("ReadyForOpsAppr1").Value = "Y"
            Sheet2.Range("FinalAppr").Value        = "Y"
            Sheet2.Range("AcctAppr").Value         = "Y"
        ' Case "Open" and "" → all flags already "N"
    End Select

    GoTo 9999
errexit:
    MsgBox "There was an error in the UseExistingBatch Routine. " & Err.Description, vbOKOnly

9999:
    ProtectUnProtect ("On")
    Application.EnableEvents = True

End Sub


' Phase 1: UsePostedBatch — no WipDb posted batch check needed
Public Function UsePostedBatch() As Boolean
    ' Phase 1: No posted batch management
    UsePostedBatch = False
End Function


Function GetMissingItems(ByVal listIn As String, ByVal listCheckAgainst As String) As String
    Dim arrIn() As String
    Dim arrCheck() As String
    Dim item As Variant
    Dim Dict As Object
    Dim missingItems As Collection
    Dim missing As String

    ' Split the strings into arrays
    arrIn = Split(listIn, ",")
    arrCheck = Split(listCheckAgainst, ",")

    ' Create a dictionary for fast lookup
    Set Dict = CreateObject("Scripting.Dictionary")
    Set missingItems = New Collection

    ' Add items from listCheckAgainst to dictionary
    For Each item In arrCheck
        If Not Dict.Exists(Trim(item)) Then
            Dict.Add Trim(item), 1
        End If
    Next item

    ' Check for missing items
    For Each item In arrIn
        If Not Dict.Exists(Trim(item)) Then
            missingItems.Add Trim(item)
        End If
    Next item

    ' Convert collection to comma-delimited string
    For Each item In missingItems
        missing = missing & item & ","
    Next item

    ' Remove the trailing comma if there were missing items
    If Len(missing) > 0 Then
        missing = Left(missing, Len(missing) - 1)
    End If

    GetMissingItems = missing
End Function


' CreateBatch (Start sheet button stub)
' Batch creation is now automatic — LylesWIPData.CreateBatch is called from
' GetWipDetail2 after every successful Vista load (C2). This sub exists only
' because the Start sheet may have a "Create Batch" button pointing to it.
Public Sub CreateBatch()
    Exit Sub
End Sub


' Phase 1 guard — no WipDb batch clear
Public Sub ClearWIPDetail()
    ' Phase 1: WipDb write path disabled
    Exit Sub
End Sub


Public Function GetGAAPRev(row As Long, sh As Worksheet) As Double
On Error Resume Next
GetGAAPRev = 0
If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOR")).Value = "T" Then
    GetGAAPRev = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRevNew")).Value
Else
    GetGAAPRev = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRev")).Value
End If
End Function

Public Function GetGAAPRevPlug(row As Long, sh As Worksheet) As String
On Error Resume Next
GetGAAPRevPlug = "N"
If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOR")).Value = "T" Then
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRevNew")).Font.Bold = True Then GetGAAPRevPlug = "Y"
Else
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPRev")).Font.Bold = True Then GetGAAPRevPlug = "Y"
End If
End Function


Public Function GetGAAPCost(row As Long, sh As Worksheet) As Double
On Error Resume Next
GetGAAPCost = 0
If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOP")).Value = "T" Then
    GetGAAPCost = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCostNew")).Value
Else
    GetGAAPCost = sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCost")).Value
End If
End Function

Public Function GetGAAPCostPlug(row As Long, sh As Worksheet) As String
On Error Resume Next
GetGAAPCostPlug = "N"
If sh.Cells(row, NumDict(sh.CodeName)("COLZJCOP")).Value = "T" Then
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCostNew")).Font.Bold = True Then GetGAAPCostPlug = "Y"
Else
    If sh.Cells(row, NumDict(sh.CodeName)("COLZGAAPCost")).Font.Bold = True Then GetGAAPCostPlug = "Y"
End If
End Function


Public Function GetOpsCost(row As Long, sh As Worksheet) As Double
On Error Resume Next
GetOpsCost = 0
If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCChg")).Value = "T" Then
    GetOpsCost = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCostNew")).Value
Else
    GetOpsCost = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCost")).Value
End If
End Function

Public Function GetOpsCostPlug(row As Long, sh As Worksheet) As String
On Error Resume Next
GetOpsCostPlug = "N"
If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCChg")).Value = "T" Then
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCostNew")).Font.Bold = True Then GetOpsCostPlug = "Y"
Else
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsCost")).Font.Bold = True Then GetOpsCostPlug = "Y"
End If
End Function


Public Function GetOpsRev(row As Long, sh As Worksheet) As Double
On Error Resume Next
GetOpsRev = 0
If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRChg")).Value = "T" Then
    GetOpsRev = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRevNew")).Value
Else
    GetOpsRev = sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRev")).Value
End If
End Function

Public Function GetOpsRevPlug(row As Long, sh As Worksheet) As String
On Error Resume Next
GetOpsRevPlug = "N"
If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRChg")).Value = "T" Then
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRevNew")).Font.Bold = True Then GetOpsRevPlug = "Y"
Else
    If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsRev")).Font.Bold = True Then GetOpsRevPlug = "Y"
End If
End Function


Public Function GetOpsBonusPlug(row As Long, sh As Worksheet) As String
On Error Resume Next
GetOpsBonusPlug = "N"
If sh.Cells(row, NumDict(sh.CodeName)("COLZOPsBonusNew")).Font.Bold = True Then GetOpsBonusPlug = "Y"
End Function


Function HexToByteArray(hexString As String) As Byte()
    Dim byteArray() As Byte
    Dim i As Long
    Dim length As Long

    ' Remove any spaces from the hex string
    hexString = Replace(hexString, " ", "")

    ' Ensure the string length is even (since each byte is represented by two characters)
    If Len(hexString) Mod 2 <> 0 Then
        Err.Raise 5, "HexToByteArray", "Hex string must have an even number of characters."
        Exit Function
    End If

    length = Len(hexString)
    ReDim byteArray(length \ 2 - 1) ' Each byte needs two chars

    ' Iterate over the string two characters at a time
    For i = 0 To length - 1 Step 2
        ' Convert two hex characters to a byte
        byteArray(i \ 2) = CLng("&H" & Mid(hexString, i + 1, 2))
    Next i

    HexToByteArray = byteArray
End Function


Public Sub UpdateAllRows(sh As Worksheet, Complete As String)

Dim cell As Range
Dim rng As Range
Dim ColOffset As Integer
Dim cp As String
If Complete = "Y" Then
    cp = "P"
Else
    cp = ""
End If


If sh.CodeName = "Sheet11" Then ' Ops

    ColOffset = NumDict(sh.CodeName)("COLJobNumber") - NumDict(sh.CodeName)("COLDone")


    Set rng = Sheet11.Range("Done")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If

        If cp = "" Then
            cell.ClearComments
        End If
    Next cell



Else ' GAAP

    ColOffset = NumDict(sh.CodeName)("COLJobNumber") - NumDict(sh.CodeName)("COLGAAPDone")

    Set rng = Sheet12.Range("DoneGAAP")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If

        If cp = "" Then
            cell.ClearComments
        End If

    Next cell

End If




For Each cell In sh.Range("MyJobNos")

If cell.Value <> "" Then

    Call UpdateRow(cell.row, sh)

End If

Next cell




End Sub


' C4/C5: Save job row to LylesWIP
' Called from Sheet11/Sheet12 BeforeDoubleClick on col H (OpsDone) or col I (GAAPDone),
' and from col G (Close). Also called from Sheet12 for GAAP overrides.
' SaveJobRow reads all current override fields from the row and upserts to LylesWIP.
Public Sub UpdateRow(row As Long, sh As Worksheet)
    Dim co       As Integer
    Dim wipMonth As Date
    Dim userName As String

    co       = CInt(Sheet17.Range("StartCompany").Value)
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)
    userName = CStr(Sheet2.Range("UserName2").Value)
    If userName = "" Then userName = Environ("UserName")

    LylesWIPData.SaveJobRow sh, row, co, wipMonth, userName
End Sub


' UpdateRowJV — JV sheet write path (Sprint 5, not yet implemented)
Public Sub UpdateRowJV(row As Long, sh As Worksheet)
    ' JV override persistence deferred — confirm JV override workflow with Nicole first
    Exit Sub
End Sub


' ========================================
' UpdateApprovals — C7 Implementation
' All six approval buttons (RFOYes/No, OFAYes/No, AFAYes/No) set their flags
' then call this sub. We read the four flags, derive the current state string,
' and persist it to LylesWIP.WipBatches.
'
' Flag → state derivation (checked highest to lowest):
'   AcctAppr="Y"            → "AcctApproved"
'   FinalAppr="Y"           → "OpsApproved"
'   ReadyForOpsAppr1="Y"    → "ReadyForOps"
'   else                    → "Open"
'
' Note: FormButtons_RFOYes_Updated.bas is OBSOLETE — do not deploy.
' FormButtons.bas already calls this sub; no changes to FormButtons.bas needed.
' ========================================
Public Sub UpdateApprovals()
    On Error GoTo errexit

    Dim co       As Integer
    Dim wipMonth As Date
    Dim dept     As String
    Dim userName As String

    co       = CInt(Sheet17.Range("StartCompany").Value)
    wipMonth = CDate(Sheet17.Range("StartMonth").Value)
    dept     = CStr(Sheet17.Range("StartDept").Value)
    userName = CStr(Sheet2.Range("UserName2").Value)
    If userName = "" Then userName = Environ("UserName")

    Dim newState As String
    If Sheet2.Range("AcctAppr").Value = "Y" Then
        newState = "AcctApproved"
    ElseIf Sheet2.Range("FinalAppr").Value = "Y" Then
        newState = "OpsApproved"
    ElseIf Sheet2.Range("ReadyForOpsAppr1").Value = "Y" Then
        newState = "ReadyForOps"
    Else
        newState = "Open"
    End If

    LylesWIPData.UpdateBatchState co, wipMonth, dept, newState, userName

    ' December AcctApproved: capture year-end snapshot only when ALL departments
    ' for this company/month are AcctApproved. Check WipBatches for any non-approved.
    If newState = "AcctApproved" And Month(wipMonth) = 12 Then
        If LylesWIPData.AllBatchesApproved(co, wipMonth) Then
            Dim snapCount As Long
            snapCount = LylesWIPData.SaveYearEndSnapshot(co, wipMonth)
            If snapCount >= 0 Then
                MsgBox "All divisions approved. Year-end snapshot saved: " & snapCount & _
                       " jobs archived for " & Year(wipMonth) & ".", vbInformation, "Year-End Snapshot"
            End If
        End If
    End If

    GoTo 9999
errexit:
    MsgBox "There was an error saving the approval state. " & Err.Description, vbOKOnly
9999:
End Sub


' ========================================
' CompleteCheck — C8 Implementation
' Original: Called LCGWIPCompleteCheck stored proc on WipDb
' C8: Sheet-based check — iterate Done / DoneGAAP named range and verify
'     every row that has a job number is marked "P" (complete).
'
' CompType:
'   "O" → check Jobs-Ops  (Sheet11 Done range,     COLDone column)
'   "G" → check Jobs-GAAP (Sheet12 DoneGAAP range, COLGAAPDone column)
'
' Caller:
'   ""        → interactive: show MsgBox if incomplete (called from approval buttons)
'   "OrigRun" → silent: no MsgBox, result stored in CompleteAll/CompleteAllGAAP by caller
'
' A row is skipped if its job number cell is empty (subtotal/grand-total rows).
' ========================================
Public Function CompleteCheck(CompType As String, Caller As String) As Boolean
    On Error GoTo errexit

    Dim rng       As Range
    Dim cell      As Range
    Dim colOffset As Integer
    Dim remaining As Integer
    remaining = 0

    If CompType = "O" Then
        colOffset = NumDict("Sheet11")("COLJobNumber") - NumDict("Sheet11")("COLDone")
        Set rng = Sheet11.Range("Done")
        For Each cell In rng
            If cell.Offset(0, colOffset).Value <> "" Then
                If cell.Value <> "P" Then remaining = remaining + 1
            End If
        Next cell
    Else
        colOffset = NumDict("Sheet12")("COLJobNumber") - NumDict("Sheet12")("COLGAAPDone")
        Set rng = Sheet12.Range("DoneGAAP")
        For Each cell In rng
            If cell.Offset(0, colOffset).Value <> "" Then
                If cell.Value <> "P" Then remaining = remaining + 1
            End If
        Next cell
    End If

    If remaining > 0 Then
        If Caller = "" Then
            Dim typeLabel As String
            typeLabel = IIf(CompType = "O", "Ops", "GAAP")
            MsgBox remaining & " job(s) are not yet marked " & typeLabel & " Done." & vbCrLf & _
                   "Please mark all rows complete before approving.", vbInformation, "Complete Check"
        End If
        CompleteCheck = False
    Else
        CompleteCheck = True
    End If

    GoTo 9999
errexit:
    MsgBox "There was an error in the CompleteCheck Routine. " & Err.Description, vbOKOnly
    CompleteCheck = False
9999:
End Function


' Phase 1: Sheet-only update — no WipDb call
Public Function MarkAllComplete(CodeName As String, Complete As String, Field As String)

Dim rng As Range
Dim cell As Range
Dim ColOffset As Integer
Dim cp As String
If Complete = "Y" Then
    cp = "P"
Else
    cp = ""
End If


If Field = "Done" Then

    ColOffset = NumDict(CodeName)("COLJobNumber") - NumDict(CodeName)("COLDone")

    Set rng = Sheet11.Range("Done")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
    Next cell

    Set rng = Sheet12.Range("Done")
    ColOffset = NumDict("Sheet12")("COLJobNumber") - NumDict(CodeName)("COLDone")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
    Next cell



Else

    ColOffset = NumDict(CodeName)("COLJobNumber") - NumDict(CodeName)("COLGAAPDone")

    Set rng = Sheet12.Range("DoneGAAP")

    For Each cell In rng
        If cell.Offset(0, ColOffset).Value <> "" Then
            cell.Value = cp
        End If
    Next cell

End If

End Function


' ========================================
' GLCheck — Phase 1 Replacement (Vista Direct)
' Original: Connected to WipDb, called LCGWIPGLCheck stored proc
' Phase 1:  Queries bGLCO.LastMthSubClsd directly from Vista.
' ========================================
Public Sub GLCheck()
    On Error GoTo errexit

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        ' Can't check — default to Open
        Sheet17.Range("StartMonth").Offset(0, 1).Value = ""
        GoTo 9999
    End If

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim co As Integer
    co = CInt(Sheet17.Range("StartCompany").Value)

    sql = "SELECT LastMthSubClsd FROM bGLCO WITH (NOLOCK) WHERE GLCo = " & co

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        Sheet2.Range("LastClosedMth").Value = rs.Fields("LastMthSubClsd").Value

        If Sheet2.Range("LastClosedMth").Value >= Sheet17.Range("StartMonth").Value Then
            ' Red — month is closed
            With Sheet17.Range("StartMonth").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Sheet17.Range("StartMonth").Offset(0, 1).Value = "Closed!"
        Else
            ' Green — month is open
            With Sheet17.Range("StartMonth").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 5287936
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Sheet17.Range("StartMonth").Offset(0, 1).Value = ""
        End If
    Else
        ' No GL record for this company — show as open
        Sheet17.Range("StartMonth").Offset(0, 1).Value = ""
    End If

    rs.Close
    Set rs = Nothing

    GoTo 9999
errexit:
    Sheet17.Range("StartMonth").Offset(0, 1).Value = ""

9999:

End Sub


' Phase 1 guard — no WipDb approval check
Public Sub ApprCheck2()
    ' Phase 1: WipDb write path disabled
    Exit Sub
End Sub


Public Sub SetTotals(r As Integer, StartRow As Integer, EndRow As Integer, Range As String, sh As Worksheet)
Dim Col As String
Dim colNum As String
Col = LetDict(sh.CodeName)(Range)
colNum = NumDict(sh.CodeName)(Range)

sh.Range("SummaryData").Cells(r, CInt(colNum)).Formula = "=Sum(" & Col & CStr(StartRow) & ":" & Col & CStr(EndRow) & ")"
sh.Range("SummaryData").Cells(r, CInt(colNum)).Font.Bold = True

End Sub

Public Sub SetDTotals(r As Integer, StartRow As Integer, EndRow As Integer, Range As String, sh As Worksheet)
Dim Col As String
Dim colNum As String

Col = LetDict(sh.CodeName)(Range)
colNum = NumDict(sh.CodeName)(Range)

sh.Range("SummaryData").Cells(r, CInt(colNum)).Formula = "=Sum(" & Col & CStr(StartRow) & ":" & Col & CStr(EndRow) & ")/2"
sh.Range("SummaryData").Cells(r, CInt(colNum)).Font.Bold = True

End Sub


Public Sub SetGrandTotals(r As Integer, OpenRowTotal As Integer, EndRow As Integer, Range As String, sh As Worksheet)
Dim Col As String
Dim colNum As String
Dim ContStat As String
Dim Div As String
Col = LetDict(sh.CodeName)(Range)
colNum = NumDict(sh.CodeName)(Range)
ContStat = LetDict(sh.CodeName)("COLZContractStatus")

If Sheet2.Range("SortOption").Value = 1 Then
    ' Open Job Totals
    sh.Range("SummaryData").Cells(r, CInt(colNum)).Formula = "=SUMIF(" & ContStat & "8:" & ContStat & EndRow & ",1, " & Col & "8" & ":" & Col & EndRow & ")"
    sh.Range("SummaryData").Cells(r, CInt(colNum)).Font.Bold = True

    ' Closed Job Totals
    sh.Range("SummaryData").Cells(r + 1, CInt(colNum)).Formula = "=SUMIF(" & ContStat & "8:" & ContStat & EndRow & ",2, " & Col & "8" & ":" & Col & EndRow & ")"
    sh.Range("SummaryData").Cells(r + 1, CInt(colNum)).Font.Bold = True
    Div = "4"
Else
    Div = "2"
End If


' Grand Total
sh.Range("SummaryData").Cells(r + 2, CInt(colNum)).Formula = "=Sum(" & Col & "8" & ":" & Col & EndRow + 1 & ")/" & Div
sh.Range("SummaryData").Cells(r + 2, CInt(colNum)).Font.Bold = True


End Sub


'   pass the column letter and returns the number of the column
Function ColLetToNum(InputLetter As String) As Integer
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If


Dim OutputNumber As Integer
Dim Leng As Integer
Dim i As Integer
On Error GoTo errexit
Leng = Len(InputLetter)
OutputNumber = 0

For i = 1 To Leng
   OutputNumber = (Asc(VBA.UCase(VBA.Mid(InputLetter, i, 1))) - 64) + OutputNumber * 26
Next i

ColLetToNum = OutputNumber  'Output the corresponding number
GoTo 9999
errexit:
MsgBox "There was an error in the ColLetToNum Routine. " & Err, vbOKOnly
9999:
End Function


Public Sub ProtectUnProtect(OnOff As String)

Application.ScreenUpdating = False

If OnOff = "On" Then

    If Sheet2.Range("ProtectSheet").Value = "True" Then
        Sheet11.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet12.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet13.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet14.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet15.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Sheet17.Protect "password", AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True

    Else
        Sheet11.Unprotect "password"
        Sheet12.Unprotect "password"
        Sheet13.Unprotect "password"
        Sheet14.Unprotect "password"
        Sheet15.Unprotect "password"
        Sheet17.Unprotect "password"

    End If

Else

    Sheet11.Unprotect "password"
    Sheet12.Unprotect "password"
    Sheet13.Unprotect "password"
    Sheet14.Unprotect "password"
    Sheet15.Unprotect "password"
    Sheet17.Unprotect "password"

End If

9999:

'Application.ScreenUpdating = True


End Sub
