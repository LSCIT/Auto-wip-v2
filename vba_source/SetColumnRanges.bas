Option Explicit
Public NumDict As Object ' Outer dictionary: SheetCodeName -> Inner Dictionary (ColumnName -> ColumnNumber)
Public LetDict As Object ' Outer dictionary: SheetCodeName -> Inner Dictionary (ColumnName -> ColumnLetter)

Public Sub SetColumns()
    ' Replaced by InitializeColumnDictionaries — stub for legacy UploadData references
End Sub

Public Sub InitCols()
Set NumDict = Nothing
Set LetDict = Nothing

If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If

Call WriteDictionaryToSheet

End Sub



Sub InitializeColumnDictionaries(NumDict As Object, LetDict As Object, rowNum As Long)
    On Error GoTo ErrorHandler
    
    ' Validate row number
    If rowNum < 1 Or rowNum > 1048576 Then
        MsgBox "Error: Invalid row number. Must be between 1 and 1048576.", vbExclamation
        Exit Sub
    End If
    
    ' Define the list of desired CodeNames
    Dim tempList As Object
    Set tempList = CreateObject("Scripting.Dictionary")
    ' Add desired CodeNames (adjust as needed)
    tempList.Add "Sheet11", True
    tempList.Add "Sheet12", True
    tempList.Add "Sheet13", True
    tempList.Add "Sheet14", True
    tempList.Add "Sheet15", True
    tempList.Add "Sheet16", True
    
    
    ' Collect matching worksheets
    Dim worksheets() As Worksheet
    Dim validCount As Long
    Dim ws As Worksheet
    validCount = 0
    
    ' First pass: count valid worksheets
    For Each ws In ThisWorkbook.worksheets
        If tempList.Exists(ws.CodeName) Then
            validCount = validCount + 1
        End If
    Next ws
    
    ' Check if any valid worksheets were found
    If validCount = 0 Then
        MsgBox "Error: No worksheets with specified CodeNames found.", vbExclamation
        Exit Sub
    End If
    
    ' Allocate array and populate with matching worksheets
    ReDim worksheets(0 To validCount - 1)
    Dim i As Long
    i = 0
    For Each ws In ThisWorkbook.worksheets
        If tempList.Exists(ws.CodeName) Then
            Set worksheets(i) = ws
            i = i + 1
        End If
    Next ws
    
    ' Validate worksheet list
    If IsEmpty(worksheets) Or Not IsArray(worksheets) Then
        MsgBox "Error: Worksheet list is empty or invalid.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize outer dictionaries
    If NumDict Is Nothing Then
        Set NumDict = CreateObject("Scripting.Dictionary")
        Set LetDict = CreateObject("Scripting.Dictionary")
    End If
    
    Dim initializedCount As Long
    initializedCount = 0
    
    ' Process each worksheet
    Dim wsVar As Variant ' Use Variant for For Each over array
    For Each wsVar In worksheets
        If wsVar Is Nothing Then
            MsgBox "Warning: Invalid worksheet in list.", vbExclamation
            GoTo NextSheet
        End If
        
        ' Cast to Worksheet
        Set ws = wsVar
        
        ' Create inner dictionaries
        Dim innerNumDict As Object
        Dim innerLetDict As Object
        Set innerNumDict = CreateObject("Scripting.Dictionary")
        Set innerLetDict = CreateObject("Scripting.Dictionary")
        
        Dim nm As Name
        Dim rng As Range
        Dim colNum As Long
        Dim colLetter As String
        Dim count As Long
        Dim localName As String
        
        
        count = 0
        ' Process sheet-scoped names
        For Each nm In ws.Names
            On Error Resume Next
            Set rng = nm.RefersToRange
            On Error GoTo ErrorHandler

            If Not rng Is Nothing Then
                If rng.Cells.count = 1 And rng.row = rowNum Then
                    colNum = rng.Column
                    colLetter = Split(Cells(1, colNum).Address, "$")(1)

                    ' Strip tab name from named range
                    localName = Replace(nm.Name, Chr(39) & Chr(39), "'")
                    localName = Replace(localName, "'" & ws.Name & "'!", "")
                    localName = Replace(localName, ws.Name & "!", "")

                    innerNumDict(localName) = colNum
                    innerLetDict(localName) = colLetter

                    count = count + 1
                Else
                End If
            End If
        Next nm


        
        ' Only add non-empty dictionaries to NumDict and LetDict
        If count > 0 Then
            If NumDict.Exists(ws.CodeName) Then
                NumDict(ws.CodeName).RemoveAll
                LetDict(ws.CodeName).RemoveAll
            End If
            NumDict.Add key:=ws.CodeName, item:=innerNumDict ' Correct assignment
            LetDict.Add key:=ws.CodeName, item:=innerLetDict ' Correct assignment
            initializedCount = initializedCount + 1
        End If
NextSheet:
    Next wsVar
    
    If initializedCount = 0 Then
        MsgBox "Error: No valid sheets were initialized.", vbExclamation
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In InitializeColumnDictionaries, processing sheet: " & IIf(ws Is Nothing, "Unknown", ws.CodeName) & ", row: " & rowNum, vbCritical
End Sub

Sub WriteDictionaryToSheet()
    On Error GoTo ErrorHandler
    
    ' Check if dictionaries are initialized
    If NumDict Is Nothing Or LetDict Is Nothing Then
        MsgBox "Error: Dictionaries are not initialized. Run InitializeColumnDictionaries first.", vbExclamation
        Exit Sub
    End If
    
    ' Define output worksheet (create new or use existing)
    Dim outputWs As Worksheet
    On Error Resume Next
    Set outputWs = ThisWorkbook.worksheets("DictionaryOutput")
    On Error GoTo ErrorHandler
    If outputWs Is Nothing Then
        Set outputWs = ThisWorkbook.worksheets.Add
        outputWs.Name = "DictionaryOutput"
    Else
        ' Clear existing content
        outputWs.Cells.Clear
    End If
    
    ' Write headers
    With outputWs
        .Cells(1, 1).Value = "Sheet CodeName"
        .Cells(1, 2).Value = "Named Range"
        .Cells(1, 3).Value = "Column Number"
        .Cells(1, 4).Value = "Column Letter"
    End With
    
    ' Initialize row counter for output
    Dim row As Long
    row = 2 ' Start writing data from row 2
    
    ' Iterate through outer dictionaries
    Dim sheetCodeName As Variant
    Dim innerNumDict As Object
    Dim innerLetDict As Object
    Dim namedRange As Variant
    
    For Each sheetCodeName In NumDict.Keys
        ' Get inner dictionaries
        Set innerNumDict = NumDict(sheetCodeName)
        Set innerLetDict = LetDict(sheetCodeName)
        
        ' Iterate through named ranges in inner dictionaries
        For Each namedRange In innerNumDict.Keys
            With outputWs
                .Cells(row, 1).Value = sheetCodeName
                .Cells(row, 2).Value = namedRange
                .Cells(row, 3).Value = innerNumDict(namedRange)
                .Cells(row, 4).Value = innerLetDict(namedRange)
            End With
            row = row + 1
        Next namedRange
    Next sheetCodeName
    
    ' AutoFit columns for better readability
    outputWs.Columns("A:D").AutoFit
    
    MsgBox "Dictionary values have been written to " & outputWs.Name & ".", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In WriteDictionaryToSheet", vbCritical
End Sub



Sub DeleteBrokenNamedRangesWithCOLinJobsGAAP()
    Dim ws As Worksheet
    Dim nm As Name
    Dim nameToCheck As String
    
    ' Try to get the "Jobs-GAAP" worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Jobs-Ops")
    On Error GoTo 0
    
    ' Check if worksheet exists
    If ws Is Nothing Then
        MsgBox "Worksheet 'Jobs-GAAP' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through named ranges scoped to the "Jobs-GAAP" worksheet
    For Each nm In ws.Names
        ' Get the name to check
        nameToCheck = nm.Name
        ' Remove 'Jobs-GAAP'! prefix if present
        If InStr(1, nameToCheck, "'Jobs-GAAP'!") = 1 Then
            nameToCheck = Mid(nameToCheck, Len("'Jobs-GAAP'!") + 1)
        End If
        
        ' Diagnostic output to verify values
        Debug.Print "Name: " & nm.Name & ", RefersTo: " & nm.RefersTo & ", nameToCheck: " & nameToCheck
        
        ' Check if the name starts with "COL" (case-insensitive)
        If UCase(Left(nameToCheck, 3)) = "COL" Then
            ' Check if RefersTo contains "#REF!"
            If InStr(1, nm.RefersTo, "#REF!") > 0 Then
                nm.Delete
'                Debug.Print "Deleted broken named range: " & nm.Name & " in 'Jobs-GAAP'"
            End If
        End If
    Next nm
    
    MsgBox "Completed deleting named ranges starting with 'COL' with '#REF!' in 'Jobs-GAAP' scope.", vbInformation
End Sub