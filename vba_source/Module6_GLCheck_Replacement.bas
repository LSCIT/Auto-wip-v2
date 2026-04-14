' ========================================
' GLCheck — Phase 1 Replacement (Vista Direct)
' REPLACES the GLCheck sub in Module6 (around line 1412)
'
' Original: Connected to WipDb, called LCGWIPGLCheck stored proc
'           to get last closed GL month from Michael's manual table.
'
' Phase 1:  Queries bGLCO.LastMthSubClsd directly from Vista.
'           Sub-ledger closed date is the correct field since WIP = JC sub-ledger.
'
' TO DEPLOY: In Module6, find "Public Sub GLCheck()" (around line 1412).
'            Select from there down to its "End Sub" (around line 1498).
'            Delete that entire block and paste this replacement.
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
