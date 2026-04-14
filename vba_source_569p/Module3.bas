
Public Function GetCol(Range As String, sh As Worksheet) As Integer

 If InStr(1, Range, "JV") > 0 Then
    GetCol = sh.Range(Range).Column
 Else
    GetCol = sh.Range(Range).Column
 End If
 
End Function

'Pass Range and get column Leter
Public Function GetColLet(ColRange As String, sh As Worksheet) As String
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

Dim Col As Integer
'MsgBox ColRange

If InStr(1, ColRange, "JV") > 0 Then
    Col = sh.Range(ColRange).Column
Else
    Col = sh.Range(ColRange).Column
End If

Dim vArr
vArr = Split(Cells(1, Col).Address(True, False), "$")
GetColLet = vArr(0)
GoTo 9999
errexit:
MsgBox "There was an error in the GetColLet Routine. " & Err, vbOKOnly
9999:


End Function

Public Sub CheckForUnSavedRows()
Dim c As Range
For Each c In Sheet7.Range("Done").Cells

    If c.Borders(xlEdgeTop).Weight = xlMedium Then
        
        If MsgBox("There are Un-Saved Changes on the form." & vbCrLf & "             Continue?", vbYesNo) = vbNo Then
            Cancel = True
        End If
    
    End If
Next c

End Sub


Public Function JTDProfitCheck(sh As Worksheet, TargetCol As Integer, TargetValue As Double, TargetRow As Integer) As Boolean

Dim OPR As Double
Dim OPC As Double
Dim JTDC As Double
Dim JTDER As Double
Dim JTDP As Double
Dim PP As Double

If TargetCol = CInt(NumDict(sh.CodeName)("COLOvrCostProj")) Then
    OPR = sh.Range(LetDict(sh.CodeName)("COLOvrRevProj") & TargetRow).Value
    OPC = TargetValue
Else
    OPR = TargetValue
    OPC = sh.Range(LetDict(sh.CodeName)("COLOvrCostProj") & TargetRow).Value
End If

JTDC = sh.Range(LetDict(sh.CodeName)("COLJTDCost") & TargetRow).Value

If OPC <> 0 Then
    JTDER = (JTDC / OPC) * OPR
Else
    JTDER = 0
End If

PP = OPR - OPC

JTDP = JTDER - JTDC

If JTDP > PP Then
    JTDProfitCheck = False
Else
    JTDProfitCheck = True
End If

End Function





Public Sub OpsToGAAP()
    On Error GoTo errexit

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Copy Ops overrides to GAAP in LylesWIP (only where GAAP is NULL)
    Dim rowsCopied As Long
    rowsCopied = LylesWIPData.CopyOpsToGAAP()

    If rowsCopied >= 0 Then
        ' Re-merge overrides onto GAAP sheet to show copied values
        ' (no Vista reload needed — just refresh from LylesWIP)
        Dim co As Integer
        Dim wipMonth As Date
        co = CInt(Sheet17.Range("StartCompany").Value)
        wipMonth = CDate(Sheet17.Range("StartMonth").Value)
        Sheet12.Unprotect "password"
        LylesWIPData.MergeOverridesOntoSheet Sheet12, co, wipMonth
        Sheet12.Protect "password"
    End If

    GoTo 9999
errexit:
    MsgBox "There was an error in the Copy Ops to GAAP routine: " & Err.Description, vbOKOnly
9999:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Public Sub UpdateCostBill()
If Sheet2.Range("ErrorCtl").Value = True Then
    On Error GoTo errexit
End If

If NumDict Is Nothing Then
    InitializeColumnDictionaries NumDict, LetDict, 1
End If

ProtectUnProtect ("Off")
'Declare variables'
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim DataSQL As String
Dim Datars As ADODB.Recordset
Set cnn = New ADODB.Connection

Set cmd = New ADODB.Command
Set Datars = New ADODB.Recordset
Dim i As Integer
Dim intTotalElements As Integer
Dim RetMsg As String
Dim frm As DataRetrievalStatus

Set frm = New DataRetrievalStatus
frm.Label1.Caption = "Updating Cost and Billed Amounts.."
frm.StartUpPosition = 0
' Calculate the position where the UserForm should appear to be centered in Excel
frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
frm.Show vbModeless
DoEvents


' Open Connection
' If username and password are blank on settings tab use integrated windows user and password for sql query
If CStr(Sheet2.Range("PPPassword").Value) = "" Or CStr(Sheet2.Range("PPUsername").Value) = "" Then
    cnn.connectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;"
Else
    cnn.connectionString = "Provider=SQLOLEDB.1;Password=" & CStr(Sheet2.Range("PPPassword").Value) & ";Persist Security Info=True;User ID=" & CStr(Sheet2.Range("PPUsername").Value) & ";Initial Catalog=" & CStr(Sheet2.Range("WIPDBName").Value) & ";Data Source=" & CStr(Sheet2.Range("PPServerName").Value) & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MROBERTS;Use Encryption for Data=False;Tag with column collation when possible=False;"
End If

cnn.Open

' Get Batch Count

Set cmd = New ADODB.Command
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LCGWIPUpdateCostBill"
cmd.CommandTimeout = 180

Set cmdCo = cmd.CreateParameter("@Co", adInteger, adParamInput, 30, Sheet17.Range("StartCompany").Value)
cmd.Parameters.Append cmdCo
Set cmdDept = cmd.CreateParameter("@Dept", adVarChar, adParamInput, 200, Sheet17.Range("StartDept").Value)
cmd.Parameters.Append cmdDept
Set cmdMonth = cmd.CreateParameter("@Month", adDate, adParamInput, 30, Sheet17.Range("StartMonth").Value)
cmd.Parameters.Append cmdMonth
Set cmdRetMsg = cmd.CreateParameter("@Msg", adVarChar, adParamOutput, 200)
cmd.Parameters.Append cmdRetMsg

RetMsg = cmd.Parameters("@Msg").Value

cmd.Execute

cnn.Close
Unload frm

GetWipDetail2 Sheet11
GetWipDetail2 Sheet12

Dim sh As Worksheet
Dim targetSheets As Variant
Dim r As Range

targetSheets = Array(Sheet11, Sheet12)

For i = LBound(targetSheets) To UBound(targetSheets)
    Set sh = targetSheets(i)
    sh.Unprotect "password"

        
    For Each r In sh.Range("SummaryData").Rows
        
        r.Cells(1, NumDict(sh.CodeName)("COLJTDCost")).ClearComments
        r.Cells(1, NumDict(sh.CodeName)("COLBILLBillings")).ClearComments
        r.Cells(1, NumDict(sh.CodeName)("COLCYCost")).ClearComments
        
        ' JTD Cost
        If r.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Value <> "" And r.Cells(1, NumDict(sh.CodeName)("COLJTDCost")).Value <> _
        r.Cells(1, NumDict(sh.CodeName)("COLZORGJTDCost")).Value Then
        
            r.Cells(1, NumDict(sh.CodeName)("COLJTDCost")).Font.Bold = True
            
            With r.Cells(1, NumDict(sh.CodeName)("COLJTDCost"))
                .AddComment.Text Text:="Original = " & CStr(Format(r.Cells(1, NumDict(sh.CodeName)("COLZORGJTDCost")).Value, "#,##0;(#,##0)"))
                .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 25
                .Comment.Shape.Width = 125
            End With
        
        End If
        ' CY Cost
        If r.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Value <> "" And r.Cells(1, NumDict(sh.CodeName)("COLCYCost")).Value <> _
        r.Cells(1, NumDict(sh.CodeName)("COLZORGCYCost")).Value Then
        
            r.Cells(1, NumDict(sh.CodeName)("COLCYCost")).Font.Bold = True
            
            With r.Cells(1, NumDict(sh.CodeName)("COLCYCost"))
                .AddComment.Text Text:="Original = " & CStr(Format(r.Cells(1, NumDict(sh.CodeName)("COLZORGCYCost")).Value, "#,##0;(#,##0)"))
                .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 25
                .Comment.Shape.Width = 125
            End With
        
        End If
        
        ' JTD Billing
        If r.Cells(1, NumDict(sh.CodeName)("COLJobNumber")).Value <> "" And r.Cells(1, NumDict(sh.CodeName)("COLBILLBillings")).Value <> _
        r.Cells(1, NumDict(sh.CodeName)("COLZORGBilledAmt")).Value Then
        
            r.Cells(1, NumDict(sh.CodeName)("COLBILLBillings")).Font.Bold = True
            
            With r.Cells(1, NumDict(sh.CodeName)("COLBILLBillings"))
                .AddComment.Text Text:="Original = " & CStr(Format(r.Cells(1, NumDict(sh.CodeName)("COLZORGBilledAmt")).Value, "#,##0;(#,##0)"))
                .Comment.Shape.AutoShapeType = msoShapeRoundedRectangle
                .Comment.Shape.TextFrame.Characters.Font.Name = "Arial"
                .Comment.Shape.TextFrame.Characters.Font.Size = 10
                .Comment.Shape.Height = 25
                .Comment.Shape.Width = 125
            End With
        
        End If
    
    Next r


Next i

GoTo 9999
errexit:
cnn.Close
If Datars.State <> 0 Then
    Datars.Close
End If
Unload frm

MsgBox "There was an error in the UpdateCostBill Routine. " & Err, vbOKOnly

9999:

If Sheet2.Range("ProtectSheet").Value = True Then
    ProtectUnProtect ("On")
End If


End Sub