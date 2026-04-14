Attribute VB_Name = "VistaData"
Option Explicit
' =============================================================================
' VistaData.bas — Vista Direct Connection Module
' Purpose:  Connect to Viewpoint/Vista SQL Server and retrieve WIP data
'           Replaces WipDb stored procedure calls (LCGWIPGetDetailPM)
' Created:  March 2026 — Phase 1 (read-only, replaces middleman DB)
' =============================================================================

' Module-level Vista connection
Private mVistaConn As ADODB.Connection

' =============================================================================
' OpenVistaConnection
' Opens ADODB connection to the Viewpoint database using Settings sheet config
' =============================================================================
Public Function OpenVistaConnection() As Boolean
    On Error GoTo ErrorHandler

    If Not mVistaConn Is Nothing Then
        If mVistaConn.State = adStateOpen Then
            OpenVistaConnection = True
            Exit Function
        End If
    End If

    Set mVistaConn = New ADODB.Connection

    Dim server As String
    Dim uid As String
    Dim pwd As String

    ' Read Vista connection from Settings sheet named ranges
    server = CStr(Sheet2.Range("VistaServerName").Value)

    If server = "" Then
        MsgBox "Vista server not configured in Settings sheet (VistaServerName).", vbCritical
        OpenVistaConnection = False
        Exit Function
    End If

    ' Read database name from Settings (defaults to Viewpoint if blank)
    Dim dbName As String
    dbName = ""
    On Error Resume Next
    dbName = CStr(Sheet2.Range("VistaDBName").Value)
    On Error GoTo ErrorHandler
    If dbName = "" Then dbName = "Viewpoint"

    ' Build connection string
    ' Try SQL Auth first (for test/dev), fall back to Windows Auth
    uid = ""
    pwd = ""
    On Error Resume Next
    uid = CStr(Sheet2.Range("VPUsername").Value)
    pwd = CStr(Sheet2.Range("VPPassword").Value)
    On Error GoTo ErrorHandler

    Dim connStr As String
    If uid <> "" And pwd <> "" Then
        ' SQL Authentication (test/dev environment)
        connStr = "Provider=MSOLEDBSQL;" & _
                  "Server=" & server & ";" & _
                  "Database=" & dbName & ";" & _
                  "UID=" & uid & ";" & _
                  "PWD=" & pwd & ";" & _
                  "Persist Security Info=False;" & _
                  "TrustServerCertificate=yes;" & _
                  "Packet Size=4096;"
    Else
        ' Windows Authentication (production)
        connStr = "Provider=MSOLEDBSQL;" & _
                  "Server=" & server & ";" & _
                  "Database=" & dbName & ";" & _
                  "Integrated Security=SSPI;" & _
                  "Persist Security Info=False;" & _
                  "TrustServerCertificate=yes;" & _
                  "Packet Size=4096;"
    End If

    mVistaConn.ConnectionString = connStr
    mVistaConn.CommandTimeout = 180  ' 3 minutes for large queries
    mVistaConn.Open

    OpenVistaConnection = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to connect to Vista server (" & server & "):" & vbCrLf & _
           Err.Description, vbCritical, "Vista Connection Error"
    OpenVistaConnection = False
End Function

' =============================================================================
' CloseVistaConnection
' Cleanly closes the Vista connection
' =============================================================================
Public Sub CloseVistaConnection()
    On Error Resume Next
    If Not mVistaConn Is Nothing Then
        If mVistaConn.State = adStateOpen Then
            mVistaConn.Close
        End If
        Set mVistaConn = Nothing
    End If
End Sub

' =============================================================================
' GetVistaConnection
' Returns the active Vista connection (opens if needed)
' =============================================================================
Public Function GetVistaConnection() As ADODB.Connection
    ' VBA does NOT short-circuit Or/And — must use nested If to avoid
    ' accessing .State on a Nothing object
    Dim needsConnection As Boolean
    needsConnection = False

    If mVistaConn Is Nothing Then
        needsConnection = True
    ElseIf mVistaConn.State <> adStateOpen Then
        needsConnection = True
    End If

    If needsConnection Then
        If Not OpenVistaConnection() Then
            Set GetVistaConnection = Nothing
            Exit Function
        End If
    End If
    Set GetVistaConnection = mVistaConn
End Function

' =============================================================================
' TestVistaConnection
' Quick connectivity test — called from Settings sheet button
' =============================================================================
Public Sub TestVistaConnection()
    If OpenVistaConnection() Then
        Dim rs As ADODB.Recordset
        Set rs = mVistaConn.Execute("SELECT @@SERVERNAME AS ServerName, DB_NAME() AS DatabaseName")
        MsgBox "Connected successfully!" & vbCrLf & _
               "Server: " & rs.Fields("ServerName").Value & vbCrLf & _
               "Database: " & rs.Fields("DatabaseName").Value, _
               vbInformation, "Vista Connection Test"
        rs.Close
        Set rs = Nothing
        CloseVistaConnection
    End If
End Sub

' =============================================================================
' BuildWIPQuery
' Assembles the Vista SQL query for WIP data retrieval
' Parameters:
'   co      — Company number (tinyint, e.g., 15 for WML)
'   month   — WIP month as date (first of month)
'   dept    — Department filter (comma-separated, e.g., "10,20,30")
'   groupBy — Sort order: "Department" or "PM"
' Returns:   SQL string ready for execution
' =============================================================================
Public Function BuildWIPQuery(co As Integer, wipMonth As Date, dept As String, groupBy As String) As String
    Dim sql As String
    Dim cutOffDate As String
    Dim startDate As String
    Dim priorYrEnd As String
    Dim marchPlug As String
    Dim monthStr As String

    ' Format dates for SQL Server
    monthStr = Format(wipMonth, "yyyy-MM-dd")
    cutOffDate = Format(DateSerial(Year(wipMonth), Month(wipMonth) + 1, 0), "yyyy-MM-dd") ' EOMONTH
    startDate = Format(DateSerial(Year(wipMonth), 1, 1), "yyyy-MM-dd")
    priorYrEnd = Format(DateSerial(Year(wipMonth) - 1, 12, 31), "yyyy-MM-dd")
    marchPlug = Format(DateSerial(Year(wipMonth), 3, 31), "yyyy-MM-dd")

    sql = ""

    ' Variable declarations — Rev 5.70p (2026-04-06)
    ' All date filters use Mth (fiscal month) to match Crystal Report "JC Cost and Revenue"
    sql = sql & "DECLARE @Co tinyint = " & co & ";" & vbCrLf
    sql = sql & "DECLARE @CutOffDate date = '" & cutOffDate & "';" & vbCrLf
    sql = sql & "DECLARE @StartDate date = '" & startDate & "';" & vbCrLf
    sql = sql & "DECLARE @PriorYrEnd date = '" & priorYrEnd & "';" & vbCrLf
    sql = sql & "DECLARE @CurrentDate date = GETDATE();" & vbCrLf
    sql = sql & "DECLARE @MarchPlug date = '" & marchPlug & "';" & vbCrLf
    sql = sql & "DECLARE @Month date = '" & monthStr & "';" & vbCrLf
    sql = sql & "DECLARE @BillingMth date = '" & monthStr & "';" & vbCrLf
    sql = sql & vbCrLf

    ' Department filter CTE
    sql = sql & ";WITH DeptFilter AS (" & vbCrLf
    sql = sql & "    SELECT LTRIM(RTRIM(value)) AS Department" & vbCrLf
    sql = sql & "    FROM STRING_SPLIT('" & Replace(dept, "'", "''") & "', ',')" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Job List CTE — Rev 5.70p: cost-exists inclusion + Mth filters
    sql = sql & "JobList AS (" & vbCrLf
    sql = sql & "    SELECT" & vbCrLf
    sql = sql & "        j.JCCo, j.Job, j.Description AS JobDescription," & vbCrLf
    sql = sql & "        j.Contract, j.JobStatus, j.ProjectMgr, j.ProjMinPct," & vbCrLf
    sql = sql & "        c.Department, c.Description AS ContractDescription," & vbCrLf
    sql = sql & "        c.OrigContractAmt, c.ContractAmt, c.BilledAmt, c.ReceivedAmt," & vbCrLf
    sql = sql & "        c.CurrentRetainAmt, c.ActualCloseDate AS CompletionDate," & vbCrLf
    sql = sql & "        c.ContractStatus, c.MonthClosed," & vbCrLf
    sql = sql & "        d.Description AS DeptDescription," & vbCrLf
    sql = sql & "        ISNULL(pm.Name, '') AS PM" & vbCrLf
    sql = sql & "    FROM bJCJM j WITH (NOLOCK)" & vbCrLf
    sql = sql & "    JOIN bJCCM c WITH (NOLOCK) ON j.JCCo = c.JCCo AND j.Contract = c.Contract" & vbCrLf
    sql = sql & "    JOIN bJCDM d WITH (NOLOCK) ON c.JCCo = d.JCCo AND c.Department = d.Department" & vbCrLf
    sql = sql & "    JOIN DeptFilter df ON c.Department = df.Department" & vbCrLf
    sql = sql & "    LEFT JOIN bJCMP pm WITH (NOLOCK) ON j.JCCo = pm.JCCo AND j.ProjectMgr = pm.ProjectMgr" & vbCrLf
    sql = sql & "    WHERE j.JCCo = @Co" & vbCrLf
    sql = sql & "      AND (c.StartMonth <= @CutOffDate" & vbCrLf
    sql = sql & "           OR EXISTS (SELECT 1 FROM bJCCD cd WITH (NOLOCK)" & vbCrLf
    sql = sql & "                      WHERE cd.JCCo = j.JCCo AND cd.Job = j.Job" & vbCrLf
    sql = sql & "                      AND cd.JCTransType NOT IN ('OE','CO','PF') AND cd.Mth <= @Month))" & vbCrLf
    sql = sql & "      AND (" & vbCrLf
    sql = sql & "        j.JobStatus = 1" & vbCrLf
    sql = sql & "        OR (j.JobStatus = 2 AND c.MonthClosed >= @StartDate)" & vbCrLf
    sql = sql & "        OR (j.JobStatus = 3 AND c.MonthClosed >= @Month)" & vbCrLf
    sql = sql & "        OR EXISTS (SELECT 1 FROM bJCCD cd WITH (NOLOCK)" & vbCrLf
    sql = sql & "            WHERE cd.JCCo = j.JCCo AND cd.Job = j.Job" & vbCrLf
    sql = sql & "            AND cd.JCTransType NOT IN ('OE','CO','PF')" & vbCrLf
    sql = sql & "            AND cd.Mth BETWEEN @StartDate AND @Month)" & vbCrLf
    sql = sql & "        OR EXISTS (SELECT 1 FROM bARTH ar WITH (NOLOCK)" & vbCrLf
    sql = sql & "            WHERE ar.JCCo = j.JCCo AND ar.Contract = c.Contract" & vbCrLf
    sql = sql & "            AND ar.ARTransType = 'I'" & vbCrLf
    sql = sql & "            AND ar.Mth BETWEEN @StartDate AND @BillingMth)" & vbCrLf
    sql = sql & "        OR EXISTS (SELECT 1 FROM bJCCD cd WITH (NOLOCK)" & vbCrLf
    sql = sql & "            WHERE cd.JCCo = j.JCCo AND cd.Job = j.Job" & vbCrLf
    sql = sql & "            AND cd.JCTransType NOT IN ('OE','CO','PF')" & vbCrLf
    sql = sql & "            AND cd.Mth BETWEEN @StartDate AND @Month" & vbCrLf
    sql = sql & "            AND EXISTS (SELECT 1 FROM bJCCD cd2 WITH (NOLOCK)" & vbCrLf
    sql = sql & "                WHERE cd2.JCCo = j.JCCo AND cd2.Job = j.Job" & vbCrLf
    sql = sql & "                AND cd2.JCTransType NOT IN ('OE','CO','PF')" & vbCrLf
    sql = sql & "                AND cd2.Mth <= @PriorYrEnd))" & vbCrLf
    sql = sql & "      )" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Job Costs CTE — Rev 5.70p: all filters use Mth (fiscal month)
    sql = sql & "JobCosts AS (" & vbCrLf
    sql = sql & "    SELECT d.JCCo, d.Job," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth <= @Month THEN d.ActualCost ELSE 0 END) AS ActualCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth BETWEEN @StartDate AND @Month THEN d.ActualCost ELSE 0 END) AS CYActualCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth = @Month THEN d.ActualCost ELSE 0 END) AS PeriodActualCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType = 'OE' AND d.Mth <= @Month THEN d.EstCost ELSE 0 END) AS OrigEstCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType = 'CO' AND d.Mth <= @Month THEN d.EstCost ELSE 0 END) AS COEstCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType = 'PF' AND d.Mth <= @Month THEN d.ProjCost ELSE 0 END) AS ProjFinalCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType = 'PF' AND d.Mth < @Month THEN d.ProjCost ELSE 0 END) AS PriorProjCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType = 'PF' AND d.Mth <= @MarchPlug THEN d.ProjCost ELSE 0 END) AS MarchProjCost" & vbCrLf
    sql = sql & "    FROM bJCCD d WITH (NOLOCK)" & vbCrLf
    sql = sql & "    WHERE d.JCCo = @Co" & vbCrLf
    sql = sql & "    GROUP BY d.JCCo, d.Job" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' CO Contract Amounts CTE — Rev 5.70p: Mth filter
    sql = sql & "COContracts AS (" & vbCrLf
    sql = sql & "    SELECT id.JCCo, j.Job," & vbCrLf
    sql = sql & "        SUM(CASE WHEN id.JCTransType = 'CO' AND id.Mth <= @Month THEN id.ContractAmt ELSE 0 END) AS COContractAmt" & vbCrLf
    sql = sql & "    FROM bJCID id WITH (NOLOCK)" & vbCrLf
    sql = sql & "    JOIN bJCJM j WITH (NOLOCK) ON id.JCCo = j.JCCo AND id.Contract = j.Contract" & vbCrLf
    sql = sql & "    WHERE id.JCCo = @Co" & vbCrLf
    sql = sql & "    GROUP BY id.JCCo, j.Job" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' March Baseline CTE
    sql = sql & "MarchBaseline AS (" & vbCrLf
    sql = sql & "    SELECT jor.JCCo, j.Job, jor.OtherAmount AS MarchProjRevenue" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "        SELECT JCCo, Contract, OtherAmount," & vbCrLf
    sql = sql & "            ROW_NUMBER() OVER (PARTITION BY JCCo, Contract ORDER BY Month DESC) AS rn" & vbCrLf
    sql = sql & "        FROM bJCOR WITH (NOLOCK)" & vbCrLf
    sql = sql & "        WHERE JCCo = @Co AND Month <= @MarchPlug" & vbCrLf
    sql = sql & "    ) jor" & vbCrLf
    sql = sql & "    JOIN bJCJM j WITH (NOLOCK) ON jor.JCCo = j.JCCo AND jor.Contract = j.Contract" & vbCrLf
    sql = sql & "    WHERE jor.rn = 1" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Prior-quarter GAAP Revenue snapshot (bJCOR) — most recent udPlugged='Y' before WIP month
    sql = sql & "PriorGAAPRev AS (" & vbCrLf
    sql = sql & "    SELECT jor.JCCo, j.Job," & vbCrLf
    sql = sql & "           ISNULL(jor.RevCost, 0) AS LastGAAPRev," & vbCrLf
    sql = sql & "           jor.udPlugged AS LastGAAPRevPlugged" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "        SELECT JCCo, Contract, RevCost, udPlugged," & vbCrLf
    sql = sql & "               ROW_NUMBER() OVER (PARTITION BY JCCo, Contract ORDER BY Month DESC) AS rn" & vbCrLf
    sql = sql & "        FROM bJCOR WITH (NOLOCK)" & vbCrLf
    sql = sql & "        WHERE JCCo = @Co AND udPlugged = 'Y' AND Month < @Month" & vbCrLf
    sql = sql & "    ) jor" & vbCrLf
    sql = sql & "    JOIN bJCJM j WITH (NOLOCK) ON j.JCCo = jor.JCCo AND j.Contract = jor.Contract" & vbCrLf
    sql = sql & "    WHERE jor.rn = 1" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Prior-quarter GAAP Cost snapshot (bJCOP) — most recent udPlugged='Y' before WIP month
    sql = sql & "PriorGAAPCost AS (" & vbCrLf
    sql = sql & "    SELECT op.JCCo, op.Job," & vbCrLf
    sql = sql & "           ISNULL(op.ProjCost, 0) AS LastGAAPCost," & vbCrLf
    sql = sql & "           op.udPlugged AS LastGAAPCostPlugged" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "        SELECT JCCo, Job, ProjCost, udPlugged," & vbCrLf
    sql = sql & "               ROW_NUMBER() OVER (PARTITION BY JCCo, Job ORDER BY Month DESC) AS rn" & vbCrLf
    sql = sql & "        FROM bJCOP WITH (NOLOCK)" & vbCrLf
    sql = sql & "        WHERE JCCo = @Co AND udPlugged = 'Y' AND Month < @Month" & vbCrLf
    sql = sql & "    ) op WHERE op.rn = 1" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Prior-year JTD cost — Rev 5.70p: Mth filter
    sql = sql & "PriorYearJobCosts AS (" & vbCrLf
    sql = sql & "    SELECT d.JCCo, d.Job," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF')" & vbCrLf
    sql = sql & "                      AND d.Mth <= @PriorYrEnd THEN d.ActualCost ELSE 0 END)" & vbCrLf
    sql = sql & "        AS PriorYrJTDCost" & vbCrLf
    sql = sql & "    FROM bJCCD d WITH (NOLOCK)" & vbCrLf
    sql = sql & "    WHERE d.JCCo = @Co" & vbCrLf
    sql = sql & "    GROUP BY d.JCCo, d.Job" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Billed amount from JB Progress Bills — Rev 5.70p: matches Crystal Report exactly
    sql = sql & "BilledThruMonth AS (" & vbCrLf
    sql = sql & "    SELECT pb.JBCo AS JCCo, pb.Contract," & vbCrLf
    sql = sql & "        SUM(pb.AmountBilled_ThisBill) AS BilledAmt" & vbCrLf
    sql = sql & "    FROM vrvJBProgressBills pb WITH (NOLOCK)" & vbCrLf
    sql = sql & "    WHERE pb.JBCo = @Co AND pb.BillMonth <= @Month" & vbCrLf
    sql = sql & "    GROUP BY pb.JBCo, pb.Contract" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Prior month JTD cost — Rev 5.70p: for recognized profit calculation
    sql = sql & "PriorMonthCosts AS (" & vbCrLf
    sql = sql & "    SELECT d.JCCo, d.Job," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('OE','CO','PF') AND d.Mth < @Month" & vbCrLf
    sql = sql & "                 THEN d.ActualCost ELSE 0 END) AS PriorMonthJTDCost" & vbCrLf
    sql = sql & "    FROM bJCCD d WITH (NOLOCK)" & vbCrLf
    sql = sql & "    WHERE d.JCCo = @Co" & vbCrLf
    sql = sql & "    GROUP BY d.JCCo, d.Job" & vbCrLf
    sql = sql & ")," & vbCrLf
    sql = sql & vbCrLf

    ' Most recent bJCOP projected cost as of Dec 31 prior year (udPlugged='Y')
    ' Provides the denominator for % complete at prior year end
    sql = sql & "PriorYearProjCost AS (" & vbCrLf
    sql = sql & "    SELECT op.JCCo, op.Job," & vbCrLf
    sql = sql & "           ISNULL(op.ProjCost, 0) AS PriorYrProjCost" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "        SELECT JCCo, Job, ProjCost," & vbCrLf
    sql = sql & "               ROW_NUMBER() OVER (PARTITION BY JCCo, Job ORDER BY Month DESC) AS rn" & vbCrLf
    sql = sql & "        FROM bJCOP WITH (NOLOCK)" & vbCrLf
    sql = sql & "        WHERE JCCo = @Co AND udPlugged = 'Y' AND Month <= @PriorYrEnd" & vbCrLf
    sql = sql & "    ) op WHERE op.rn = 1" & vbCrLf
    sql = sql & ")" & vbCrLf
    sql = sql & vbCrLf

    ' Final SELECT — Rev 5.70p: uses BilledThruMonth, PriorMonthCosts, PeriodActualCost
    sql = sql & "SELECT jl.JCCo, LTRIM(RTRIM(jl.Job)) AS Job, RTRIM(jl.Contract) AS Contract, jl.ContractDescription, jl.JobDescription," & vbCrLf
    sql = sql & "    jl.PM, jl.Department, jl.DeptDescription, jl.JobStatus, jl.ContractStatus," & vbCrLf
    sql = sql & "    jl.MonthClosed," & vbCrLf
    sql = sql & "    CASE WHEN jl.ContractStatus = 2 AND (jl.MonthClosed IS NULL OR jl.MonthClosed < @Month) THEN 1 ELSE 0 END AS VistaClosedAtWipDate," & vbCrLf
    sql = sql & "    jl.CompletionDate, jl.ProjectMgr, jl.ProjMinPct," & vbCrLf
    sql = sql & "    jl.OrigContractAmt, ISNULL(co.COContractAmt, 0) AS COContractAmt," & vbCrLf
    sql = sql & "    jl.ContractAmt AS ProjContract," & vbCrLf
    sql = sql & "    ISNULL(bt.BilledAmt, 0) AS BilledAmt," & vbCrLf
    sql = sql & "    jl.ReceivedAmt," & vbCrLf
    sql = sql & "    ISNULL(jc.ActualCost, 0) AS ActualCost," & vbCrLf
    sql = sql & "    ISNULL(jc.CYActualCost, 0) AS CYActualCost," & vbCrLf
    sql = sql & "    ISNULL(jc.PeriodActualCost, 0) AS PeriodActualCost," & vbCrLf
    sql = sql & "    ISNULL(jc.OrigEstCost, 0) AS OrigEstCost," & vbCrLf
    sql = sql & "    ISNULL(jc.COEstCost, 0) AS COEstCost," & vbCrLf
    sql = sql & "    CASE WHEN ISNULL(jc.ProjFinalCost, 0) = 0 THEN ISNULL(jc.ActualCost, 0)" & vbCrLf
    sql = sql & "         ELSE jc.ProjFinalCost END AS ProjCost," & vbCrLf
    sql = sql & "    ISNULL(jc.MarchProjCost, 0) AS MarchProjCost," & vbCrLf
    sql = sql & "    ISNULL(mb.MarchProjRevenue, 0) AS MarchProjRevenue," & vbCrLf

    ' Phase 1 defaults for override fields
    sql = sql & "    ISNULL(jc.ActualCost, 0) AS OrgActualCost," & vbCrLf
    sql = sql & "    ISNULL(jc.CYActualCost, 0) AS OrgCYActualCost," & vbCrLf
    sql = sql & "    ISNULL(bt.BilledAmt, 0) AS OrgBilledAmt," & vbCrLf
    sql = sql & "    0 AS OrgCYBilledAmt," & vbCrLf
    sql = sql & "    '' AS [Close], '' AS Completed, '' AS CompletedGAAP," & vbCrLf
    sql = sql & "    '' AS UserName, 0 AS BatchSeq, CAST(0 AS varbinary(8)) AS RowVersion," & vbCrLf

    ' Override fields (all zero/empty for Phase 1)
    sql = sql & "    0 AS OpsRev, '' AS OpsRevPlugged, 0 AS OpsCost, '' AS OpsCostPlugged," & vbCrLf
    sql = sql & "    0 AS GAAPRev, '' AS GAAPRevPlugged, 0 AS GAAPCost, '' AS GAAPCostPlugged," & vbCrLf
    sql = sql & "    0 AS BonusProfit, '' AS BonusProfitPlugged, '' AS BonusProfitNotes," & vbCrLf
    sql = sql & "    0 AS PriorYrBonusProfit," & vbCrLf

    ' Trend fields (all zero for Phase 1)
    sql = sql & "    0 AS LastProjContract, 0 AS LastProjContract2, 0 AS LastProjContract3," & vbCrLf
    sql = sql & "    0 AS LastProjContract4, 0 AS LastProjContract5, 0 AS LastProjContract6," & vbCrLf
    sql = sql & "    0 AS LastProjCost, 0 AS LastProjCost2, 0 AS LastProjCost3," & vbCrLf
    sql = sql & "    0 AS LastProjCost4, 0 AS LastProjCost5, 0 AS LastProjCost6," & vbCrLf
    sql = sql & "    0 AS LastOpsRev, 0 AS LastOpsRev2, 0 AS LastOpsRev3," & vbCrLf
    sql = sql & "    0 AS LastOpsRev4, 0 AS LastOpsRev5, 0 AS LastOpsRev6," & vbCrLf
    sql = sql & "    '' AS LastOpsRevPlugged, '' AS LastOpsRevPlugged2, '' AS LastOpsRevPlugged3," & vbCrLf
    sql = sql & "    '' AS LastOpsRevPlugged4, '' AS LastOpsRevPlugged5, '' AS LastOpsRevPlugged6," & vbCrLf
    sql = sql & "    0 AS LastOpsCost, 0 AS LastOpsCost2, 0 AS LastOpsCost3," & vbCrLf
    sql = sql & "    0 AS LastOpsCost4, 0 AS LastOpsCost5, 0 AS LastOpsCost6," & vbCrLf
    sql = sql & "    '' AS LastOpsCostPlugged, '' AS LastOpsCostPlugged2, '' AS LastOpsCostPlugged3," & vbCrLf
    sql = sql & "    '' AS LastOpsCostPlugged4, '' AS LastOpsCostPlugged5, '' AS LastOpsCostPlugged6," & vbCrLf
    sql = sql & "    ISNULL(pgr.LastGAAPRev, 0) AS LastGAAPRev, 0 AS LastGAAPRev2, 0 AS LastGAAPRev3," & vbCrLf
    sql = sql & "    0 AS LastGAAPRev4, 0 AS LastGAAPRev5, 0 AS LastGAAPRev6," & vbCrLf
    sql = sql & "    ISNULL(pgr.LastGAAPRevPlugged, '') AS LastGAAPRevPlugged, '' AS LastGAAPRevPlugged2, '' AS LastGAAPRevPlugged3," & vbCrLf
    sql = sql & "    '' AS LastGAAPRevPlugged4, '' AS LastGAAPRevPlugged5, '' AS LastGAAPRevPlugged6," & vbCrLf
    sql = sql & "    ISNULL(pgc.LastGAAPCost, 0) AS LastGAAPCost, 0 AS LastGAAPCost2, 0 AS LastGAAPCost3," & vbCrLf
    sql = sql & "    0 AS LastGAAPCost4, 0 AS LastGAAPCost5, 0 AS LastGAAPCost6," & vbCrLf
    sql = sql & "    ISNULL(pgc.LastGAAPCostPlugged, '') AS LastGAAPCostPlugged, '' AS LastGAAPCostPlugged2, '' AS LastGAAPCostPlugged3," & vbCrLf
    sql = sql & "    '' AS LastGAAPCostPlugged4, '' AS LastGAAPCostPlugged5, '' AS LastGAAPCostPlugged6," & vbCrLf
    sql = sql & "    0 AS LastBonusProfit, ISNULL(pmc.PriorMonthJTDCost, 0) AS LastActualCost," & vbCrLf
    ' Prior-year revenue/cost — apply 10% GAAP / 30% OPS threshold (Bug 4)
    ' If pct complete at Dec 31 < threshold: use JTD cost (cost-recovery method)
    ' If pct complete at Dec 31 >= threshold: use pct × ContractAmt
    sql = sql & "    CASE" & vbCrLf
    sql = sql & "        WHEN ISNULL(pyJOP.PriorYrProjCost, 0) > 0" & vbCrLf
    sql = sql & "             AND ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost >= 0.10" & vbCrLf
    sql = sql & "        THEN ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost * jl.ContractAmt" & vbCrLf
    sql = sql & "        ELSE ISNULL(pyJC.PriorYrJTDCost, 0)" & vbCrLf
    sql = sql & "    END AS PriorYearGAAPRev," & vbCrLf
    sql = sql & "    ISNULL(pyJOP.PriorYrProjCost, 0) AS PriorYearGAAPCost," & vbCrLf
    sql = sql & "    CASE" & vbCrLf
    sql = sql & "        WHEN ISNULL(pyJOP.PriorYrProjCost, 0) > 0" & vbCrLf
    sql = sql & "             AND ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost >= 0.30" & vbCrLf
    sql = sql & "        THEN ISNULL(pyJC.PriorYrJTDCost, 0) * 1.0 / pyJOP.PriorYrProjCost * jl.ContractAmt" & vbCrLf
    sql = sql & "        ELSE ISNULL(pyJC.PriorYrJTDCost, 0)" & vbCrLf
    sql = sql & "    END AS PriorYearOpsRev," & vbCrLf
    sql = sql & "    ISNULL(pyJOP.PriorYrProjCost, 0) AS PriorYearOpsCost," & vbCrLf
    sql = sql & "    '' AS OpsRevNotes, '' AS OpsCostNotes, '' AS GAAPRevNotes, '' AS GAAPCostNotes" & vbCrLf

    sql = sql & "FROM JobList jl" & vbCrLf
    sql = sql & "LEFT JOIN JobCosts jc ON jl.JCCo = jc.JCCo AND jl.Job = jc.Job" & vbCrLf
    sql = sql & "LEFT JOIN COContracts co ON jl.JCCo = co.JCCo AND jl.Job = co.Job" & vbCrLf
    sql = sql & "LEFT JOIN MarchBaseline mb ON jl.JCCo = mb.JCCo AND jl.Job = mb.Job" & vbCrLf
    sql = sql & "LEFT JOIN PriorGAAPRev pgr ON jl.JCCo = pgr.JCCo AND jl.Job = pgr.Job" & vbCrLf
    sql = sql & "LEFT JOIN PriorGAAPCost pgc ON jl.JCCo = pgc.JCCo AND jl.Job = pgc.Job" & vbCrLf
    sql = sql & "LEFT JOIN BilledThruMonth bt ON jl.JCCo = bt.JCCo AND jl.Contract = bt.Contract" & vbCrLf
    sql = sql & "LEFT JOIN PriorMonthCosts pmc ON jl.JCCo = pmc.JCCo AND jl.Job = pmc.Job" & vbCrLf
    sql = sql & "LEFT JOIN PriorYearJobCosts pyJC ON jl.JCCo = pyJC.JCCo AND jl.Job = pyJC.Job" & vbCrLf
    sql = sql & "LEFT JOIN PriorYearProjCost pyJOP ON jl.JCCo = pyJOP.JCCo AND jl.Job = pyJOP.Job" & vbCrLf

    ' Order by
    If groupBy = "Department" Then
        sql = sql & "ORDER BY jl.Department, jl.ContractStatus, jl.Contract;"
    Else
        sql = sql & "ORDER BY jl.PM, jl.Contract;"
    End If

    BuildWIPQuery = sql
End Function

' =============================================================================
' GetWIPDataFromVista
' Executes the WIP query and returns an ADODB Recordset
' Parameters:
'   co      — Company number
'   month   — WIP month (first of month)
'   dept    — Comma-separated department codes
'   groupBy — "Department" or "PM"
' Returns:   ADODB.Recordset (client-side cursor for bidirectional navigation)
' =============================================================================
' =============================================================================
' GetVistaCompanyList
' Returns company list from Vista (replaces LCGWIPGetCoList1 on WipDb)
' Only returns companies that have JC jobs
' =============================================================================
Public Function GetVistaCompanyList() As ADODB.Recordset
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        Set GetVistaCompanyList = Nothing
        Exit Function
    End If

    Dim sql As String
    sql = "SELECT HQCo AS JCCo, Name " & _
          "FROM bHQCO WITH (NOLOCK) " & _
          "WHERE HQCo IN (1,2,4,12,13,14,15,16,151) " & _
          "ORDER BY HQCo"

    Set GetVistaCompanyList = conn.Execute(sql)
    Exit Function

ErrorHandler:
    MsgBox "Error retrieving company list from Vista:" & vbCrLf & _
           Err.Description, vbCritical, "Vista Company List Error"
    Set GetVistaCompanyList = Nothing
End Function

' =============================================================================
' GetVistaCompanyName
' Returns company name from Vista (replaces LCGWIPGetCoName1 on WipDb)
' =============================================================================
Public Function GetVistaCompanyName(co As Integer) As String
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        GetVistaCompanyName = ""
        Exit Function
    End If

    Dim sql As String
    sql = "SELECT Name FROM bHQCO WITH (NOLOCK) WHERE HQCo = " & co

    Dim rs As ADODB.Recordset
    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        GetVistaCompanyName = rs.Fields("Name").Value
    Else
        GetVistaCompanyName = ""
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    GetVistaCompanyName = ""
End Function

' =============================================================================
' GetVistaDepartmentList
' Returns department list from Vista (replaces LCGWIPGetDeptData1 on WipDb)
' Returns all departments for the given company
' =============================================================================
Public Function GetVistaDepartmentList(co As Integer) As ADODB.Recordset
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        Set GetVistaDepartmentList = Nothing
        Exit Function
    End If

    Dim sql As String
    sql = "SELECT d.Department, d.Description " & _
          "FROM bJCDM d WITH (NOLOCK) " & _
          "WHERE d.JCCo = " & co & " " & _
          "ORDER BY d.Department"

    Set GetVistaDepartmentList = conn.Execute(sql)
    Exit Function

ErrorHandler:
    MsgBox "Error retrieving department list from Vista:" & vbCrLf & _
           Err.Description, vbCritical, "Vista Department List Error"
    Set GetVistaDepartmentList = Nothing
End Function

' =============================================================================
' GetJVDataFromVista
' Returns JV (Joint Venture) data from Vista
' Joins udWIPJV (JV master) with JC tables for financial data
' Replaces LCGWIPGetDetailJV stored proc on WipDb
' =============================================================================
Public Function GetJVDataFromVista(co As Integer, wipMonth As Date) As ADODB.Recordset
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        Set GetJVDataFromVista = Nothing
        Exit Function
    End If

    ' udWIPJV does not exist on the production Vista server (10.112.11.8).
    ' Guard: return Nothing (empty) so the JV tab loads blank rather than erroring.
    Dim chk As ADODB.Recordset
    Set chk = conn.Execute("SELECT COUNT(*) FROM sys.objects WHERE name='udWIPJV' AND type='U'")
    If CLng(chk.Fields(0).Value) = 0 Then
        chk.Close
        Set GetJVDataFromVista = Nothing
        Exit Function
    End If
    chk.Close
    Set chk = Nothing

    Dim sql As String
    Dim cutOffDate As String
    Dim startDate As String
    Dim priorYrEnd As String

    cutOffDate = Format(DateSerial(Year(wipMonth), Month(wipMonth) + 1, 0), "yyyy-MM-dd")
    startDate = Format(DateSerial(Year(wipMonth), 1, 1), "yyyy-MM-dd")
    priorYrEnd = Format(DateSerial(Year(wipMonth) - 1, 12, 31), "yyyy-MM-dd")

    sql = ""
    sql = sql & "DECLARE @Co tinyint = " & co & ";" & vbCrLf
    sql = sql & "DECLARE @CutOffDate date = '" & cutOffDate & "';" & vbCrLf
    sql = sql & "DECLARE @StartDate date = '" & startDate & "';" & vbCrLf
    sql = sql & "DECLARE @PriorYrEnd date = '" & priorYrEnd & "';" & vbCrLf
    sql = sql & "DECLARE @CurrentDate date = GETDATE();" & vbCrLf
    sql = sql & vbCrLf

    ' JV Cost CTE — aggregate costs for JV jobs
    sql = sql & ";WITH JVCosts AS (" & vbCrLf
    sql = sql & "    SELECT d.JCCo, d.Job," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate <= @CutOffDate THEN d.ActualCost ELSE 0 END)" & vbCrLf
    sql = sql & "        + SUM(CASE WHEN d.JCTransType IN ('PR','EM') AND d.ActualDate <= @CutOffDate THEN d.ActualCost ELSE 0 END) AS JTDCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType = 'PF' AND d.PostedDate <= @CutOffDate THEN d.ProjCost ELSE 0 END) AS ProjCost," & vbCrLf
    sql = sql & "        SUM(CASE WHEN d.JCTransType NOT IN ('PR','EM','OE','CO','PF') AND d.PostedDate <= @PriorYrEnd THEN d.ActualCost ELSE 0 END)" & vbCrLf
    sql = sql & "        + SUM(CASE WHEN d.JCTransType IN ('PR','EM') AND d.ActualDate <= @PriorYrEnd THEN d.ActualCost ELSE 0 END) AS PYJTDCost" & vbCrLf
    sql = sql & "    FROM bJCCD d WITH (NOLOCK)" & vbCrLf
    sql = sql & "    WHERE d.JCCo = @Co" & vbCrLf
    sql = sql & "    AND d.Job IN (SELECT IntJobNum FROM udWIPJV WITH (NOLOCK) WHERE Co = @Co)" & vbCrLf
    sql = sql & "    GROUP BY d.JCCo, d.Job" & vbCrLf
    sql = sql & ")" & vbCrLf
    sql = sql & vbCrLf

    ' Final SELECT — field names match what GetudWIPJVSub expects
    sql = sql & "SELECT" & vbCrLf
    sql = sql & "    jv.JVJobNum," & vbCrLf
    sql = sql & "    jv.IntJobNum," & vbCrLf
    sql = sql & "    ISNULL(jv.SupJobNumber, '') AS SupJobNumber," & vbCrLf
    sql = sql & "    jv.JVJobDesc," & vbCrLf
    sql = sql & "    jv.JVPartners," & vbCrLf
    sql = sql & "    jv.OurJVPct," & vbCrLf
    sql = sql & "    0 AS BatchSeq," & vbCrLf

    ' Ops fields
    sql = sql & "    ISNULL(c.ContractAmt, 0) AS OpsContractAmt," & vbCrLf
    sql = sql & "    CASE WHEN ISNULL(jc.ProjCost, 0) = 0 THEN c.ContractAmt" & vbCrLf
    sql = sql & "         ELSE c.ContractAmt END AS OpsProjectedRevenue," & vbCrLf
    sql = sql & "    CASE WHEN ISNULL(jc.ProjCost, 0) = 0 THEN ISNULL(jc.JTDCost, 0)" & vbCrLf
    sql = sql & "         ELSE jc.ProjCost END AS OpsProjectedCost," & vbCrLf
    sql = sql & "    0 AS OpsEarnedRev," & vbCrLf
    sql = sql & "    ISNULL(jc.JTDCost, 0) AS OpsJTDCost," & vbCrLf
    sql = sql & "    ISNULL(c.BilledAmt, 0) AS OpsJTDBillings," & vbCrLf

    ' GAAP fields (same source for Phase 1)
    sql = sql & "    ISNULL(c.ContractAmt, 0) AS GAAPContractAmt," & vbCrLf
    sql = sql & "    0 AS GAAPEarnedRev," & vbCrLf
    sql = sql & "    ISNULL(jc.JTDCost, 0) AS GAAPJTDCost," & vbCrLf
    sql = sql & "    ISNULL(c.BilledAmt, 0) AS GAAPJTDBillings," & vbCrLf
    sql = sql & "    CASE WHEN ISNULL(jc.ProjCost, 0) = 0 THEN 0" & vbCrLf
    sql = sql & "         ELSE c.ContractAmt - jc.ProjCost END AS GAAPProjectedFinalProfit," & vbCrLf

    ' Prior year
    sql = sql & "    0 AS PYOpsEarnedRevenue," & vbCrLf
    sql = sql & "    ISNULL(jc.PYJTDCost, 0) AS PYOpsPJTDCost," & vbCrLf
    sql = sql & "    0 AS PYGAAPEarnedRevenue," & vbCrLf
    sql = sql & "    ISNULL(jc.PYJTDCost, 0) AS PYGAAPJTDCost," & vbCrLf

    ' Completion/workflow defaults (Phase 1)
    sql = sql & "    '' AS OCompleted, '' AS GCompleted," & vbCrLf
    sql = sql & "    '' AS OUserName, '' AS GUserName," & vbCrLf
    sql = sql & "    CAST(0 AS varbinary(8)) AS ORowVersion," & vbCrLf
    sql = sql & "    CAST(0 AS varbinary(8)) AS GRowVersion" & vbCrLf

    sql = sql & "FROM udWIPJV jv WITH (NOLOCK)" & vbCrLf
    sql = sql & "JOIN bJCJM j WITH (NOLOCK) ON j.JCCo = jv.Co AND j.Job = jv.IntJobNum" & vbCrLf
    sql = sql & "JOIN bJCCM c WITH (NOLOCK) ON j.JCCo = c.JCCo AND j.Contract = c.Contract" & vbCrLf
    sql = sql & "LEFT JOIN JVCosts jc ON jc.JCCo = jv.Co AND jc.Job = jv.IntJobNum" & vbCrLf
    sql = sql & "WHERE jv.Co = @Co" & vbCrLf
    sql = sql & "ORDER BY jv.IntJobNum"

    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Open sql, conn

    Set GetJVDataFromVista = rs
    Exit Function

ErrorHandler:
    MsgBox "Error retrieving JV data from Vista:" & vbCrLf & _
           Err.Description, vbCritical, "Vista JV Query Error"
    Set GetJVDataFromVista = Nothing
End Function

Public Function GetWIPDataFromVista(co As Integer, wipMonth As Date, dept As String, groupBy As String) As ADODB.Recordset
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = GetVistaConnection()

    If conn Is Nothing Then
        Set GetWIPDataFromVista = Nothing
        Exit Function
    End If

    Dim sql As String
    sql = BuildWIPQuery(co, wipMonth, dept, groupBy)

    ' Use client-side cursor for bidirectional navigation (matches original behavior)
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Open sql, conn

    Set GetWIPDataFromVista = rs
    Exit Function

ErrorHandler:
    MsgBox "Error retrieving WIP data from Vista:" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "SQL query length: " & Len(sql) & " chars", _
           vbCritical, "Vista Query Error"
    Set GetWIPDataFromVista = Nothing
End Function
