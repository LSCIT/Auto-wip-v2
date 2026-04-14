Attribute VB_Name = "Permissions"
' Permissions — Security role and department permission checks
' Modified: March 2026 — Phase 1 Vista Direct Connection
' Modified: April 2026  — C10: Real pnp.WIPSECGetRole lookup
' Changes:
'   - GetSecurity: Phase 1 hardcoded WIPAccounting → C10 live lookup via pnp.WIPSECGetRole
'   - GetDept: kept as no-op (departments line was commented out in original — never enforced)
'   - Bug C4 FIX: uses MSOLEDBSQL provider (no Workstation ID=MROBERTS)


' ========================================
' PNP_SECURITY_DB — name of the PNP4 production database on PPServerName.
' pnp.WIPSECGetRole and pnp.WipPermissionsView2 live in this database.
' If the actual database name differs from "PnpMain", update this constant.
' ========================================
Private Const PNP_SECURITY_DB As String = "PnpMain"


' ========================================
' GetSecurity — C10 Implementation
' Connects to PPServerName / PNP_SECURITY_DB and calls pnp.WIPSECGetRole.
' Stores the returned role in Sheet2.Range("Role").
'
' Role values returned by pnp.WIPSECGetRole:
'   WIPAccounting       — full cycle: load, ready-for-ops, GAAP edit, final approval
'   WIPLevel2           — ops cycle: load, edit Jobs-Ops, ops final approval
'   WipInitialApproval  — ops edit only (Done column)
'   WipFinalApproval    — ops final approval only
'   WipViewOnly         — read-only access
'
' If the user is not in pnp.WipPermissionsView2, the proc returns
' @RetMsg = "Security Settings Not Valid" and Role is left blank.
' Josh/Dane must ensure all users are in the security table before deploying.
' ========================================
Public Sub GetSecurity(User As String)
    On Error GoTo errexit

    Sheet2.Range("Role").Value       = ""
    Sheet2.Range("Companies").Value  = ""
    Sheet2.Range("Departments").Value = ""

    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command

    ' Connect to PNP security database using P&P server + credentials
    cnn.connectionString = "Provider=MSOLEDBSQL;" & _
                           "Server="    & CStr(Sheet2.Range("PPServerName").Value) & ";" & _
                           "Database="  & PNP_SECURITY_DB & ";" & _
                           "UID="       & CStr(Sheet2.Range("PPUsername").Value)  & ";" & _
                           "PWD="       & CStr(Sheet2.Range("PPPassword").Value)  & ";" & _
                           "TrustServerCertificate=yes;Encrypt=no;"
    cnn.CommandTimeout = 10
    cnn.Open

    cmd.ActiveConnection = cnn
    cmd.CommandType      = adCmdStoredProc
    cmd.CommandText      = "pnp.WIPSECGetRole"
    cmd.CommandTimeout   = 30

    Dim cmdUser     As ADODB.Parameter
    Dim cmdRoleName As ADODB.Parameter
    Dim cmdRetMsg   As ADODB.Parameter
    Set cmdUser     = cmd.CreateParameter("@user",     adVarChar, adParamInput,  30,  User)
    Set cmdRoleName = cmd.CreateParameter("@roleName", adVarChar, adParamOutput, 30)
    Set cmdRetMsg   = cmd.CreateParameter("@RetMsg",   adVarChar, adParamOutput, 500)
    cmd.Parameters.Append cmdUser
    cmd.Parameters.Append cmdRoleName
    cmd.Parameters.Append cmdRetMsg

    cmd.Execute

    Dim roleName As String
    Dim retMsg   As String
    roleName = IIf(IsNull(cmd.Parameters("@roleName").Value), "", cmd.Parameters("@roleName").Value)
    retMsg   = IIf(IsNull(cmd.Parameters("@RetMsg").Value),   "", cmd.Parameters("@RetMsg").Value)

    cnn.Close
    Set cnn = Nothing
    Set cmd = Nothing

    If retMsg <> "" Then
        MsgBox retMsg, vbInformation, "WIP Security"
        Exit Sub
    End If

    Sheet2.Range("Role").Value = roleName
    GoTo 9999

errexit:
    If Not cnn Is Nothing Then
        If cnn.State <> 0 Then cnn.Close
    End If
    MsgBox "There was an error in the GetSecurity Routine. " & Err.Description, vbOKOnly
9999:
End Sub


' ========================================
' GetDept — kept as no-op
' The original called pnp.WIPSECGetDepartments but had a bug:
' the result line was commented out ('Departments = cmd.Parameters(...).Value)
' so departments were NEVER stored or enforced in any version of this tool.
' Implementing real department filtering is deferred to a future sprint.
' ========================================
Public Sub GetDept(User As String, Company As Integer)
    Sheet2.Range("Departments").Value = ""
End Sub
