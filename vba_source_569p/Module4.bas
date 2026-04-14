Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function MoveWindow Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal x As LongPtr, ByVal y As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr, ByVal bRepaint As LongPtr) As LongPtr

Const MESSAGE_TITLE As String = "My Message Box"

Sub ShowAndMoveMsgBox()
    Dim hWnd As LongPtr
    Dim RetVal As LongPtr
    Dim frm As DataRetrievalStatus
    Set frm = New DataRetrievalStatus
    
    frm.Label1.Caption = "Getting Data......."
    
    
    frm.StartUpPosition = 0
    ' Calculate the position where the UserForm should appear to be centered in Excel
    frm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frm.Width)
    frm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frm.Height)
    
    
    
    
    frm.Show vbModeless
    
    

    ' Small delay to ensure the MsgBox has time to appear
    Application.Wait (Now + TimeValue("0:00:01"))

End Sub
