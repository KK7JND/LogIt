Attribute VB_Name = "basSysTray"
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
  ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_ACTIVATEAPP = &H1C
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64
Public Const GWL_WNDPROC = (-4)

Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * MAX_TOOLTIP
End Type
Public nfIconData As NOTIFYICONDATA
Private FHandle As Long
Private WndProc As Long
Private Hooking As Boolean
Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo Err_WindowProc

    If Hooking = True Then
        If uMsg = WM_RBUTTONUP And lParam = WM_RBUTTONDOWN Then
            frmMain.SysTrayMouseEventHandler
            WindowProc = True
            Exit Function
        End If
        WindowProc = CallWindowProc(WndProc, hw, uMsg, wParam, lParam)
    End If

Exit_WindowProc:
    Exit Function
    
Err_WindowProc:
    MsgBox "Error in basSysTray:WindowProc: " & Err.Description
    Resume Exit_WindowProc

End Function
Public Sub Hook(Lwnd As Long)
On Error GoTo Err_Hook

    If Hooking = False Then
        FHandle = Lwnd
        WndProc = SetWindowLong(Lwnd, GWL_WNDPROC, AddressOf WindowProc)
        Hooking = True
    End If

Exit_Hook:
    Exit Sub
    
Err_Hook:
    MsgBox "Error in basSysTray:Hook: " & Err.Description
    Resume Exit_Hook

End Sub
Public Sub Unhook()
On Error GoTo Err_Unhook

    If Hooking = True Then
        SetWindowLong FHandle, GWL_WNDPROC, WndProc
        Hooking = False
    End If

Exit_Unhook:
    Exit Sub
    
Err_Unhook:
    MsgBox "Error in basSysTray:Unhook: " & Err.Description
    Resume Exit_Unhook

End Sub
Public Sub AddIconToTray(MeHwnd As Long, MeIcon As Long, MeIconHandle As Long, Tip As String)
On Error GoTo Err_AddIconToTray

    With nfIconData
        .hwnd = MeHwnd
        .uID = MeIcon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_RBUTTONUP
        .hIcon = MeIconHandle
        .szTip = Tip & Chr$(0)
        .cbSize = Len(nfIconData)
    End With
    Shell_NotifyIcon NIM_ADD, nfIconData

Exit_AddIconToTray:
    Exit Sub
    
Err_AddIconToTray:
    MsgBox "Error in basSysTray:AddIconToTray: " & Err.Description
    Resume Exit_AddIconToTray

End Sub
Public Sub RemoveIconFromTray()
On Error GoTo Err_RemoveIconFromTray

    Shell_NotifyIcon NIM_DELETE, nfIconData

Exit_RemoveIconFromTray:
    Exit Sub
    
Err_RemoveIconFromTray:
    MsgBox "Error in basSysTray:RemoveIconFromTray: " & Err.Description
    Resume Exit_RemoveIconFromTray

End Sub


