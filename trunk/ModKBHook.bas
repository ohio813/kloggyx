Attribute VB_Name = "ModKBHook"
Option Explicit
Public hKbdHook As Long
Private Const WH_KEYBOARD_LL As Integer = 13
Private Const HC_ACTION As Integer = 0
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_CHAR As Long = &H102

Private Type KBDLLHOOKSTRUCT
    vkCode As Integer
    scanCode As Integer
    flags As Integer
    time As Integer
    dwExtraInfo As Integer
End Type

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_PRESSED As Integer = &H1000

Public Enum OtrosVK
        VK_LControl = &HA2
        VK_RControl = &HA3
        VK_Delete = &H2E
        VK_LShift = &HA0
        VK_RShift = &HA1
        VK_Pause = &H13
        VK_PrintScreen = 44
        VK_LWin = &H5B
        VK_RWin = &H5C
        VK_Alt = &H12
        VK_LAlt = &HA4
        VK_RAlt = &HA5
End Enum

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Integer
    Dim kbdllhs As KBDLLHOOKSTRUCT
    On Error GoTo ErrHandler
    CopyMemory kbdllhs, ByVal lParam, Len(kbdllhs)
    If nCode = HC_ACTION Then
        LowLevelKeyboardProc = CallNextHookEx(hKbdHook, nCode, wParam, lParam)
        Select Case wParam
            Case WM_KEYDOWN, WM_SYSKEYDOWN
               FrmMain.MyEventRaiser.RaiseKBHKeyDown kbdllhs.vkCode, kbdllhs.scanCode, kbdllhs.flags
            Case WM_KEYUP, WM_SYSKEYUP
               FrmMain.MyEventRaiser.RaiseKBHKeyUp kbdllhs.vkCode, kbdllhs.scanCode, kbdllhs.flags
        End Select
    Else: LowLevelKeyboardProc = CallNextHookEx(hKbdHook, nCode, wParam, lParam)
    End If
    Exit Function
ErrHandler:
    FrmMain.MyEventRaiser.RaiseErrorDetected "Error en LowLevelKeyBoardProc"
    Err.Clear
End Function

Public Sub Hook()
    On Error Resume Next '<--- LOL
    hKbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
    If hKbdHook = 0 Then
        FrmMain.MyEventRaiser.RaiseErrorDetected "     Error al tratar de hacer el Hook"
    End If
End Sub

Public Sub UnHook()
    On Error Resume Next '<--- LOL
    Call UnhookWindowsHookEx(hKbdHook)
End Sub

