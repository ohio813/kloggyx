Attribute VB_Name = "Module1"
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
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByRef lpbKeyState As Byte, ByRef lpwTransKey As Integer, ByVal fuState As Integer) As Integer

Private Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Integer
    Dim kbdllhs As KBDLLHOOKSTRUCT
    CopyMemory kbdllhs, ByVal lParam, Len(kbdllhs)
    If nCode = HC_ACTION Then
        LowLevelKeyboardProc = CallNextHookEx(hKbdHook, nCode, wParam, lParam)
        Select Case wParam
            Case WM_KEYDOWN, WM_SYSKEYDOWN
               Form1.MyEventRaiser.RaiseKBHKeyDown kbdllhs.vkCode, kbdllhs.scanCode, kbdllhs.flags
            Case WM_KEYUP, WM_SYSKEYUP
               Form1.MyEventRaiser.RaiseKBHKeyUp kbdllhs.vkCode, kbdllhs.scanCode, kbdllhs.flags
        End Select
    Else: LowLevelKeyboardProc = CallNextHookEx(hKbdHook, nCode, wParam, lParam)
    End If
End Function

Public Sub Hook()
    hKbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
    If hKbdHook = 0 Then
        Open "C:\status.akl" For Append As #1
            Print #1, "     Error al tratar de hacer el Hook"
        Close #1
        Exit Sub
    End If
End Sub

Public Sub UnHook()
    Call UnhookWindowsHookEx(hKbdHook)
End Sub

