Attribute VB_Name = "Rutinas"
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
                         ByVal lpString As String, ByVal cch As Long) As Long
       
Public Function ActiveWindow() As String
    Dim hwndActivo As Long
    Dim Buff As String * 255
    Dim Title As String
    'Obtener el Handle de la aplicacion activa
    hwndActivo = GetForegroundWindow
    If hwndActivo = 0 Then Exit Function
    Title = Left$(Buff, GetWindowText(hwndActivo, ByVal Buff, 255))
    ActiveWindow = Title
End Function

