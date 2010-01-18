Attribute VB_Name = "Rutinas"
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
                         ByVal lpString As String, ByVal cch As Long) As Long

Public UploadUrl As String

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

Public Function UploadString(ByVal DatosASubir As String, NombreArchivo) As String
    Dim HttpRequest As String
    Dim Cuerpo As String
    Dim Identificador As String
    Dim HttpLength As Long
    Dim Host As String
    Dim UploadPath As String
    Dim TmpUploadURL As String
    Dim tmpPos As Integer
    
    TmpUploadURL = UploadUrl
    TmpUploadURL = Replace(TmpUploadURL, "http://", "")
    tmpPos = InStr(1, TmpUploadURL, "/")
    
    If tmpPos > 0 Then
        Host = Mid(TmpUploadURL, 1, tmpPos - 1)
        UploadPath = Replace(TmpUploadURL, Host, "")
    Else
    End If
    
    HttpLength = Len(strBody)
    ' Es necesario generar un random string de 32 caracteres alphanumericos
    ' para establecerlo como frontera, inicio y fin, del contenido del los datos a enviar
    Dim RandomStr As String, j As Integer, Charz As Byte, Condicion As Boolean
    
    Randomize
    
    For j = 1 To 32
        Charz = Int(Rnd() * 127) 'numero aleatoreo entre 1 y 127, 127 es el numero de caracteres ascii estandar que existen
        Condicion = Condicion And (Charz >= Asc("A") And Charz <= Asc("Z")) 'letras A...Z
        Condicion = Condicion And (Charz >= Asc("a") And Charz <= Asc("z")) 'letras a...z
        Condicion = Condicion And (Charz >= Asc("0") And Charz <= Asc("9")) 'numeros 0...9
        If Condicion Then
            RandomStr = RandomStr & Chr(Charz)
        Else 'si no fue caracter alphanumerico, lo intentamos de nuevo
            j = j - 1
        End If
    Next j
    
    Identificador = RandomStr
    
    Cuerpo = "--" & Identificador & vbCrLf
    Cuerpo = Cuerpo & "Content-Disposition: form-data; name=""abupload"";"
    Cuerpo = Cuerpo & " filename=""" & NombreArchivo & """" & vbCrLf
    Cuerpo = Cuerpo & "Content-Type: text/plain" & vbCrLf
    Cuerpo = Cuerpo & vbCrLf & DatosASubir
    Cuerpo = Cuerpo & vbCrLf & "--" & Identificador & "--"
    HttpRequest = "POST " & UploadPath & " HTTP/1.0" & vbCrLf
    HttpRequest = HttpRequest & "Host: " & Host & vbCrLf
    HttpRequest = HttpRequest & "Content-Type: multipart/form-data, boundary=" & Identificador & vbCrLf
    HttpRequest = HttpRequest & "Content-Length: " & HttpLength & vbCrLf & vbCrLf
    HttpRequest = HttpRequest & Cuerpo

    BuildFileUploadRequest = HttpRequest
End Function

Public Sub NuevoError(ByVal ErrStr As String)
    On Error Resume Next 'Otro LOL!!! haha
    FrmMain.MyEventRaiser.RaiseErrorDetected ErrStr
End Sub
