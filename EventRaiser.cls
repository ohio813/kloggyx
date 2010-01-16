VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventRaiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta clase es solo para generar eventos en el formulario
'ya que en los modules no pueden generarse xD, ingenio al 100%

Public Event KBHKeyUp(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
Public Event KBHKeyDown(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)

Public Sub RaiseKBHKeyDown(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    RaiseEvent KBHKeyDown(vkCode, scanCode, flags)
End Sub

Public Sub RaiseKBHKeyUp(ByVal vkCode As Integer, ByVal scanCode As Integer, ByVal flags As Integer)
    RaiseEvent KBHKeyUp(vkCode, scanCode, flags)
End Sub
