Attribute VB_Name = "basAlwaysOnTop"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub putMeOnTop(Form As Form)
    SetWindowPos Form.hwnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub


Public Sub takeMeDown(Form As Form)
    SetWindowPos Form.hwnd, -2, 0, 0, 0, 0, 1 Or 2
End Sub
