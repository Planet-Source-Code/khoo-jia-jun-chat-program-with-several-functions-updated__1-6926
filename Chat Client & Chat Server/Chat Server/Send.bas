Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Function StayOnTop(Windows)
SetWindowPos Windows.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
End Function
Public Function Send(Instruct As String)
On Error GoTo skip
Form1.Winsock2.SendData Instruct
skip:
End Function
Public Function ChatSend(Chat As String)
On Error GoTo skip
Form1.Winsock1.SendData Chat
skip:
End Function
