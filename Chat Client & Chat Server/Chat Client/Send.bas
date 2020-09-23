Attribute VB_Name = "Module1"
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

