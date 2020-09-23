VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JiaJun Chat Application - Client Mode"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6945
   Icon            =   "Chat Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Connect"
      Height          =   255
      Index           =   8
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&About"
      Height          =   255
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Clear"
      Height          =   255
      Index           =   6
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&See Mouse"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Process Manager"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&File Manager"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Function"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Attention"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Invisible Server"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   6360
      Width           =   1305
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3240
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Server To Client: Not Connected"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   17
      Top             =   6960
      Width           =   2310
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Client To Server: Not Connected"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   16
      Top             =   6720
      Width           =   2310
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "My IP: "
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4980
      TabIndex        =   15
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Connect To:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4560
      TabIndex        =   5
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = 0 And Command1(0).Caption = "&Invisible Server" Then Send "InvisibleF": Command1(0).Caption = "&Visible Server": Exit Sub
If Index = 0 And Command1(0).Caption = "&Visible Server" Then Send "VisibleF": Command1(0).Caption = "&Invisible Server": Exit Sub
If Index = 1 Then Send "Attention"
If Index = 2 Then Form2.Visible = True
If Index = 3 Then Form3.Visible = True
If Index = 4 Then Form4.Visible = True
If Index = 5 Then Form5.Visible = True: Send "SeeMouse"
If Index = 6 Then Text1 = ""
If Index = 7 Then Form7.Timer1.Enabled = True: Form7.Visible = True

If Index = 8 And Command1(8).Caption = "&Connect" Then
If Text3 = "" Then msg = MsgBox("Please Enter A Valid IP", vbCritical, "Error"): Exit Sub
Command1(8).Caption = "&Disconnect"
Winsock1.RemoteHost = Text3
Winsock2.RemoteHost = Text3
Send "RequestConnect"
Exit Sub
End If

If Index = 8 And Command1(8).Caption = "&Disconnect" Then
Command1(8).Caption = "&Connect"
Label6 = "Client To Server: Not Connected"
Text1.Locked = True
For enable = 0 To 5
Command1(enable).Enabled = False
Next enable
Winsock1.Close
Winsock2.Close
Send "Disconnect"
Exit Sub
End If

End Sub
Private Sub Form_Load()
Label3 = Winsock1.LocalIP
With Winsock1
.RemotePort = 1002
.Bind 1001
End With
With Winsock2
.RemotePort = 1004
.Bind 1003
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Label6 = "Client To Server: Not Connected" Then End
Send "Disconnect"
End
End Sub
Private Sub Text1_Change()
ChatSend Text1.Text
End Sub
Private Sub text2_Change()
On Error Resume Next
Text2.SelLength = 0
If Len(Text2.Text) > 0 Then
If Right$(Text2.Text, 1) = vbCrLf Then
Text2.SelStart = Len(Text2.Text) - 1
Exit Sub
End If
Text2.SelStart = Len(Text2.Text)
End If
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Receive Text
Dim Ferry As String
Winsock1.GetData Ferry
Text2 = Ferry
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
'Receive File Info
Dim FileInfo As String
Winsock2.GetData FileInfo
If FileInfo = "MsgboxOk" Then msg = MsgBox("User Clicked Ok", vbInformation, "")
If FileInfo = "MsgboxCancel" Then msg = MsgBox("User Clicked Cancel", vbInformation, "")
If FileInfo = "MsgboxAbort" Then msg = MsgBox("User Clicked Abort", vbInformation, "")
If FileInfo = "MsgboxRetry" Then msg = MsgBox("User Clicked Retry", vbInformation, "")
If FileInfo = "MsgboxIgnore" Then msg = MsgBox("User Clicked Ignore", vbInformation, "")
If FileInfo = "MsgboxYes" Then msg = MsgBox("User Clicked Yes", vbInformation, "")
If FileInfo = "MsgboxNo" Then msg = MsgBox("User Clicked No", vbInformation, "")

If FileInfo = "Accepted" Then
Label6 = "Client To Server: Connected"
Text1.Locked = False
For enable = 0 To 5
Command1(enable).Enabled = True
Next enable
End If
If FileInfo = "RequestConnect" Then Label7 = "Server To Client: Connected": Send "Accepted"
If FileInfo = "Disconnect" Then Label7 = "Server To Client: Not Connected"
If InStr(1, FileInfo, "RunList") <> 0 Then Form4.Text1 = Right$(FileInfo, Len(FileInfo) - 7)
If InStr(1, FileInfo, "Login") <> 0 Then msg = MsgBox(Right$(FileInfo, Len(FileInfo) - 5), vbInformation, "")
If InStr(FileInfo, "Dir") Then Form3.Text1 = Right$(FileInfo, Len(FileInfo) - 3)
If InStr(FileInfo, "File") Then Form3.Text2 = Right$(FileInfo, Len(FileInfo) - 4)
If InStr(FileInfo, "Success") Then Form3.List2.RemoveItem (Form3.List2.ListIndex)
If InStr(1, FileInfo, "MX") <> 0 Then Form5.Image1.Left = Right$(FileInfo, Len(FileInfo) - 2)
If InStr(1, FileInfo, "MY") <> 0 Then Form5.Image1.Top = Right$(FileInfo, Len(FileInfo) - 2)
End Sub
