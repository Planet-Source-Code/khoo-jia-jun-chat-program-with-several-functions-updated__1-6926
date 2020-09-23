VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Manager"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   Icon            =   "File.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Exit"
      Height          =   255
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Delete File"
      Height          =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.ListBox List4 
      Height          =   255
      ItemData        =   "File.frx":000C
      Left            =   840
      List            =   "File.frx":0019
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "File.frx":0039
      Left            =   120
      List            =   "File.frx":008B
      TabIndex        =   5
      Top             =   4200
      Width           =   615
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3570
      ItemData        =   "File.frx":00DD
      Left            =   3960
      List            =   "File.frx":00DF
      TabIndex        =   4
      Top             =   480
      Width           =   3855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3570
      ItemData        =   "File.frx":00E1
      Left            =   120
      List            =   "File.frx":00E3
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Run:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   10
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Delete A File
msg = MsgBox("Are you Sure?", vbYesNo + vbQuestion, "Confirm Delete")
If msg = vbYes And Len(Label1.Caption) <= 3 Then Send "DelFile" + Label1.Caption + List2.List(List2.ListIndex): Exit Sub
If msg = vbYes And Len(Label1.Caption) > 3 Then Send "DelFile" + Label1.Caption + "\" + List2.List(List2.ListIndex): Exit Sub
If msg = vbNo Then Exit Sub
End Sub
Private Sub Command2_Click()
Unload Form3
End Sub
Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
'Delete A File
If KeyCode = vbKeyDelete Then
msg = MsgBox("Are you Sure?", vbYesNo + vbQuestion, "Confirm Delete")
If msg = vbYes And Len(Label1.Caption) <= 3 Then Send "DelFile" + Label1.Caption + List2.List(List2.ListIndex): Exit Sub
If msg = vbYes And Len(Label1.Caption) > 3 Then Send "DelFile" + Label1.Caption + "\" + List2.List(List2.ListIndex): Exit Sub
If msg = vbNo Then Exit Sub
End If
End Sub
Private Sub List3_Click()
Label1 = List3.List(List3.ListIndex) + ":\"
Send "DirPath" + List3.List(List3.ListIndex) + ":\"
End Sub
Private Sub List4_Click()
If List4.ListIndex = 0 Then Send "DirPath" + "C:\Windows\Desktop": Label1.Caption = "C:\Windows\Desktop"
If List4.ListIndex = 1 Then Send "DirPath" + "C:\My Documents": Label1.Caption = "C:\My Documents"
If List4.ListIndex = 2 Then Send "DirPath" + "C:\Program Files\ICQ": Label1.Caption = "C:\Program Files\ICQ"
End Sub
Private Sub Text1_Change()
List1.Clear
On Error GoTo skip
Foundpos = 1
For i = 0 To Len(Text1) / 3
Foundpos1 = InStr(Foundpos, Text1, Chr$(13))
List1.AddItem Mid(Text1, Foundpos, Foundpos1 - Foundpos), i
Foundpos = Foundpos + (Foundpos1 - Foundpos) + 1
Next
skip:
End Sub
Private Sub text2_Change()
List2.Clear
On Error GoTo skip
Foundpos = 1
For i = 0 To Len(Text2) / 3
Foundpos1 = InStr(Foundpos, Text2, Chr$(13))
List2.AddItem Mid(Text2, Foundpos, Foundpos1 - Foundpos), i
Foundpos = Foundpos + (Foundpos1 - Foundpos) + 1
Next
skip:
End Sub
Private Sub List1_DblClick()
Label1.Caption = List1.List(List1.ListIndex)
Send "DirPath" + List1.List(List1.ListIndex)
End Sub
Private Sub list2_DblClick()
If Len(Label1) <= 3 Then Send "Execute" + Label1 + List2.List(List2.ListIndex)
If Len(Label1) > 3 Then Send "Execute" + Label1 + "\" + List2.List(List2.ListIndex)
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If Text3 = "" Then Exit Sub
Send "Execute" + Text3
Text3 = ""
End If
End Sub
