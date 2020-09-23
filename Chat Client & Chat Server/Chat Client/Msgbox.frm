VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MessageBox Manager"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "00"
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Send"
      Height          =   255
      Index           =   1
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Test"
      Height          =   255
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Ok, Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   3240
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Retry, Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Abort, Retry, Ignore"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Yes, No, Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Yes, No"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Ok, Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Ok"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   3
      Left            =   3240
      Picture         =   "Msgbox.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   2
      Left            =   2400
      Picture         =   "Msgbox.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   1
      Left            =   1560
      Picture         =   "Msgbox.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Height          =   615
      Index           =   0
      Left            =   720
      Picture         =   "Msgbox.frx":24C6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Title:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   345
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Code()
If Check1(0).Value = 1 And Option1(0).Value = True Then Text3 = "00"
If Check1(0).Value = 1 And Option1(1).Value = True Then Text3 = "01"
If Check1(0).Value = 1 And Option1(2).Value = True Then Text3 = "02"
If Check1(0).Value = 1 And Option1(3).Value = True Then Text3 = "03"
If Check1(0).Value = 1 And Option1(4).Value = True Then Text3 = "04"
If Check1(0).Value = 1 And Option1(5).Value = True Then Text3 = "05"
If Check1(0).Value = 1 And Option1(6).Value = True Then Text3 = "06"

If Check1(1).Value = 1 And Option1(0).Value = True Then Text3 = "10"
If Check1(1).Value = 1 And Option1(1).Value = True Then Text3 = "11"
If Check1(1).Value = 1 And Option1(2).Value = True Then Text3 = "12"
If Check1(1).Value = 1 And Option1(3).Value = True Then Text3 = "13"
If Check1(1).Value = 1 And Option1(4).Value = True Then Text3 = "14"
If Check1(1).Value = 1 And Option1(5).Value = True Then Text3 = "15"
If Check1(1).Value = 1 And Option1(6).Value = True Then Text3 = "16"

If Check1(2).Value = 1 And Option1(0).Value = True Then Text3 = "20"
If Check1(2).Value = 1 And Option1(1).Value = True Then Text3 = "21"
If Check1(2).Value = 1 And Option1(2).Value = True Then Text3 = "22"
If Check1(2).Value = 1 And Option1(3).Value = True Then Text3 = "23"
If Check1(2).Value = 1 And Option1(4).Value = True Then Text3 = "24"
If Check1(2).Value = 1 And Option1(5).Value = True Then Text3 = "25"
If Check1(2).Value = 1 And Option1(6).Value = True Then Text3 = "26"

If Check1(3).Value = 1 And Option1(0).Value = True Then Text3 = "30"
If Check1(3).Value = 1 And Option1(1).Value = True Then Text3 = "31"
If Check1(3).Value = 1 And Option1(2).Value = True Then Text3 = "32"
If Check1(3).Value = 1 And Option1(3).Value = True Then Text3 = "33"
If Check1(3).Value = 1 And Option1(4).Value = True Then Text3 = "34"
If Check1(3).Value = 1 And Option1(5).Value = True Then Text3 = "35"
If Check1(3).Value = 1 And Option1(6).Value = True Then Text3 = "36"

If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(0).Value = True Then Text3 = "40"
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(1).Value = True Then Text3 = "41"
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(2).Value = True Then Text3 = "42"
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(3).Value = True Then Text3 = "43"
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(4).Value = True Then Text3 = "44"
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(5).Value = True Then Text3 = "45"
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Option1(6).Value = True Then Text3 = "46"
End Function
Private Sub Check1_Click(Index As Integer)
If Index = 0 And Check1(0).Value = 1 Then Check1(1).Value = 0: Check1(2).Value = 0: Check1(3).Value = 0
If Index = 1 And Check1(1).Value = 1 Then Check1(0).Value = 0: Check1(2).Value = 0: Check1(3).Value = 0
If Index = 2 And Check1(2).Value = 1 Then Check1(0).Value = 0: Check1(1).Value = 0: Check1(3).Value = 0
If Index = 3 And Check1(3).Value = 1 Then Check1(0).Value = 0: Check1(1).Value = 0: Check1(2).Value = 0
Call Code
End Sub
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
If Text3 = "00" Then msg = MsgBox(Text2, vbCritical + vbOKOnly, Text1)
If Text3 = "01" Then msg = MsgBox(Text2, vbCritical + vbOKCancel, Text1)
If Text3 = "02" Then msg = MsgBox(Text2, vbCritical + vbYesNo, Text1)
If Text3 = "03" Then msg = MsgBox(Text2, vbCritical + vbYesNoCancel, Text1)
If Text3 = "04" Then msg = MsgBox(Text2, vbCritical + vbAbortRetryIgnore, Text1)
If Text3 = "05" Then msg = MsgBox(Text2, vbCritical + vbRetryCancel, Text1)
If Text3 = "06" Then msg = MsgBox(Text2, vbCritical + vbMsgBoxHelpButton, Text1)

If Text3 = "10" Then msg = MsgBox(Text2, vbExclamation + vbOKOnly, Text1)
If Text3 = "11" Then msg = MsgBox(Text2, vbExclamation + vbOKCancel, Text1)
If Text3 = "12" Then msg = MsgBox(Text2, vbExclamation + vbYesNo, Text1)
If Text3 = "13" Then msg = MsgBox(Text2, vbExclamation + vbYesNoCancel, Text1)
If Text3 = "14" Then msg = MsgBox(Text2, vbExclamation + vbAbortRetryIgnore, Text1)
If Text3 = "15" Then msg = MsgBox(Text2, vbExclamation + vbRetryCancel, Text1)
If Text3 = "16" Then msg = MsgBox(Text2, vbExclamation + vbMsgBoxHelpButton, Text1)

If Text3 = "20" Then msg = MsgBox(Text2, vbInformation + vbOKOnly, Text1)
If Text3 = "21" Then msg = MsgBox(Text2, vbInformation + vbOKCancel, Text1)
If Text3 = "22" Then msg = MsgBox(Text2, vbInformation + vbYesNo, Text1)
If Text3 = "23" Then msg = MsgBox(Text2, vbInformation + vbYesNoCancel, Text1)
If Text3 = "24" Then msg = MsgBox(Text2, vbInformation + vbAbortRetryIgnore, Text1)
If Text3 = "25" Then msg = MsgBox(Text2, vbInformation + vbRetryCancel, Text1)
If Text3 = "26" Then msg = MsgBox(Text2, vbInformation + vbMsgBoxHelpButton, Text1)

If Text3 = "30" Then msg = MsgBox(Text2, vbQuestion + vbOKOnly, Text1)
If Text3 = "31" Then msg = MsgBox(Text2, vbQuestion + vbOKCancel, Text1)
If Text3 = "32" Then msg = MsgBox(Text2, vbQuestion + vbYesNo, Text1)
If Text3 = "33" Then msg = MsgBox(Text2, vbQuestion + vbYesNoCancel, Text1)
If Text3 = "34" Then msg = MsgBox(Text2, vbQuestion + vbAbortRetryIgnore, Text1)
If Text3 = "35" Then msg = MsgBox(Text2, vbQuestion + vbRetryCancel, Text1)
If Text3 = "36" Then msg = MsgBox(Text2, vbQuestion + vbMsgBoxHelpButton, Text1)

If Text3 = "40" Then msg = MsgBox(Text2, vbOKOnly, Text1)
If Text3 = "41" Then msg = MsgBox(Text2, vbOKCancel, Text1)
If Text3 = "42" Then msg = MsgBox(Text2, vbYesNo, Text1)
If Text3 = "43" Then msg = MsgBox(Text2, vbYesNoCancel, Text1)
If Text3 = "44" Then msg = MsgBox(Text2, vbAbortRetryIgnore, Text1)
If Text3 = "45" Then msg = MsgBox(Text2, vbRetryCancel, Text1)
If Text3 = "46" Then msg = MsgBox(Text2, vbMsgBoxHelpButton, Text1)
End If
If Index = 1 Then Send "MsgBox" + Text3 + Text1 + Chr$(255) + Text2
End Sub
Private Sub Option1_Click(Index As Integer)
Call Code
End Sub
