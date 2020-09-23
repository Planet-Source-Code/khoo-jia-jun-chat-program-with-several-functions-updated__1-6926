VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Process Manager"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   Icon            =   "Process.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Close"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Refresh"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Send "CloseAPP°°°°°°"
End Sub
Private Sub Command2_Click()
Unload Form4
End Sub
Private Sub Form_Load()
Send "CloseAPP°°°°°°"
End Sub
Private Sub List1_DblClick()
Send "CloseAPP" + List1.List(List1.ListIndex)
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
