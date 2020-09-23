VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3975
   Icon            =   "Aboutc.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "2000"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   8
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Chat Program Created By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Khoo Jia Jun"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Country Of Origin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Singapore"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   1395
         TabIndex        =   4
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   3
         Top             =   3240
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Bedok South Secondary School (3A)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   2
         Top             =   3480
         Width           =   2625
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   6
         Left            =   1560
         TabIndex        =   1
         Top             =   3840
         Width           =   345
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   45
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form7
End Sub
Private Sub Form_Load()
For i = 0 To 7
Label1(i).Left = (Picture1.Width - Label1(i).Width) / 2
Next i
End Sub
Private Sub Timer1_Timer()
For i = 0 To 7
Label1(i).Top = Label1(i).Top - 10
Next i
If Label1(7).Top = -250 Then
topping = 2040
For i = 0 To 7
Label1(i).Top = topping
If i = 1 Or i = 3 Or i = 5 Then topping = topping + 120
topping = topping + 240
Next i
End If
End Sub

