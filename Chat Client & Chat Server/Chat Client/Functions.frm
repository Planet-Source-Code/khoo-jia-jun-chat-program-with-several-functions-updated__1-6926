VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Functions"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   Icon            =   "Functions.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Off"
      Height          =   255
      Index           =   18
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Clock"
      Height          =   255
      Index           =   17
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Tray"
      Height          =   255
      Index           =   16
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Tabs"
      Height          =   255
      Index           =   15
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Icons"
      Height          =   255
      Index           =   14
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Button"
      Height          =   255
      Index           =   13
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide"
      Height          =   255
      Index           =   12
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide"
      Height          =   255
      Index           =   11
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Swap"
      Height          =   255
      Index           =   10
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Confine"
      Height          =   255
      Index           =   9
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Freeze"
      Height          =   255
      Index           =   8
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "MessageBox"
      Height          =   255
      Index           =   7
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Close CG"
      Height          =   255
      Index           =   5
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clear Text"
      Height          =   255
      Index           =   4
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Start Url"
      Height          =   255
      Index           =   3
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete Run + RunServices"
      Height          =   255
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add RunServices"
      Height          =   255
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Run"
      Height          =   255
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Monitor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   22
      Top             =   2400
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Taskbar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Desktop:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Misc:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Registry"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Mouse:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = "0" Then AddRegName$ = InputBox("Enter String Value Name", ""): Send "RegNameRun" + AddRegName$: AddRegData$ = InputBox("Enter String Value Data", ""): Send "RegDataRun" + AddRegData$ 'Add Run
If Index = "1" Then AddRegName$ = InputBox("Enter String Value Name", ""): Send "RegNameRunSerVices" + AddRegName$: AddRegData$ = InputBox("Enter String Value Data", ""): Send "RegDataRunSerVices" + AddRegData$ 'Add RunServices
If Index = "2" Then Send "RegDel": msg = MsgBox("Registry Startup Contents Deleted!!!", vbInformation, "") 'Delete Run + RunServices
If Index = "3" Then URL$ = InputBox("Enter URL to Start", ""): Send "URL" + URL$ 'Start URL
If Index = "4" Then Send "CText" 'Clear Text
If Index = "5" Then Send "CloseAPPic": Send "CloseAPPcgmenu": Send "CloseAPPcg16eh": Send "CloseAPPcg32eh"  'Close Norton CrashGuard"
If Index = "7" Then Form6.Visible = True 'MessageBox
If Index = "8" And Command1(8).Caption = "Freeze" Then Send "JMouse": Command1(8).Caption = "Free": Exit Sub 'Freeze Mouse
If Index = "8" And Command1(8).Caption = "Free" Then Send "FMouse": Command1(8).Caption = "Freeze": Exit Sub 'Free Mouse
If Index = "9" And Command1(9).Caption = "Confine" Then Send "CMouse": Command1(9).Caption = "Free": Exit Sub 'Confine Mouse"
If Index = "9" And Command1(9).Caption = "Free" Then Send "FMouse": Command1(9).Caption = "Confine": Exit Sub 'Free Mouse
If Index = "10" And Command1(10).Caption = "Swap" Then Send "SMouse": Command1(10).Caption = "Restore": Exit Sub 'Swap Mouse Buttons
If Index = "10" And Command1(10).Caption = "Restore" Then Send "RMouse": Command1(10).Caption = "Swap": Exit Sub 'Restore Mouse Buttons

If Index = "11" And Command1(11).Caption = "Hide" Then Send "HideD": Command1(11).Caption = "Show": Exit Sub 'Hide Desktop
If Index = "11" And Command1(11).Caption = "Show" Then Send "ShowD": Command1(11).Caption = "Hide": Exit Sub 'Show Desktop
If Index = "12" And Command1(12).Caption = "Hide" Then Send "HideT": Command1(12).Caption = "Show": Exit Sub 'Hide Taskbar
If Index = "12" And Command1(12).Caption = "Show" Then Send "ShowT": Command1(12).Caption = "Hide": Exit Sub 'Show Taskbar
If Index = "13" And Command1(13).Caption = "Hide Button" Then Send "HideTB": Command1(13).Caption = "Show Button": Exit Sub 'Hide Taskbar Button
If Index = "13" And Command1(13).Caption = "Show Button" Then Send "ShowTB": Command1(13).Caption = "Hide Button": Exit Sub 'Show Taskbar Button
If Index = "14" And Command1(14).Caption = "Hide Icons" Then Send "HideTI": Command1(14).Caption = "Show Icons": Exit Sub 'Hide Taskbar Icons
If Index = "14" And Command1(14).Caption = "Show Icons" Then Send "ShowTI": Command1(14).Caption = "Hide Icons": Exit Sub 'Show Taskbar Icons
If Index = "15" And Command1(15).Caption = "Hide Tabs" Then Send "HideTTabs": Command1(15).Caption = "Show Tabs": Exit Sub 'Hide Taskbar Tabs
If Index = "15" And Command1(15).Caption = "Show Tabs" Then Send "ShowTTabs": Command1(15).Caption = "Hide Tabs": Exit Sub 'Show Taskbar Tabs
If Index = "16" And Command1(16).Caption = "Hide Tray" Then Send "HideTTray": Command1(16).Caption = "Show Tray": Exit Sub 'Hide Taskbar Tray
If Index = "16" And Command1(16).Caption = "Show Tray" Then Send "ShowTTray": Command1(16).Caption = "Hide Tray": Exit Sub 'Show Taskbar Tray
If Index = "17" And Command1(17).Caption = "Hide Clock" Then Send "HideTClock": Command1(17).Caption = "Show Clock": Exit Sub 'Hide Taskbar Clock
If Index = "17" And Command1(17).Caption = "Show Clock" Then Send "ShowTClock": Command1(17).Caption = "Hide Clock": Exit Sub 'Show Taskbar Clock

If Index = "18" And Command1(18).Caption = "Off" Then Send "MonitorOff": Command1(18).Caption = "On": Exit Sub 'Off Monitor
If Index = "18" And Command1(18).Caption = "On" Then Send "MonitorOn": Command1(18).Caption = "Off": Exit Sub 'On Monitor

End Sub
