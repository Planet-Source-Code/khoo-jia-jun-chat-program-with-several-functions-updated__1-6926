VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mouse Positions"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4290
   Icon            =   "Mouse.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Mouse.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Send "NoSeeMouse"
End Sub
