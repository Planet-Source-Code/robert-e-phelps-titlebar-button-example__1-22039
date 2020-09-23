VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TitleBar Button Example Menu"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   375
      Left            =   3180
      TabIndex        =   2
      Top             =   1380
      Width           =   915
   End
   Begin VB.CommandButton cmdSizableToolWindow 
      Caption         =   "Sizeable ToolWindow"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1755
   End
   Begin VB.CommandButton cmdSizable 
      Caption         =   "Sizable Form"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1755
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()

    Unload frmSizable
    Unload frmSizableToolWindow
    Unload Me

End Sub

Private Sub cmdSizable_Click()
frmSizable.Show
End Sub

Private Sub cmdSizableToolWindow_Click()
frmSizableToolWindow.Show
End Sub

Private Sub Form_Load()

    ' Position Menu form
    Me.Left = (Screen.Width - Me.Width) / 4
    Me.Top = (Screen.Height - Me.Height) / 2

End Sub
