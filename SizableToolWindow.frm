VERSION 5.00
Begin VB.Form frmSizableToolWindow 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Command Button Example"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTitleBarButton 
      Caption         =   "Fake Command Button"
      Height          =   735
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoveButton 
      Caption         =   "Remove TitleBar Button"
      Height          =   375
      Left            =   2460
      TabIndex        =   3
      Top             =   960
      Width           =   2115
   End
   Begin VB.CommandButton cmdAddButton 
      Caption         =   "Add TitleBar Button"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   480
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This time we use the default caption button size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmSizableToolWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const USE_HOOK As Boolean = True
Private cBtnTitleBar As clsBtnTitleBar

Private Sub cmdAddButton_Click()

    If USE_HOOK Then
        Set cBtnTitleBar = Nothing
        Set cBtnTitleBar = New clsBtnTitleBar
        cBtnTitleBar.AddButton Me, False, Me.cmdTitleBarButton.hWnd, "?", 0, True
    End If

End Sub

Private Sub cmdRemoveButton_Click()
Set cBtnTitleBar = Nothing
End Sub

Private Sub cmdTitleBarButton_Click()
MsgBox "The Titlebar button was clicked", vbInformation, "Command button clicked"
End Sub

Private Sub Form_Load()

    Me.Left = (Screen.Width - Me.Width) * 2 / 3
    Me.Top = (Screen.Height - Me.Height) * 2 / 3

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set cBtnTitleBar = Nothing
End Sub

