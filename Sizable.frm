VERSION 5.00
Begin VB.Form frmSizable 
   Caption         =   "Toggle Button Example"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTitleBarButton 
      Caption         =   "Fake TitleBarButton"
      Height          =   555
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdRemoveButton 
      Caption         =   "Remove Button"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   540
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddButton 
      Caption         =   "Add Button"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Remove button on TitleBar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "frmSizable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const USE_HOOK As Boolean = True
Private cBtnTitleBar As clsBtnTitleBar

Private Sub chkTitleBarButton_Click()

    If Me.chkTitleBarButton.Value = vbChecked Then
        MsgBox "The button is ON" & vbCrLf & vbCrLf & "This is where code " & _
               "can be placed to execute", vbInformation, "Toggle button"
    Else
        MsgBox "The button is OFF" & vbCrLf & vbCrLf & "This is where code " & _
               "can be placed to execute", vbInformation, "Toggle button"
    End If

End Sub

Private Sub cmdAddButton_Click()

    If USE_HOOK Then
        Set cBtnTitleBar = Nothing
        Set cBtnTitleBar = New clsBtnTitleBar
        cBtnTitleBar.AddButton Me, True, Me.chkTitleBarButton.hWnd, "Show Description", 100, False
    End If

End Sub

Private Sub cmdRemoveButton_Click()
Set cBtnTitleBar = Nothing
End Sub

Private Sub Form_Load()

    Me.Left = (Screen.Width - Me.Width) * 2 / 3
    Me.Top = (Screen.Height - Me.Height) / 3

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Always destroy class to end sub-classing or app will crash!
    Set cBtnTitleBar = Nothing

End Sub
