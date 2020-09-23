VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBtnTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************
' Developer:  Robert E. Phelps
' Feel free to use this code as you wish, but please give credit to the
' author.  Do not sell this code; if you do, I want a piece of it :) !!!
' ************************************************************************

Private hWndButton As Long
Private hWndMain As Long

Public Sub AddButton(ByRef frm As Form, ByVal bToggleButton As Boolean, _
                     ByVal hWndFakeButton As Long, ByVal sText As String, _
                     ByVal lWidthInPixels As Long, ByVal bCaptionButtonStyle As Boolean)

    Dim udtNCM As NONCLIENTMETRICS
    Dim udtLogFontCap As LOGFONT
    Dim hFont As Long
    Dim hWndParent As Long
    Dim lpWndProc As Long


    ' Validation
    If frm.BorderStyle = 0 Then Exit Sub
    If hWndFakeButton = 0 Then Exit Sub
    If bCaptionButtonStyle Then
        ' Use default caption button width
        lWidthInPixels = 0
    Else
        If lWidthInPixels <= 0 Then Exit Sub
    End If

    ' The button will be a free-floating window that has no parent (0)
    hWndMain = frm.hWnd
    hWndParent = 0

    ' Create button, remove window border, and set parent to none
    If Not bToggleButton Then
        ' Command button
        hWndButton = CreateWindowEx(0, "Button", sText, BS_PUSHBUTTON, 20, 20, 20, 20, hWndMain, 0, App.hInstance, 0)
    Else
        ' Toggle button
        hWndButton = CreateWindowEx(0, "Button", sText, BS_PUSHLIKE + BS_AUTOCHECKBOX, 20, 20, 20, 20, hWndMain, 0, App.hInstance, 0)
    End If
    Call SetWindowLong(hWndButton, GWL_STYLE, GetWindowLong(hWndButton, GWL_STYLE) - WS_CAPTION)
    Call SetParent(hWndButton, hWndParent)

    ' Get Non-Client Metrics
    udtNCM.cbSize = Len(udtNCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, udtNCM, 0)

    ' Select caption font for BorderStyle
    Select Case frm.BorderStyle
        Case 1, 2, 3: udtLogFontCap = udtNCM.lfCaptionFont
        Case 4, 5: udtLogFontCap = udtNCM.lfSMCaptionFont
    End Select

    ' Use absolute value for font height
    udtLogFontCap.lfHeight = Abs(udtLogFontCap.lfHeight)
    If WS_EX_TOOLWINDOW = (GetWindowLong(hWndMain, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) Then
        ' Reduce font size for ToolWindow
        udtLogFontCap.lfHeight = udtLogFontCap.lfHeight - 1
    End If

    ' Get handle to retrieved caption font and set font to button
    hFont = CreateFontIndirect(udtLogFontCap)
    Call SendMessage(hWndButton, WM_SETFONT, hFont, ByVal True)

    ' Sub-class Main form and set window properties for callback function
    lpWndProc = SetWindowLong(hWndMain, GWL_WNDPROC, AddressOf WindowProc)
    Call SetProp(hWndMain, HWND_PROP_lpPrevWndProc, lpWndProc)
    Call SetProp(hWndMain, HWND_PROP_hWndMainForm, hWndMain)
    Call SetProp(hWndMain, HWND_PROP_hWndTitleBarButton, hWndButton)

    ' Sub-class Titlebar button and set window properties for callback function
    lpWndProc = SetWindowLong(hWndButton, GWL_WNDPROC, AddressOf WindowProc)
    Call SetProp(hWndButton, HWND_PROP_lpPrevWndProc, lpWndProc)
    Call SetProp(hWndButton, HWND_PROP_hWndMainForm, hWndMain)
    Call SetProp(hWndButton, HWND_PROP_hWndTitleBarButton, hWndButton)
    Call SetProp(hWndButton, HWND_PROP_TitleBarButtonWidth, lWidthInPixels)
    Call SetProp(hWndButton, HWND_PROP_hWndTitleBarFakeButton, hWndFakeButton)

    ' Set the Fake TitleBar toggle button to Unchecked by default
    If bToggleButton Then Call SendMessage(hWndFakeButton, BM_SETCHECK, 0, 0)

    ' Set button pos before showing, must be called twice
    Call PostMessage(hWndMain, WM_SIZE, 0, 0)
    Call PostMessage(hWndMain, WM_SIZE, 0, 0)
    Call ShowWindow(hWndButton, SW_SHOWNOACTIVATE)

    ' Return focus to Main form
    Call SetFocus(hWndMain)

End Sub

Private Sub Class_Terminate()

    ' Remove window properties from button, unhook, and destroy
    Call SetWindowLong(hWndButton, GWL_WNDPROC, GetProp(hWndButton, HWND_PROP_lpPrevWndProc))
    Call RemoveProp(hWndButton, HWND_PROP_lpPrevWndProc)
    Call RemoveProp(hWndButton, HWND_PROP_hWndMainForm)
    Call RemoveProp(hWndButton, HWND_PROP_hWndTitleBarButton)
    Call RemoveProp(hWndButton, HWND_PROP_TitleBarButtonWidth)
    Call RemoveProp(hWndButton, HWND_PROP_hWndTitleBarFakeButton)
    Call DestroyWindow(hWndButton)

    ' Remove window properties from Main form and unhook
    Call SetWindowLong(hWndMain, GWL_WNDPROC, GetProp(hWndMain, HWND_PROP_lpPrevWndProc))
    Call RemoveProp(hWndMain, HWND_PROP_lpPrevWndProc)
    Call RemoveProp(hWndMain, HWND_PROP_hWndMainForm)
    Call RemoveProp(hWndMain, HWND_PROP_hWndTitleBarButton)

End Sub

