VERSION 5.00
Begin VB.Form frmAppBar 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   1680
   ClientWidth     =   15255
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1200
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrHide 
      Interval        =   1
      Left            =   3720
      Top             =   120
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   15255
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton Command1 
         Caption         =   "Click here to quit, or double click on the form"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmAppBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BarData As APPBARDATA

Dim bAutoHide As Boolean
Dim bAnimate As Boolean








Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

    Dim lResult As Long

    Move 0, 0, 0, 0
    Screen.MousePointer = vbDefault
    
    bAutoHide = True
    bAnimate = True
    
    BarData.cbSize = Len(BarData)
    BarData.hwnd = hwnd
    BarData.uCallbackMessage = WM_MOUSEMOVE
    lResult = SHAppBarMessage(ABM_NEW, BarData)
    lResult = SetRect(BarData.rc, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN))
    BarData.uEdge = ABE_TOP
    lResult = SHAppBarMessage(ABM_QUERYPOS, BarData)
    If bAutoHide Then
        BarData.rc.Bottom = BarData.rc.Top + 2
        lResult = SHAppBarMessage(ABM_SETPOS, BarData)
        BarData.lParam = True
        lResult = SHAppBarMessage(ABM_SETAUTOHIDEBAR, BarData)
        If lResult = 0 Then
            bAutoHide = False
        Else
            lResult = SetWindowPos(BarData.hwnd, HWND_TOP, BarData.rc.Left, BarData.rc.Top - 42, BarData.rc.Right - BarData.rc.Left, 44, SWP_NOACTIVATE)
        End If
    End If
    If Not bAutoHide Then
        BarData.rc.Bottom = BarData.rc.Top + 42
        lResult = SHAppBarMessage(ABM_SETPOS, BarData)
        lResult = SetWindowPos(BarData.hwnd, HWND_TOP, BarData.rc.Left, BarData.rc.Top, BarData.rc.Right - BarData.rc.Left, 42, SWP_NOACTIVATE)
    End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Static bRecieved As Boolean
    Dim lResult As Long
    Dim newRC As RECT
    Dim lMessage As Long
    
    lMessage = x / Screen.TwipsPerPixelX
    
    If bRecieved = False Then
        bRecieved = True
        Select Case lMessage
            Case WM_ACTIVATE
                lResult = SHAppBarMessage(ABM_ACTIVATE, BarData)
            Case WM_WINDOWPOSCHANGED
                lResult = SHAppBarMessage(ABM_WINDOWPOSCHANGED, BarData)
            Case ABN_STATECHANGE
                lResult = SetRect(BarData.rc, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN))
                BarData.uEdge = ABE_TOP
                lResult = SHAppBarMessage(ABM_QUERYPOS, BarData)
                If bAutoHide Then
                    BarData.rc.Bottom = BarData.rc.Top + 2
                    lResult = SHAppBarMessage(ABM_SETPOS, BarData)
                    BarData.lParam = True
                    lResult = SHAppBarMessage(ABM_SETAUTOHIDEBAR, BarData)
                    If lResult = 0 Then
                        bAutoHide = False
                    Else
                        lResult = SetWindowPos(BarData.hwnd, HWND_TOP, BarData.rc.Left, BarData.rc.Top - 42, BarData.rc.Right - BarData.rc.Left, 44, SWP_NOACTIVATE)
                    End If
                End If
                If Not bAutoHide Then
                    BarData.rc.Bottom = BarData.rc.Top + 42
                    lResult = SHAppBarMessage(ABM_SETPOS, BarData)
                    lResult = SetWindowPos(BarData.hwnd, HWND_TOP, BarData.rc.Left, BarData.rc.Top, BarData.rc.Right - BarData.rc.Left, 42, SWP_NOACTIVATE)
                End If
            Case ABN_FULLSCREENAPP
                Beep
        End Select
        bRecieved = False
    End If
End Sub

Private Sub Form_Resize()
    picFrame.Move 0, 0, Width, Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If BarData.hwnd <> 0 Then SHAppBarMessage ABM_REMOVE, BarData
End Sub









Private Sub picFrame_DblClick()
    Unload Me
End Sub

Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lResult As Long
    Dim iCounter As Integer
    If Top < 0 Then
        If bAnimate Then
            For iCounter = -36 To -1
                BarData.rc.Top = iCounter
                lResult = SetWindowPos(BarData.hwnd, 0&, BarData.rc.Left, BarData.rc.Top, BarData.rc.Right - BarData.rc.Left, 42, SWP_NOACTIVATE)
            Next
        End If
        BarData.rc.Top = 0
        lResult = SetWindowPos(BarData.hwnd, HWND_TOPMOST, BarData.rc.Left, BarData.rc.Top, BarData.rc.Right - BarData.rc.Left, 42, SWP_SHOWWINDOW)
        tmrHide.Enabled = True
    End If
End Sub

















Private Sub tmrHide_Timer()
    Dim lResult As Long
    Dim lpPoint As POINTAPI
    Dim iCounter As Integer
    lResult = GetCursorPos(lpPoint)
    If lpPoint.x < Left \ Screen.TwipsPerPixelX Or lpPoint.x > (Left + Width) \ Screen.TwipsPerPixelX Or lpPoint.y < Top \ Screen.TwipsPerPixelY Or lpPoint.y - 10 > (Top + Height) \ Screen.TwipsPerPixelY Then
        If bAnimate Then
            For iCounter = -1 To -37 Step -1
                BarData.rc.Top = iCounter
                lResult = SetWindowPos(BarData.hwnd, 0&, BarData.rc.Left, BarData.rc.Top, BarData.rc.Right - BarData.rc.Left, 42, SWP_NOACTIVATE)
            Next
        End If
        BarData.rc.Top = -42
        lResult = SetWindowPos(BarData.hwnd, HWND_TOPMOST, BarData.rc.Left, BarData.rc.Top, BarData.rc.Right - BarData.rc.Left, 44, SWP_NOACTIVATE)
        tmrHide.Enabled = False
    End If
End Sub


