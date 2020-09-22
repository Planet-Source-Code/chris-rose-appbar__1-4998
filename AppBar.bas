Attribute VB_Name = "modAppBar"
Option Explicit

Type POINTAPI
        x As Long
        y As Long
End Type

Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type


Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const WM_MOUSEMOVE = &H200
Public Const WM_ACTIVATE = &H6
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const ABE_BOTTOM = 3
Public Const ABE_LEFT = 0
Public Const ABE_RIGHT = 2
Public Const ABE_TOP = 1
Public Const ABM_ACTIVATE = &H6
Public Const ABM_GETAUTOHIDEBAR = &H7
Public Const ABM_GETSTATE = &H4
Public Const ABM_GETTASKBARPOS = &H5
Public Const ABM_NEW = &H0
Public Const ABM_QUERYPOS = &H2
Public Const ABM_REMOVE = &H1
Public Const ABM_SETAUTOHIDEBAR = &H8
Public Const ABM_SETPOS = &H3
Public Const ABM_WINDOWPOSCHANGED = &H9
Public Const ABN_FULLSCREENAPP = &H2
Public Const ABN_POSCHANGED = &H1
Public Const ABN_STATECHANGE = &H0
Public Const ABN_WINDOWARRANGE = &H3

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

