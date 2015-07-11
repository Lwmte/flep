Attribute VB_Name = "apiWindowSize"
'==========================================================================
'  SYSTEM METRICS AND WINDOW SIZE THIRD-PARTY API MODULE
'=========================================================================


Option Explicit

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
   
Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
      ByVal x As Long, ByVal y As Long, _
      ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Const SM_CYCAPTION      As Long = 4
Public Const SM_CXBORDER       As Long = 5
Public Const SM_CYBORDER       As Long = 6
Public Const SM_CXDLGFRAME     As Long = 7
Public Const SM_CYDLGFRAME     As Long = 8
Public Const SM_CXEDGE         As Long = 45
Public Const SM_CYEDGE         As Long = 46
Public Const SM_CXFRAME        As Long = 32
Public Const SM_CYFRAME        As Long = 33

Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000


Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum


Public Type POINTAPI
    x As Long
    y As Long
End Type


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public BorderWidth, BorderHeight, CaptionHeight As Long
