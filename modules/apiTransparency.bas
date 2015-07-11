Attribute VB_Name = "apiTransparency"
'==========================================================================
' WINDOW TRANSPARENCY THIRD-PARTY API MODULE
'=========================================================================


Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long



Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_Defaut        As Long = &H2
Private Const WS_EX_LAYERED     As Long = &H80000


'==========================================================================
' FUNCTION: Transparency 1
'
'==========================================================================
Public Function Transparency(ByVal hwnd As Long, Optional ByVal Col As Long = vbBlack, _
    Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
' Return : True if there is no error.
' hWnd   : hWnd of the window to make transparent
' Col : Color to make transparent if TrMode=False
' PcTransp  : 0 A~  255 >> 0 = transparent  -:- 255 = Opaque
Dim DisplayStyle As Long
    On Error GoTo ErrorHandler
    DisplayStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
        DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
        Call SetWindowLong(hwnd, GWL_EXSTYLE, DisplayStyle)
    End If
    Transparency = (SetLayeredWindowAttributes(hwnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)
         
ErrorHandler:
    If Not Err.Number = 0 Then Err.Clear
End Function

'==========================================================================
' FUNCTION: ACTIVE TRANSPARENCY
' Sets window transparency.
'==========================================================================
Public Sub ActiveTransparency(M As Long, d As Boolean, F As Boolean, _
     T_Transparency As Integer, Optional Color As Long)
Dim B As Boolean
        If d And F Then
        'Makes color (here the background color of the shape) transparent
        'upon value of T_Transparency
            B = Transparency(M, Color, T_Transparency, False)
        ElseIf d Then
            'Makes form, including all components, transparent
            'upon value of T_Transparency
            B = Transparency(M, 0, T_Transparency, True)
        Else
            'Restores the form opaque.
            B = Transparency(M, , 255, True)
        End If
End Sub


'==========================================================================
' FUNCTION: Fade In Window
' Fades in window from Alpha=0 to Alhpa=FinalAlpha with specified Speed.
'==========================================================================
Public Sub FadeIn(Window As Object, Speed As Byte, FinalAlpha As Byte)

Dim i As Integer
    
    ActiveTransparency Window.hwnd, True, False, 0
    Window.Show
    For i = 0 To FinalAlpha Step Speed
        ActiveTransparency Window.hwnd, True, False, i
        Window.Refresh
    Next i
    
End Sub
