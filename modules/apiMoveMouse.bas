Attribute VB_Name = "apiMoveMouse"
'==========================================================================
'  code from vbkeys.com and/or vbstreets.ru. all
'  other non-original code sequences belong to their owners.
'=========================================================================

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'Узнаем координаты курсора мыши
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinini As _
Long) As Long 'В данном случае узнаем размеры рабочего стола для прилипаний
Private Const SPI_GETWORKAREA = 48 'Константа для вышенаписанной функции

Private IsMouseDwn As Boolean 'Служит для определения была ли нажата кнопка мыши
Dim NewCur As POINTAPI, FrmCur As POINTAPI, RC As RECT 'Работа с курсором и размером рабочей обл. экрана

Public Sub Object_MouseMove(frm As Form, ByVal x As Single, ByVal y As Single, ByVal Button As Integer, Optional ByVal Docking As Long)

    '**********************************
    'Если надо => укажите свою кнопку *
    '**********************************
    If Not Button = vbLeftButton Then IsMouseDwn = False: Exit Sub
    
        If Not IsMouseDwn Then 'Если кнопа не нажата => запоминаем координаты, узнаем область экрана
        
            FrmCur.x = x / Screen.TwipsPerPixelX
            FrmCur.y = y / Screen.TwipsPerPixelY
            Call SystemParametersInfo(SPI_GETWORKAREA, vbNull, RC, 0)
            IsMouseDwn = True 'Запоминаем, что кнопка нажата
            
        End If
        
        
    If IsMouseDwn Then 'Если кнопка нажата, работаем...
    
        Dim TempCur As POINTAPI
        GetCursorPos NewCur 'Новые координаты
        
        TempCur.x = (NewCur.x - FrmCur.x) 'Координаты верхнего, левого угла окна
        TempCur.y = (NewCur.y - FrmCur.y)
        
        If Abs(TempCur.x) - RC.Left < Docking Then 'Abs() - на и за экраном работаем...
            frm.Left = RC.Left * Screen.TwipsPerPixelX 'Если у края - прилипаем к левому краю
        ElseIf Abs(TempCur.x + frm.Width / Screen.TwipsPerPixelX - RC.Right) <= Docking Then
            frm.Left = RC.Right * Screen.TwipsPerPixelX - frm.Width 'Иначе смотрим на другой край и то же прилипаем
        Else
            frm.Left = TempCur.x * Screen.TwipsPerPixelX 'А еще иначе - двигаемся вслед за курсором
        End If
        
        If Abs(TempCur.y) - RC.Top < Docking Then 'Аналогично...
         frm.Top = RC.Top
        ElseIf Abs(TempCur.y + frm.Height / Screen.TwipsPerPixelY - RC.Bottom) <= Docking Then
         frm.Top = RC.Bottom * Screen.TwipsPerPixelY - frm.Height
        Else
         frm.Top = TempCur.y * Screen.TwipsPerPixelY
        End If
        
    End If

End Sub

