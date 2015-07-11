Attribute VB_Name = "apiMoveMouse"
'==========================================================================
'  code from vbkeys.com and/or vbstreets.ru. all
'  other non-original code sequences belong to their owners.
'=========================================================================

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long '������ ���������� ������� ����
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinini As _
Long) As Long '� ������ ������ ������ ������� �������� ����� ��� ����������
Private Const SPI_GETWORKAREA = 48 '��������� ��� �������������� �������

Private IsMouseDwn As Boolean '������ ��� ����������� ���� �� ������ ������ ����
Dim NewCur As POINTAPI, FrmCur As POINTAPI, RC As RECT '������ � �������� � �������� ������� ���. ������

Public Sub Object_MouseMove(frm As Form, ByVal x As Single, ByVal y As Single, ByVal Button As Integer, Optional ByVal Docking As Long)

    '**********************************
    '���� ���� => ������� ���� ������ *
    '**********************************
    If Not Button = vbLeftButton Then IsMouseDwn = False: Exit Sub
    
        If Not IsMouseDwn Then '���� ����� �� ������ => ���������� ����������, ������ ������� ������
        
            FrmCur.x = x / Screen.TwipsPerPixelX
            FrmCur.y = y / Screen.TwipsPerPixelY
            Call SystemParametersInfo(SPI_GETWORKAREA, vbNull, RC, 0)
            IsMouseDwn = True '����������, ��� ������ ������
            
        End If
        
        
    If IsMouseDwn Then '���� ������ ������, ��������...
    
        Dim TempCur As POINTAPI
        GetCursorPos NewCur '����� ����������
        
        TempCur.x = (NewCur.x - FrmCur.x) '���������� ��������, ������ ���� ����
        TempCur.y = (NewCur.y - FrmCur.y)
        
        If Abs(TempCur.x) - RC.Left < Docking Then 'Abs() - �� � �� ������� ��������...
            frm.Left = RC.Left * Screen.TwipsPerPixelX '���� � ���� - ��������� � ������ ����
        ElseIf Abs(TempCur.x + frm.Width / Screen.TwipsPerPixelX - RC.Right) <= Docking Then
            frm.Left = RC.Right * Screen.TwipsPerPixelX - frm.Width '����� ������� �� ������ ���� � �� �� ���������
        Else
            frm.Left = TempCur.x * Screen.TwipsPerPixelX '� ��� ����� - ��������� ����� �� ��������
        End If
        
        If Abs(TempCur.y) - RC.Top < Docking Then '����������...
         frm.Top = RC.Top
        ElseIf Abs(TempCur.y + frm.Height / Screen.TwipsPerPixelY - RC.Bottom) <= Docking Then
         frm.Top = RC.Bottom * Screen.TwipsPerPixelY - frm.Height
        Else
         frm.Top = TempCur.y * Screen.TwipsPerPixelY
        End If
        
    End If

End Sub

