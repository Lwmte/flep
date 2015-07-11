Attribute VB_Name = "apiKernel"
'==========================================================================
'  GENERAL WINDOWS THIRD-PARTY API MODULE
'=========================================================================


Option Explicit

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETLINE = &HC4
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINEFROMCHAR = &HC9

Public Const WM_SETREDRAW = &HB

Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
                              ByVal wMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long

Declare Function SendMessageStr _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
                              ByVal wMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As String) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFilename As OPENFILENAME) As Long


Public Const IDC_HAND = 32649&
Public lHandCursorHandle As Long

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function SetCurrentDirectory Lib "kernel32" _
Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long




Public Function GetLineCount(SourcehWnd As Long) As Integer

    Dim Lines2&
    
    Lines2& = SendMessage(SourcehWnd, EM_GETLINECOUNT, 0, 0)
 
    GetLineCount = CInt(Lines2)

End Function


Public Function OpenFileDialog(ParentWindowHWnd As Long, StartDir As String, FilterTitle As String, FilterExtension As String, WindowTitle As String) As String

 Dim OFName As OPENFILENAME
 
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = ParentWindowHWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    OFName.lpstrFilter = FilterTitle + Chr$(0) + "*." + FilterExtension + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = StartDir
    'Set the title
    OFName.lpstrTitle = WindowTitle + " " + FilterTitle
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        OpenFileDialog = Trim$(OFName.lpstrFile)
    Else
        OpenFileDialog = k_IsCancelPressed
    End If
End Function


Public Function SaveFileDialog(ParentWindowHWnd As Long, StartDir As String, FilterTitle As String, FilterExtension As String, WindowTitle As String) As String

 Dim SFName As OPENFILENAME
 
    SFName.lStructSize = Len(SFName)
    'Set the parent window
    SFName.hwndOwner = ParentWindowHWnd
    'Set the application's instance
    SFName.hInstance = App.hInstance
    'Select a filter
    SFName.lpstrFilter = FilterTitle + Chr$(0) + "*." + FilterExtension + Chr$(0)
    'create a buffer for the file
    SFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    SFName.nMaxFile = 255
    'Create a buffer for the file title
    SFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    SFName.nMaxFileTitle = 255
    'Set the initial directory
    SFName.lpstrInitialDir = StartDir
    'Set the title
    SFName.lpstrTitle = WindowTitle + " " + FilterTitle
    'No flags
    SFName.flags = 0
    'Extension
    SFName.lpstrDefExt = FilterExtension

    'Show the 'Open File'-dialog
    If GetSaveFileName(SFName) Then
        SaveFileDialog = Trim$(SFName.lpstrFile)
    Else
        SaveFileDialog = k_IsCancelPressed
    End If

End Function

