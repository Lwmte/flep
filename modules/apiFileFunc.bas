Attribute VB_Name = "apiFileFunc"
'==========================================================================
'  FILE FUNCTIONS THIRD-PARTY API MODULE
'=========================================================================


Option Explicit

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type
'
Public Declare Function SHFileOperation Lib "Shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'
Public Sys As SHFILEOPSTRUCT
Public Const FO_COPY = &H2&
Public Const FO_DELETE = &H3&
Public Const FO_MOVE = &H1&
Public Const FO_RENAME = &H4&
Public Const FOF_SILENT = &H4
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_NOERRORUI = &H400
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_RENAMEONCOLLISION = &H8
'
'/------------------------------------------------ DECLARE/CONSTS FOR CUTFILE
'
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_ALWAYS = 4
Public Const FILE_BEGIN = 0
