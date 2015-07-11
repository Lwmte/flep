Attribute VB_Name = "funcFiles"
'==========================================================================
'  FLEP 1.0 FILE-RELATED OPERATIONS MODULE
'
'  Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'
'  Contains all code related to GLOBAL file modification routines.
'=========================================================================

Option Explicit

Public hFile As Long ' current filename made global, cause it gets used in other modules.


'==========================================================================
' FUNCTION: WRITE HEXADECIMAL STRING INTO FILE
' Writes hex string into file using system binary access.
'==========================================================================
Public Function WriteHex(ByVal SourceHexString As String, Offset As Long, FileHandler As Long) As Boolean

On Error GoTo ErrorHandler

 Dim OverallBytes As Integer
 Dim cntCurrentByte As Integer
 Dim tmpByteArray() As Byte
 
 Dim tmpBytesToWrite As Long

'   Check if string contains less than 2 characters, and if it is, fill it with
'   back-fill zeros.
    If Len(SourceHexString) = 0 Then Exit Function
    If Len(SourceHexString) = 1 Then SourceHexString = "0" & SourceHexString


'   Check if string contains even amount of characters, and if it is, cut it off.
    OverallBytes = Fix(Len(SourceHexString) / 2) - 1
    
'   Create byte array.
    ReDim tmpByteArray(OverallBytes)
    For cntCurrentByte = 0 To OverallBytes
        tmpByteArray(cntCurrentByte) = HxVal(Mid$(SourceHexString, ((cntCurrentByte * 2) + 1), 2))
    Next
    
'   Set file pointer to desired offset.
    Call SetFilePointer(ByVal FileHandler, ByVal Offset, 0, FILE_BEGIN)
    
'   Temporary variable, needed for external function call.
    tmpBytesToWrite = OverallBytes + 1
    
'   Write to file!
    Call WriteFile(ByVal FileHandler, ByVal VarPtr(tmpByteArray(0)), ByVal tmpBytesToWrite, ByVal VarPtr(tmpBytesToWrite), ByVal 0)
    
    If GetLastError <> 0 Then GoTo ErrorHandler 'Fail...
    
    WriteHex = True 'Success!
    Exit Function
    
ErrorHandler:

    WriteHex = False
    
End Function



'==========================================================================
' FUNCTION: BIN INTERPRETER
' Interprets raw string as a stream of commands and hex sequences, and does
' corresponding action.
'==========================================================================
Public Function BinInterpret(SrcString As String, FileName As String, Offset As Long) As Boolean

On Error GoTo ErrorHandler

Dim cntStrings As Integer '
Dim cntBinCmds As Integer '

Dim tmpCommandArray() As String '
Dim tmpStringArray() As String '
Dim tmpBinArray() As Byte

Dim tmpChar As Byte
Dim numBytesToFill As Long

Dim cntCounter As Long

Dim tmpBytesToWrite As Long


    tmpStringArray = Split(UCase$(StripOut(SrcString, " ")), kTerminator)
    
    
     For cntStrings = LBound(tmpStringArray) To UBound(tmpStringArray)
     
        If LenB(tmpStringArray(cntStrings)) > 0 Then
        
            tmpCommandArray = Split(tmpStringArray(cntStrings), kDivider)
            
                'SPEED-UP: interprets only strings that are shorter than max. command length.
                If LenB(tmpCommandArray(0)) <= bi_MaxCommandLength Then
        
                    Select Case tmpCommandArray(0)
                    
                        Case bi_kLen:  'SETFILELENGTH command...
                                    
                                    Call SetFilePointer(hFile, HxVal(tmpCommandArray(1)), 0, FILE_BEGIN)  ' Just set file pointer to desired EOF.
                                    Call SetEndOfFile(hFile) 'Set EOF in file itself.
                                    
    
                        Case bi_kFill: 'FILL command...
                        
                                    If LenB(tmpCommandArray(1)) > 0 And LenB(tmpCommandArray(2)) > 0 And HxVal(tmpCommandArray(2)) > 0 Then
                                    
                                    
                                        tmpChar = HxVal(Left$(tmpCommandArray(1), 2))
                                        numBytesToFill = HxVal(tmpCommandArray(2)) - 1
                                        ReDim tmpBinArray(numBytesToFill)
                                                                        
                                        For cntCounter = 0 To numBytesToFill
                                            tmpBinArray(cntCounter) = tmpChar
                                        Next cntCounter
                                        
                                        tmpBytesToWrite = numBytesToFill + 1
                                        
                                        Call SetFilePointer(ByVal hFile, ByVal Offset, 0, FILE_BEGIN)
                                        Call WriteFile(ByVal hFile, ByVal VarPtr(tmpBinArray(0)), ByVal tmpBytesToWrite, ByVal VarPtr(tmpBytesToWrite), ByVal 0)
    
                                        Offset = Offset + UBound(tmpBinArray) + 1
                                        
                                    End If
                                    
                                    
                        Case Else:    ' Anything else, usually hex string! :)
                        
                                    Call WriteHex(tmpCommandArray(0), Offset, hFile)
    
                    End Select
                    
                Else
                
                    Call WriteHex(tmpCommandArray(0), Offset, hFile)
                
                End If
                
            Erase tmpCommandArray
            
        End If
        
     Next cntStrings
     
     BinInterpret = True
     
     Exit Function
     
ErrorHandler:
    BinInterpret = False
    MsgBox "Cannot process with " & FileName & " modifying." & vbCrLf & "Reason: Error #" & CStr(Err.Number) & " (" & CStr(Err.Description) & ").", vbExclamation, "BinInterpret message" 'alarm user
    
End Function



'==========================================================================
' FUNCTION: CHECK EXISTENCE
' Check if file exists or not. If it is, returns True, else False.
'==========================================================================
Public Function CheckExistence(FileName As String) As Boolean

    If LenB(FileName) = 0 Then CheckExistence = False: Exit Function
    If LenB(Dir(FileName)) <> 0 Then CheckExistence = True Else CheckExistence = False

End Function



'==========================================================================
' FUNCTION: MODIFY BINARY
' Final procedure for modifying exe file.
' Only custom patch module (ex-DRACO) is going to be written.
' Returns error codes:
'   0 - Written successfully.
'   1 - Writing failed...
'==========================================================================
Public Function ModifyBinary(BinaryName As String, WindowHandler As Long) As Byte

On Error GoTo ErrorHandler
    
 Dim BinNameWithPath As String
 Dim TmpBinName As String
 
 Dim cntPatchUnit, cntDataUnit, cntParameterUnit, cntParameterOffset As Integer
 Dim tmpParamOffsets() As String
 
 Dim tmpBytesToWrite As Long
 
    
    ModifyBinary = 0  ' just in case...
    
    BinNameWithPath = BinaryName
    TmpBinName = BinNameWithPath & ".tmp"
    
    Sys.wFunc = FO_COPY
    Sys.pFrom = BinNameWithPath
    Sys.pTo = TmpBinName
    Sys.hwnd = WindowHandler
    Sys.fFlags = FOF_NOCONFIRMATION
    Call SHFileOperation(Sys)
    

    'Native Win32 function works much faster than VB's "Open as Binary".
    
    hFile = CreateFile(TmpBinName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_ALWAYS, 0, 0)
    
   
    For cntPatchUnit = LBound(PatchArray) To UBound(PatchArray)  'Cycle thru all patches...
        
        
        ' If patch enabled, and binary name matches...
        
        If PatchArray(cntPatchUnit).PatchEnabled = 1 And _
           LCase$(BinaryName) = LCase$(PatchArray(cntPatchUnit).patchFile) Then
        
        
        ' If patch selected in list, if target name is right and dependencies are OK, then...
        
           If IsPatchSelected(PatchNameToListIndex(PatchArray(cntPatchUnit).patchName)) = True And _
              DependencyTest(cntPatchUnit) = True Then
              

                       ' PUT DATA ENTRIES

                       On Error Resume Next 'It's needed for handling patches without datas, only with params.
                       For cntDataUnit = LBound(PatchArray(cntPatchUnit).patchDatas) To UBound(PatchArray(cntPatchUnit).patchDatas)
                       If Err.Number = 9 Then Err.Clear: Exit For
                       
                           On Error GoTo ErrorHandler
                       
                           If LenB(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataOffset) > 0 And _
                              LenB(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataModified) > 0 Then
                              
                                If BinInterpret(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataModified, TmpBinName, HxVal(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataOffset)) = False Then GoTo ErrorHandler

                           End If
                       
                       Next cntDataUnit
                       
                       
                       ' PUT PARAM ENTRIES
                   
                       On Error Resume Next
                       For cntParameterUnit = LBound(PatchArray(cntPatchUnit).patchParams) To UBound(PatchArray(cntPatchUnit).patchParams)
                       If Err.Number = 9 Then Err.Clear: Exit For
                       
                          
                           On Error GoTo ErrorHandler
                           If PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parEnabled = 1 And _
                              LenB(PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parModdedValue) > 0 Then
                              
                              tmpParamOffsets = Split(PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parOffset, kDivider)
                              
                                
                                On Error Resume Next
                                For cntParameterOffset = LBound(tmpParamOffsets) To UBound(tmpParamOffsets)
                                If Err.Number = 9 Then Err.Clear: Exit For
                                
                                    
                                    On Error GoTo ErrorHandler
                                    If LenB(tmpParamOffsets(cntParameterOffset)) > 0 Then
                                        
                                        If WriteParam(PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parModdedValue, _
                                             tmpParamOffsets(cntParameterOffset), _
                                             PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parType) = False Then GoTo ErrorHandler
                                        
                                    End If
                                    
                                Next cntParameterOffset
                              
                           End If
                           
                       Next cntParameterUnit
                   
           Else
                   
                   
                      ' PUT DATA ENTRIES
                   
                       On Error Resume Next
                       For cntDataUnit = LBound(PatchArray(cntPatchUnit).patchDatas) To UBound(PatchArray(cntPatchUnit).patchDatas)
                       If Err.Number = 9 Then Err.Clear: Exit For
                       
                           On Error GoTo ErrorHandler
                           If LenB(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataOffset) > 0 And _
                              LenB(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataDefault) > 0 And _
                              PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataCondBehave = 0 Then
                           
                                If BinInterpret(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataDefault, TmpBinName, HxVal(PatchArray(cntPatchUnit).patchDatas(cntDataUnit).dataOffset)) = False Then GoTo ErrorHandler
                           
                           End If
                       
                       Next cntDataUnit
                       

                       ' PUT PARAM ENTRIES
                       
                       On Error Resume Next
                       For cntParameterUnit = LBound(PatchArray(cntPatchUnit).patchParams) To UBound(PatchArray(cntPatchUnit).patchParams)
                       If Err.Number = 9 Then Err.Clear: Exit For
                       
                           If PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parEnabled = 1 And _
                              LenB(PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parValue) > 0 And _
                              PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parCondBehave = 0 Then
                              
                              tmpParamOffsets = Split(PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parOffset, kDivider)
                              
                                For cntParameterOffset = LBound(tmpParamOffsets) To UBound(tmpParamOffsets)
                                If Err.Number = 9 Then Err.Clear: Exit For
                                
                                    If LenB(tmpParamOffsets(cntParameterOffset)) > 0 Then
                                    
                                        If WriteParam(PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parValue, _
                                             tmpParamOffsets(cntParameterOffset), _
                                             PatchArray(cntPatchUnit).patchParams(cntParameterUnit).parType) = False Then GoTo ErrorHandler
                                        
                                    End If
                                
                                Next cntParameterOffset
        
                           End If
                           
                       Next cntParameterUnit
                               
                   On Error GoTo ErrorHandler
                         
           End If
           
        End If
    
    Next cntPatchUnit
    

    hFile = CloseHandle(hFile)
    
    '
    '
    ' Copy draft into exe
    '
    Sys.wFunc = FO_DELETE
    Sys.pFrom = BinNameWithPath
    Sys.pTo = BinNameWithPath
    Sys.hwnd = WindowHandler
    Sys.fFlags = FOF_NOCONFIRMATION
    Call SHFileOperation(Sys)
    '
    Sys.wFunc = FO_MOVE
    Sys.pFrom = TmpBinName
    Sys.pTo = BinNameWithPath
    Sys.hwnd = WindowHandler
    Sys.fFlags = FOF_NOCONFIRMATION
    Call SHFileOperation(Sys)

    
    ModifyBinary = 1
    
    Exit Function


ErrorHandler:

    CloseHandle hFile
    
    Sys.wFunc = FO_DELETE
    Sys.pFrom = TmpBinName
    Sys.pTo = TmpBinName
    Sys.hwnd = WindowHandler
    Sys.fFlags = FOF_NOCONFIRMATION
    Call SHFileOperation(Sys)
    '
    MsgBox "There were errors during patching." & vbCrLf & "Make sure file is not write-protected and check patch set consistency.", vbExclamation, "Cannot modify file" 'alarm user
    
    ModifyBinary = 0


End Function
