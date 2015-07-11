Attribute VB_Name = "engnCustomPatches"
'==========================================================================
'  FLEP 1.0 GENERAL CUSTOM PATCHES MODULE CODE (EX-DRACO MODULE CODE)
'
'  Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'=========================================================================
'
Option Explicit


Public Type typPatchParameter

    parEnabled As Byte          ' If parameter enabled or not
    parName As String           ' Name
    parOffset As String         ' Offsets (one parameter can have several offsets)
    parValue As String          ' Default value
    parModdedValue As String    ' Modified value
    parType As Integer          ' Parameter type
    parCondBehave As Byte       ' Conditional behaviour if disabled
    
End Type

Public Type typPatchData
    
    dataOffset As String        ' Patch offsets array
    dataDefault As String       ' Patch default data array. Must be equal to offsets array size.
    dataModified As String      ' Patch modified data array. Must be equal to offsets array size.
    dataCondBehave As Byte      ' Conditional behaviour if disabled

End Type
'
Public Type typPatchEntry

    PatchEnabled As Byte        ' If patch enabled or not
    patchName As String         ' Patch name
    patchDesc As String         ' Patch description
    patchFile As String         ' Target filename - e. g., each patch can be used for specific file.
    patchCategory As String     ' Patch category; each new category creates new entry in TreeView (prototyped).
    '
    patchDependencies As String ' Patch dependencies from other patches, separated by commas.
    '
    patchDatas() As typPatchData       ' Main patch data(s) array. See typPatchData type for description.
    patchParams() As typPatchParameter ' Patch parameter(s) array. See typPatchParameter for description.
    '
End Type


Public Type typExecData     ' Datatype for binary array.
    execName As String
    execLastCRC As Long
    execPatched As Boolean
End Type
'

'
Public CurrentDirectory As String   ' current directory with "\" symbol added. used to force searching for exe's in program's directory.
Public BackupDirectory As String    ' backup directory to use with presets or patch sets.

Public CurrentSetName As String     ' current patch set name.
Public ConfigPath As String
'
Public PatchEngineVersion As String ' Patch engine version, as read from .cfg.

'
Public PatchArray() As typPatchEntry
'
Public ExecList() As typExecData



'==========================================================================
' SUB: RESETS ALL ARRAYS.
'==========================================================================
Public Sub ResetArrays()
     
     ReDim PatchArray(0)   ' Reset patch array, as we're loading it again.
        
     PatchArray(0).PatchEnabled = 1
     PatchArray(0).patchName = k_EmptyPatchName
     PatchArray(0).patchFile = DefaultExeName
     

End Sub


'==========================================================================
' FUNCTION: CREATE EXECUTABLE LIST
' Re-creates executable list by scan-and-seek method for each patch.
'==========================================================================
Public Sub CreateExecList()

Dim cntPatchCounter As Integer

    Erase ExecList
    
    For cntPatchCounter = 0 To UBound(PatchArray)
    
        If Trim$(PatchArray(cntPatchCounter).patchFile) = vbNullString Then
            Call UpdateExecList(DefaultExeName)
        Else
            Call UpdateExecList(Trim$(PatchArray(cntPatchCounter).patchFile))
        End If
    
    Next cntPatchCounter

End Sub


'==========================================================================
' FUNCTION: UpdateExecList
' SCANS ANDS ADD EXE INTO DATABASE, IF NOT ALREADY EXISTS.
'==========================================================================
Public Function UpdateExecList(SrcName As String, Optional CheckSum As Long) As Boolean

    On Error GoTo ErrorHandler

    Dim cntExecCounter As Integer
    Dim cntEmptyUnit   As Integer
 
    If LenB(SrcName) = 0 Then Exit Function
 
    UpdateExecList = False
    SrcName = LCase$(SrcName)
    
    On Error Resume Next
 
    For cntExecCounter = LBound(ExecList) To UBound(ExecList)
    
        If Err.Number = 9 Then
            Err.Clear
            ReDim ExecList(0)
            ExecList(0).execName = SrcName
            If LenB(CheckSum) > 0 Then
                ExecList(UBound(ExecList)).execLastCRC = CheckSum
                ExecList(UBound(ExecList)).execPatched = True
            Else
                ExecList(UBound(ExecList)).execLastCRC = vbNullString
                ExecList(UBound(ExecList)).execPatched = False
            End If
            Exit Function
        End If
    
        If LCase$(ExecList(cntExecCounter).execName) = SrcName Then Exit Function ' if already exists...
        
    Next cntExecCounter
    
    On Error GoTo ErrorHandler

    ReDim Preserve ExecList(UBound(ExecList) + 1)
    ExecList(UBound(ExecList)).execName = SrcName

    
    If LenB(CheckSum) > 0 Then
        ExecList(UBound(ExecList)).execLastCRC = CheckSum
        ExecList(UBound(ExecList)).execPatched = True
    Else
        ExecList(UBound(ExecList)).execLastCRC = vbNullString
        ExecList(UBound(ExecList)).execPatched = False
    End If
    
    
    UpdateExecList = True
    
    Exit Function

ErrorHandler:
    UpdateExecList = False

End Function


'==========================================================================
' FUNCTION: CHECK CRC
' Checks if files' checksums are equal to those present in executable list.
' If not, immediately brings message box for user.
'==========================================================================
Public Sub CheckCRC()

On Error Resume Next

 Dim cntExecCounter As Integer
 Dim tmpCRCList As String

    For cntExecCounter = LBound(ExecList) To UBound(ExecList)
    If Err.Number = 9 Then Err.Clear: Exit Sub
    
        If CheckExistence(ExecList(cntExecCounter).execName) = True Then
        
            If GetCRCFromFile(ExecList(cntExecCounter).execName) <> ExecList(cntExecCounter).execLastCRC And ExecList(cntExecCounter).execLastCRC <> 0 Then tmpCRCList = tmpCRCList & "  " & ExecList(cntExecCounter).execName & vbCrLf
        
        End If
    
    Next cntExecCounter
    
    If LenB(tmpCRCList) > 0 And MarkSilently = 0 Then MsgBox "It seems that these files were externally modified since last patching: " & vbCrLf & tmpCRCList & vbCrLf & "Possibly wrong version or another patcher used?", vbInformation

End Sub


'==========================================================================
' FUNCTION: UPDATE CRC
' Updates CRC for all executable list entries.
'==========================================================================
Public Sub UpdateCRC()

 Dim cntExecCounter As Integer
 Dim tmpCRCList As String

    For cntExecCounter = LBound(ExecList) To UBound(ExecList)
    
        If CheckExistence(ExecList(cntExecCounter).execName) = True Then
        
            ExecList(cntExecCounter).execLastCRC = GetCRCFromFile(ExecList(cntExecCounter).execName)
            
        Else
        
            ExecList(cntExecCounter).execLastCRC = 0
        
        End If
    
    Next cntExecCounter

End Sub


'==========================================================================
' FUNCTION: CONVERT PATCH PRESET
' Injects patch preset string values into corresponding parameter array units.
'==========================================================================
Public Function ConvertPatchPreset(RawString As String, ByVal DesiredPatchArrayNum As Integer) As Boolean

On Error GoTo ErrorHandler 'if overflow or end string, then stop execution

 Dim cntPointer As Byte
 Dim tmpStringArray() As String
 
 
    tmpStringArray = Split(RawString, kDivider2)
    ReDim Preserve tmpStringArray(UBound(PatchArray(DesiredPatchArrayNum).patchParams))
    
    For cntPointer = 0 To UBound(tmpStringArray)
    
        If LenB(tmpStringArray(cntPointer)) > 0 Then
            PatchArray(DesiredPatchArrayNum).patchParams(cntPointer).parModdedValue = tmpStringArray(cntPointer)
        Else
            PatchArray(DesiredPatchArrayNum).patchParams(cntPointer).parModdedValue = PatchArray(DesiredPatchArrayNum).patchParams(cntPointer).parValue
        End If
    
    Next cntPointer
        
    ConvertPatchPreset = True
    
    Exit Function

ErrorHandler:
    ConvertPatchPreset = False 'fuk em...
    
End Function


'==========================================================================
' FUNCTION: RE-ASSIGN PATCH DATAS SUB-ARRAY
' This is needed every time we assign new data offset in patch parameters.
'==========================================================================
Public Function ReDimPatchDatas(PatchUnitIndex As Integer, RedimIndex As Integer) As Boolean

On Error Resume Next

    If RedimIndex > UBound(PatchArray(PatchUnitIndex).patchDatas) Then
    
        If Err.Number = 9 Then
            ReDim PatchArray(PatchUnitIndex).patchDatas(RedimIndex)
            Err.Clear
        End If
        
        ReDim Preserve PatchArray(PatchUnitIndex).patchDatas(RedimIndex)
    End If
    
    ReDimPatchDatas = True
    Exit Function

ErrorHandler:
    MsgBox "Can't resize patch datas sub-array to " + CStr(RedimIndex) + " for patch # " + CStr(PatchUnitIndex) + "."
    ReDimPatchDatas = False

End Function


'==========================================================================
' FUNCTION: RE-ASSIGN PATCH PARAMS SUB-ARRAY
' This is needed every time we assign new param offset in patch parameters.
'==========================================================================
Public Function ReDimPatchParams(PatchUnitIndex As Integer, RedimIndex As Integer) As Boolean

On Error Resume Next

    If RedimIndex > UBound(PatchArray(PatchUnitIndex).patchParams) Then
    
        If Err.Number = 9 Then
            ReDim PatchArray(PatchUnitIndex).patchParams(RedimIndex)
            Err.Clear
        End If
    
        ReDim Preserve PatchArray(PatchUnitIndex).patchParams(RedimIndex)
    End If
    
    ReDimPatchParams = True
    Exit Function

ErrorHandler:
    MsgBox "Can't resize patch parameters sub-array to " + CStr(RedimIndex) + " for patch # " + CStr(PatchUnitIndex) + "."
    ReDimPatchParams = False
    
End Function


'==========================================================================
' SUB: LOAD DEFAULT PATCH SET
' Load default patch set.
'==========================================================================
Public Sub LoadDefaultPatchSet()

 Dim tmpSuccessFlag As Boolean

    If CheckExistence(CurrentSetName) = False Then
        RecreatePatchSetFile (CurrentSetName)
    End If
    
    tmpSuccessFlag = LoadPatches(CurrentSetName, True)

End Sub


'==========================================================================
' FUNCTION: LOAD PATCHES FROM CONFIGURATION FILE
' Completely reworked LoadSet procedure to load custom patches.
'==========================================================================
Public Function LoadPatches(PatchSetFileName As String, StartUp As Boolean) As Boolean

On Error GoTo ErrorHandler

 Dim tempString As String
 Dim TempCFG As typCfgParam
 Dim TempNumStr As typNumString
 Dim TempFunctionCallbackFlag As Boolean
 
 Dim VersionFound As Boolean
 
 Dim CurPatchNumber As Integer
 Dim MaxPatchNumber As Integer
 
 
    LoadPatches = False
    VersionFound = False  ' reset flag just in case...
    MaxPatchNumber = -1
    CurPatchNumber = 0    ' reset variables just in case...
    

    If PatchSetFileName = vbNullString Then ' If no patch set specified...
        MsgBox "Patch set isn't specified. No patches loaded."
        Exit Function
    End If
    
    If CheckExistence(PatchSetFileName) = False Then ' If no patch set found...
        MsgBox "Can't find patch set file (" & PatchSetFileName & "). No patches loaded."
        Exit Function
    End If
    
    
    Open PatchSetFileName For Input As #54
    
    Do Until EOF(54)
    
        If VersionFound = False Then
        
            Do Until LCase$(TempCFG.Name) = kPatchSetVer
                Line Input #54, tempString
                TempCFG = ConvCFG(tempString)
                If EOF(54) Then GoTo ErrorHandler  ' If we've not found patch set header, exit...
            Loop
            
            
            Select Case Val(TempCFG.Value)
                Case Is <= Val(k_HeaderVersion):  PatchEngineVersion = TempCFG.Value
                                                  VersionFound = True
                                                  Call ResetArrays
                Case Is > Val(k_HeaderVersion):   Call MsgBox("Warning: specified patch set has higher version number." & vbCrLf & "Please update FLEP program!", vbCritical)
                                                  Close #54
                                                  If StartUp = True Then Call ResetArrays: PopulatePatchList: RefreshView
                                                  Exit Function
            End Select
            
        End If
    
    
        Line Input #54, tempString
        TempCFG = ConvCFG(tempString)
    
    
        If LCase$(TempCFG.Name) = kPatchHeader Then ' If patch unit header found, else exit...

            CurPatchNumber = CInt(TempCFG.Value)
            
            If CurPatchNumber > MaxPatchNumber Then
            
                MaxPatchNumber = CurPatchNumber
                ReDim Preserve PatchArray(MaxPatchNumber)
                
            End If
            
            
            
                Do Until LCase$(TempCFG.Name) = kPatchFooter  ' Fill current patch array until footer is found.
                '
                    Line Input #54, tempString
                    TempCFG = ConvCFG(tempString)
                    '
                    Select Case LCase$(TempCFG.Name)
                    
                    
                        ' Common...
                        
                        Case kPatchEnabled: PatchArray(CurPatchNumber).PatchEnabled = CByteL(TempCFG.Value)
                        Case kPatchName: PatchArray(CurPatchNumber).patchName = TempCFG.Value
                        Case kPatchDesc: PatchArray(CurPatchNumber).patchDesc = TempCFG.Value
                        Case kPatchCategory: PatchArray(CurPatchNumber).patchCategory = TempCFG.Value
                        Case kPatchDependencies: PatchArray(CurPatchNumber).patchDependencies = TempCFG.Value
                        Case kPatchFile: PatchArray(CurPatchNumber).patchFile = TempCFG.Value
                        
                        ' Datas...
                        
                        Case kDataOffset
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchDatas(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchDatas(TempNumStr.Number).dataOffset = TempNumStr.RawString
                            
                        Case kDetaDefault
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchDatas(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchDatas(TempNumStr.Number).dataDefault = TempNumStr.RawString
                            
                        Case kDataModified
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchDatas(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchDatas(TempNumStr.Number).dataModified = TempNumStr.RawString
                            
                        Case kDataCondBehave
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchDatas(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchDatas(TempNumStr.Number).dataCondBehave = CByteL(TempNumStr.RawString)

                        ' Params...
                        
                        Case kParEnabled
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchParams(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parEnabled = CByteL(TempNumStr.RawString)
                        
                        Case kParName
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchParams(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parName = TempNumStr.RawString
                        
                        Case kParOffset
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchParams(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parOffset = TempNumStr.RawString
                        
                        Case kParValue
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchParams(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parValue = TempNumStr.RawString
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parModdedValue = TempNumStr.RawString
                        
                        Case kParType
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchParams(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parType = CInt(TempNumStr.RawString)
                        
                        Case kParCondBehave
                            TempNumStr = ParamsToNumString(TempCFG.Value)
                            TempFunctionCallbackFlag = ReDimPatchParams(CurPatchNumber, CInt(TempNumStr.Number))
                            If TempFunctionCallbackFlag = True Then _
                            PatchArray(CurPatchNumber).patchParams(TempNumStr.Number).parCondBehave = CByteL(TempNumStr.RawString)
                        
                    End Select
            
                Loop
        '
        Else
        
            If EOF(54) Then Exit Do ' Later you can add common settings here.
            
        End If
        '
    Loop
    
    Close #54
    
    Call LockWindowUpdate(NewMainWindow.hwnd)
    
        Call PopulatePatchList
        
        If DeveloperMode = 1 And DeveloperView = 1 Then
            Call RefreshView
            Call BackupPatch(0)
        End If
        
        LoadPatches = True
    
    Call LockWindowUpdate(0)
    
    Exit Function

ErrorHandler:
    Dim tmpErrorString As String
    
        Close #54
        Call LockWindowUpdate(0)
        tmpErrorString = "Error "
        If Err.Number > 0 Then tmpErrorString = tmpErrorString & "#" & CStr(Err.Number)
        MsgBox tmpErrorString & " loading custom patch set (" & PatchSetFileName & ")." & vbCrLf & "Make sure file is not corrupted or empty."

End Function


'==========================================================================
' FUNCTION: WRITE SETTING
' Small helper function to form config string with specified looks.
'==========================================================================
Public Function WriteSetting(SettingName As String, SettingValue As Variant, Optional Padding As Integer, Optional SkipLines As Long) As String

Dim cntCrLf As Integer

  WriteSetting = Space$(Padding) & SettingName & kEquals & CStr(SettingValue)
  If SkipLines > 0 Then WriteSetting = WriteSetting & String$(SkipLines, vbCrLf)

End Function


'==========================================================================
' FUNCTION: SAVE PATCHES TO CONFIGURATION FILE
' Completely reworked SaveSet procedure to save custom patches.
'==========================================================================
Public Function SavePatches() As Boolean

On Error GoTo ErrorHandler

 Dim tmpSuccessFlag As Boolean
 Dim tmpFileName As String
 Dim cntPatchCounter, cntUnitCounter, cntSubUnitCounter As Integer
 
    SavePatches = False ' just In case....

    tmpFileName = SaveFileDialog(NewMainWindow.hwnd, CurrentDirectory, k_FilterTitle_PatchSet, k_FilterExt_PatchSet, k_SaveFileTitle)

    If tmpFileName = k_IsCancelPressed Then SavePatches = True: Exit Function
    
     Call ApplyPatch(PrevPatchNumber) 'update current patch
    
     tmpFileName = tmpFileName & "." & k_FilterExt_PatchSet
    
    Open tmpFileName For Output As #1


    Print #1, WriteSetting(kPatchSetVer, k_HeaderVersion, 0, 2)
    
    
    For cntPatchCounter = LBound(PatchArray) To UBound(PatchArray)
    
        Print #1, WriteSetting(kPatchHeader, cntPatchCounter, 2, 1)
        
        Print #1, WriteSetting(kPatchEnabled, PatchArray(cntPatchCounter).PatchEnabled, 6)
        Print #1, WriteSetting(kPatchName, PatchArray(cntPatchCounter).patchName, 6)
        Print #1, WriteSetting(kPatchDesc, PatchArray(cntPatchCounter).patchDesc, 6)
        Print #1, WriteSetting(kPatchCategory, PatchArray(cntPatchCounter).patchCategory, 6)
        Print #1, WriteSetting(kPatchDependencies, PatchArray(cntPatchCounter).patchDependencies, 6)
        Print #1, WriteSetting(kPatchFile, PatchArray(cntPatchCounter).patchFile, 6, 1)
        
        On Error Resume Next
        
            For cntUnitCounter = LBound(PatchArray(cntPatchCounter).patchDatas) To UBound(PatchArray(cntPatchCounter).patchDatas)
            If Err.Number = 9 Then Err.Clear: Exit For
            
                If LenB(PatchArray(cntPatchCounter).patchDatas(cntUnitCounter).dataOffset) > 0 Then
            
                         Print #1, WriteSetting(kDataOffset, CStr(cntUnitCounter) + kDivider + PatchArray(cntPatchCounter).patchDatas(cntUnitCounter).dataOffset, 6)
                         Print #1, WriteSetting(kDetaDefault, CStr(cntUnitCounter) + kDivider + CipherText(PatchArray(cntPatchCounter).patchDatas(cntUnitCounter).dataDefault), 6)
                         Print #1, WriteSetting(kDataModified, CStr(cntUnitCounter) + kDivider + CipherText(PatchArray(cntPatchCounter).patchDatas(cntUnitCounter).dataModified), 6)
                         Print #1, WriteSetting(kDataCondBehave, CStr(cntUnitCounter) + kDivider + CStr(PatchArray(cntPatchCounter).patchDatas(cntUnitCounter).dataCondBehave), 6, 1)
                         
                End If
                        
            Next cntUnitCounter
            
            
            For cntUnitCounter = LBound(PatchArray(cntPatchCounter).patchParams) To UBound(PatchArray(cntPatchCounter).patchParams)
            If Err.Number = 9 Then Err.Clear: Exit For
            
                If LenB(PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parEnabled) > 0 Then
            
                         Print #1, WriteSetting(kParEnabled, CStr(cntUnitCounter) + kDivider + CStr(PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parEnabled), 6)
                         Print #1, WriteSetting(kParName, CStr(cntUnitCounter) + kDivider + PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parName, 6)
                         Print #1, WriteSetting(kParOffset, CStr(cntUnitCounter) + kDivider + PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parOffset, 6)
                         Print #1, WriteSetting(kParValue, CStr(cntUnitCounter) + kDivider + PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parValue, 6)
                         Print #1, WriteSetting(kParType, CStr(cntUnitCounter) + kDivider + CStr(PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parType), 6)
                         Print #1, WriteSetting(kParCondBehave, CStr(cntUnitCounter) + kDivider + CStr(PatchArray(cntPatchCounter).patchParams(cntUnitCounter).parCondBehave), 6, 1)
                            

                End If
                        
            Next cntUnitCounter
    
        On Error GoTo ErrorHandler
    
        Print #1, WriteSetting(kPatchFooter, cntPatchCounter, 2, 2)
        
    Next cntPatchCounter
    
    Close #1
    
    SavePatches = True
    
    Exit Function
    
ErrorHandler:

    SavePatches = False
    Close #1
    Exit Function


End Function
