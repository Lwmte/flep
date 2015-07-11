Attribute VB_Name = "engnFLEPSystem"
'==========================================================================
'  FLEP 1.0 SYSTEM MODULE CODE
'
'  Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'
'  This module contains all procedures and functions related to
'  NewMainWindow.frm interface controls, such as list sorting, comparing
'  etc.
'
'=========================================================================
'

Option Explicit

'==========================================================================
' FUNCTION: LOAD APPLICATION
' Init routines...
'==========================================================================
Public Sub LoadProgram()

    lHandCursorHandle = LoadCursor(0, IDC_HAND)

    NewMainWindow.lbOffsetList.Clear
    
    CurrentDirectory = CurDir$
    ConfigPath = k_ConfigName
    CurrentSetName = "patches.flp"
    
     ResetArrays
         
End Sub



'==========================================================================
' FUNCTION: KILL APPLICATION
' Kills application completely.
'==========================================================================
Public Sub KillApp()

    Unload NewMainWindow
    Unload InputKor
    Unload ParamList
    Unload AboutWindow

End Sub


'==========================================================================
' FUNCTION: MODIFY ALL EXECUTABLES
' This is general function for patching.
'==========================================================================

Public Sub ModifyAllExecutables()

On Error GoTo ErrorHandler

 Dim cntExecCounter As Integer

 Dim tmpSuccessFlag As Integer

    If MarkSilently = 0 Then
        If MsgBox("Continue with patching?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    tmpSuccessFlag = 1
    
    BackupDirectory = CurDir$
    Call SetCurrentDirectory(CurrentDirectory)
    

    For cntExecCounter = 0 To UBound(ExecList)
        
        If CheckExistence(ExecList(cntExecCounter).execName) = False And MarkSilently = 0 Then
    
            If cntExecCounter = UBound(ExecList) Or UBound(ExecList) = 0 Then
                MsgBox "Can't find file: " & ExecList(cntExecCounter).execName & "." & vbCrLf & "No patches were applied.", vbExclamation: Exit Sub
            Else
                If MsgBox("Can't find: " & ExecList(cntExecCounter).execName & "." & vbCrLf & "Continue with other files?", vbYesNo) = vbNo Then Exit Sub
            End If
            
        Else
            
            Call ModifyBinary(ExecList(cntExecCounter).execName, NewMainWindow.hwnd)

        End If
        
    Next cntExecCounter
    
    Call UpdateCRC
    Call SetCurrentDirectory(BackupDirectory)
    
    NewMainWindow.btnModify.Caption = k_ModButtonSuccess  ' Neat stuff...
    NewMainWindow.tmrTipTimer.Enabled = True
    
    Exit Sub
    
ErrorHandler:
    NewMainWindow.btnModify.Caption = k_ModButton

End Sub



Public Function RecreateConfigFile(ConfigFileName As String)

    Open ConfigFileName For Output As #1 'create new file
    Print #1, StrConv(LoadResData(103, "CUSTOM"), vbUnicode) 'write down text in form
    Close #1 'close file

End Function



Public Function RecreatePatchSetFile(PatchSetFileName As String)

    Open PatchSetFileName For Output As #1 'create new file
    Print #1, StrConv(LoadResData(104, "CUSTOM"), vbUnicode) 'write down text in form
    Close #1 'close file

End Function



'==========================================================================
' FUNCTION: CHECK EQUAL NAMES
' Scans patches array for desired patch name and selects corresponding list
' item.
'==========================================================================
Public Function CheckEqualPatchNames(SrcStr As String) As Boolean

 Dim cntPatchCount As Integer
 
    For cntPatchCount = LBound(PatchArray) To UBound(PatchArray)
        If PatchArray(cntPatchCount).patchName = SrcStr Then CheckEqualPatchNames = True: Exit Function
    Next cntPatchCount
    
    CheckEqualPatchNames = False

End Function


'==========================================================================
' SUB: REFRESH LIST INDEX
' Needed to properly show list after selection reloading.
'==========================================================================
Public Sub RefreshListIndex()

    If NewMainWindow.lbPatchList.ListCount > 0 Then NewMainWindow.lbPatchList.ListIndex = 0

End Sub


'==========================================================================
' FUNCTION: SCAN AND SELECT ITEM
' Scans patches array for desired patch name and selects corresponding list
' item.
'==========================================================================
Public Function ScanAndSelect(Name As String, SelectedFlag As Byte) As Boolean

 Dim cntPatchCount As Integer
 Dim cntListCount As Integer
 
 Dim SelectItem As Boolean
'
    If SelectedFlag = 1 Then SelectItem = True Else SelectItem = False
'
    For cntPatchCount = LBound(PatchArray) To UBound(PatchArray)
    
        If PatchArray(cntPatchCount).patchName = Name Then
            '
            For cntListCount = 0 To NewMainWindow.lbPatchList.ListCount - 1
            
                If NewMainWindow.lbPatchList.List(cntListCount) = Name Then
                
                    NewMainWindow.lbPatchList.Selected(cntListCount) = SelectItem
                    Exit Function
                    
                End If
                
            Next cntListCount
            '
        End If
    '
    Next cntPatchCount

End Function


'==========================================================================
' FUNCTION: PATCH NAME TO NUMBER
' Scans patches array for desired patch name and returns patch array number
' If nothing is found, returns 32767.
'==========================================================================
Public Function PatchNameToNumber(Name As String) As Integer

 Dim cntScan As Integer
'
    For cntScan = LBound(PatchArray) To UBound(PatchArray)
    
        If PatchArray(cntScan).patchName = Name Then
            PatchNameToNumber = cntScan
            Exit Function
        End If
        
    Next cntScan

    PatchNameToNumber = 32767
    
End Function



'==========================================================================
' FUNCTION: PATCH NAME TO PATCH LIST INDEX
' Scans patches array for desired patch name and returns list index number.
' If nothing is found, returns 32767.
'==========================================================================
Public Function PatchNameToListIndex(Name As String) As Integer

 Dim cntScan As Integer
'
    For cntScan = 0 To NewMainWindow.lbPatchList.ListCount - 1
    
        If Name = NewMainWindow.lbPatchList.List(cntScan) Then
            PatchNameToListIndex = cntScan
            Exit Function
        End If
        
    Next cntScan

    PatchNameToListIndex = 32767
    
End Function


Public Function IsPatchSelected(Number As Integer) As Boolean

    IsPatchSelected = NewMainWindow.lbPatchList.Selected(Number)

End Function



'==========================================================================
' FUNCTION: POPULATE PATCH LIST
'==========================================================================
Public Sub PopulatePatchList()

 Dim cntPatchCount As Integer
 Dim cntListCount As Integer
 
    '
    NewMainWindow.lbPatchList.Clear
    
    cntListCount = 0
    
    For cntPatchCount = LBound(PatchArray) To UBound(PatchArray)
    
            NewMainWindow.lbPatchList.List(cntListCount) = PatchArray(cntPatchCount).patchName
            cntListCount = cntListCount + 1
            
    Next cntPatchCount
    
End Sub



'==========================================================================
' FUNCTION: RESET ALL PATCH PARAMETERS TO DEFAULTS.
' Resets all enabled patches' parameters values to defaults.
'==========================================================================
Public Sub ResetAllParameters()

 Dim cntListCount As Integer
 Dim cntParamCount As Integer
'
    For cntListCount = LBound(PatchArray) To UBound(PatchArray)

        For cntParamCount = LBound(PatchArray(cntListCount).patchParams) To UBound(PatchArray(cntListCount).patchParams)
        
            If PatchArray(cntListCount).patchParams(cntParamCount).parEnabled = 1 Then
            
               PatchArray(cntListCount).patchParams(cntParamCount).parModdedValue = _
               PatchArray(cntListCount).patchParams(cntParamCount).parValue
               
            End If
            
        Next cntParamCount

    Next cntListCount
    
End Sub



'==========================================================================
' FUNCTION: DEPENDENCY TEST
' Check selected patch index's dependencies from other patches.
' If failed, returns FALSE.
'==========================================================================
Function DependencyTest(ByVal PatchIndex As Integer) As Boolean

 Dim NeededPatchNames() As String
 
 Dim cntPatchCounter As Integer
 Dim cntDepCounter As Integer
 Dim cntNeededCounter As Integer
 
 
    DependencyTest = True ' Reset to true, until we find non-selected dependency.
 
    If Trim$(PatchArray(PatchIndex).patchDependencies) = vbNullString Then Exit Function
    
    NeededPatchNames = Split(PatchArray(PatchIndex).patchDependencies, kDivider)
    WarnDepString = ""
    

    For cntPatchCounter = LBound(PatchArray) To UBound(PatchArray)
    
        For cntDepCounter = 0 To UBound(NeededPatchNames)
            
            If Trim$(UCase$(NeededPatchNames(cntDepCounter))) = Trim$(UCase$(PatchArray(cntPatchCounter).patchName)) Then
            
                If NewMainWindow.lbPatchList.Selected(cntPatchCounter) = False Then
                
                    cntNeededCounter = cntNeededCounter + 1
                    WarnDepString = WarnDepString + CStr(cntNeededCounter) + ". " + PatchArray(cntPatchCounter).patchName + vbCrLf
                    DependencyTest = False
                
                End If
            
            End If
            
        Next cntDepCounter
    
    Next cntPatchCounter
    
    If LenB(Trim$(WarnDepString)) <> 0 Then WarnDepString = Left$(WarnDepString, Len(WarnDepString) - 2)
    
    
End Function
