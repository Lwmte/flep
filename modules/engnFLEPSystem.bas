Attribute VB_Name = "engnFLEPInterface"
'==========================================================================
'  FLEP 1.0 SYSTEM INTERFACE MODULE CODE
'
'  Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'
'  This module contains all procedures and functions related to
'  NewMainWindow.frm UI events.
'
'=========================================================================
'

Option Explicit

' Config settings.

Public MarkSilently As Byte      ' Silent patch flag.
Public DeveloperMode As Byte     ' Developer mode flag.
Public DeveloperView As Byte     ' Switches devmode view.
Public MaxParameters As Integer  ' Maximum allowed parameters.

Public LastPreset As String
Public DefaultExeName As String  ' Default exe name for patch editor.
Public WindowTitle As String     ' Window title.
Public WindowPositionX As Long   ' Windox X position.
Public WindowPositionY As Long   ' Window Y position.
Public AboutText(2) As String    ' About label text.


' Internal global variables and flags.

Public PrevPatchNumber As Integer         ' Used to save previous patch, if listbox event is invoked.
Public CurrentParamNumber As Integer      ' Keeps current parameter number for user-mode interconnection.

Public WarnDepString As String            ' Dependencies string used in UI info about non-selected dependencies.

Public tempBackupPatch As typPatchEntry   ' Used for backing-up previous version of patch.

Public InputCallbackString As String      'Global string variable for inputbox



' -------------------------- COMMON UI CHAPTER ---------------------------


'==========================================================================
' FUNCTION: SHOW PARAMETER LIST WINDOW
' Shows parameter list and spawns corresponding amount of list entries.
'==========================================================================
Public Function ShowParamListWindow(FillFromArray As Integer)

 Dim cntUnitCounter As Integer
 Dim MaxWidth As Integer
 Dim MaxHeight As Integer
 
 Dim CurWidth, CurHeight As Integer
 
 Dim CursorCoords As POINTAPI
 Dim CursorIndex As Long
 
 
    For cntUnitCounter = LBound(PatchArray(FillFromArray).patchParams) To UBound(PatchArray(FillFromArray).patchParams)
    
        If cntUnitCounter = 0 Then
            ' If it is the first parameter, then simply set its name...
            ParamList.lblParamLink(cntUnitCounter).Caption = PatchArray(FillFromArray).patchParams(cntUnitCounter).parName
            ParamList.lblParamLink(cntUnitCounter).Visible = True
        Else
            ' Else, spawn new parameter list entry!
            Load ParamList.lblParamLink(cntUnitCounter)
            ParamList.lblParamLink(cntUnitCounter).Visible = True
            ParamList.lblParamLink(cntUnitCounter).Top = ParamList.lblParamLink(cntUnitCounter - 1).Top + ParamList.lblParamLink(cntUnitCounter - 1).Height + Fix(ParamList.lblParamLink(cntUnitCounter).Height / 3)
            ParamList.lblParamLink(cntUnitCounter).Left = ParamList.lblParamLink(cntUnitCounter - 1).Left
            ParamList.lblParamLink(cntUnitCounter).Caption = PatchArray(FillFromArray).patchParams(cntUnitCounter).parName
            
            If CurHeight > 4200 Then  ' Max. height is 4200 twips = ~ 17 params (size which fits main FLEP window.)
                CurHeight = 0                                                               ' Reset current height, so new entries spawn at top.
                ParamList.lblParamLink(cntUnitCounter).Top = ParamList.lblParamLink(0).Top  ' Set current param list entry ver. position to first param list entry position.
                ParamList.lblParamLink(cntUnitCounter).Left = CurWidth + 100                ' Set current param list entry hor. position to new column + 100 twips spacing.
            End If

        End If
        
        If (ParamList.lblParamLink(cntUnitCounter).Left + ParamList.lblParamLink(cntUnitCounter).Width) > CurWidth Then CurWidth = (ParamList.lblParamLink(cntUnitCounter).Left + ParamList.lblParamLink(cntUnitCounter).Width)
        CurHeight = CurHeight + ParamList.lblParamLink(cntUnitCounter).Height + Fix(ParamList.lblParamLink(cntUnitCounter).Height / 3)
        
            
        
        If CurWidth > MaxWidth Then MaxWidth = CurWidth
        If CurHeight > MaxHeight Then MaxHeight = CurHeight

    Next cntUnitCounter
    

MaxWidth = MaxWidth + BorderWidth * 2                            ' BorderWidth addition is needed to prevent resize bugs with Vista/7/8.
MaxHeight = MaxHeight + CaptionHeight + Fix(BorderHeight * 1.5)  ' BorderHeight and CaptionHeight addition is needed to prevent resize bugs with Vista/7/8.

CursorIndex = GetCursorPos(CursorCoords)

ParamList.Width = MaxWidth
ParamList.Height = MaxHeight

' Set param window's position under mouse cursor's position.
ParamList.Top = (CursorCoords.y * Screen.TwipsPerPixelY) - ParamList.Height
ParamList.Left = CursorCoords.x * Screen.TwipsPerPixelX

ParamList.Show 1

Call ShowOrHideParams(CurrentParamNumber) ' Finalize list call, when something is selected there.

' This check is needed to prevent description text resizing bug with visible dependencies frame.
If NewMainWindow.frmDependencies.Visible = True Then NewMainWindow.txtPatchDesc.Height = NewMainWindow.txtPatchDesc.Height - NewMainWindow.frmDependencies.Height - 20

End Function



'==========================================================================
' FUNCTION: SHOW INPUT WINDOW
' Shows custom neat input box. Autosizes itself depentding on message and
' max. textbox length. Also automatically snaps to mouse cursor position.
' Returns resulting string.
'==========================================================================
Public Function ShowInputWindow(InputLabel As String, MaxLength As Integer, Optional SendText As String, Optional SelectText As Boolean) As String

Dim CursorCoords As POINTAPI
Dim CursorIndex As Long

    InputKor.lblInputDesc.Caption = InputLabel
    InputKor.Text1.MaxLength = MaxLength
    
    InputKor.Text1.Left = InputKor.lblInputDesc.Left + InputKor.lblInputDesc.Width + 8
    InputKor.Text1.Width = MaxLength * 11
    InputKor.btnOK.Left = InputKor.Text1.Left + InputKor.Text1.Width + 4
    
    InputKor.Width = ((InputKor.btnOK.Left + InputKor.btnOK.Width + 10) * Screen.TwipsPerPixelX) + BorderWidth
    
    CursorIndex = GetCursorPos(CursorCoords)
    
    InputKor.Top = (CursorCoords.y * Screen.TwipsPerPixelY) - InputKor.Height
    InputKor.Left = CursorCoords.x * Screen.TwipsPerPixelX
    
    If LenB(SendText) > 0 Then
     InputKor.Text1.Text = SendText
     InputKor.Text1.SelLength = Len(SendText)
    Else
     InputKor.Text1.Text = vbNullString
    End If
    
    InputKor.Show 1
    
    ShowInputWindow = InputCallbackString
    
End Function



'==========================================================================
' SUB: CLEAR PARAMETERS WINDOW
' Clears parameter settings for forthcoming new param or whatever.
'==========================================================================
Public Sub ClearParamWindow()

If DeveloperMode = 0 And DeveloperView = 0 Then Exit Sub

NewMainWindow.txtParTitle.Text = vbNullString
NewMainWindow.txtParOffset.Text = vbNullString
NewMainWindow.txtParDefault.Text = vbNullString
NewMainWindow.txtParStringLength.Text = vbNullString
NewMainWindow.optParType(3).Value = True
NewMainWindow.optParIgnore(0).Value = True

End Sub


Public Sub UpdateCurrentParam()
PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(CurrentParamNumber).parModdedValue = NewMainWindow.txtParamValue.Text
End Sub


'==========================================================================
' SUB: SHOW DEPENDENCIES MESSAGE BOX
' Show message box with dependencies. Also allows to automatically select
' needed patches.
'==========================================================================
Public Sub WarnDependencies()

    MsgBox "Please enable these custom patches: " & vbCrLf & vbCrLf & WarnDepString, vbExclamation, "Dependencies warning"
    
End Sub


'==========================================================================
' SUB: SWITCH VIEW
' Changes window right panel from user view to dev view and back.
'==========================================================================
Sub SwitchView()

    If DeveloperView = 0 Then
        NewMainWindow.frmDevModeControls.Visible = False
        NewMainWindow.lbPatchList.Height = 369
    Else
        NewMainWindow.frmDevModeControls.Visible = True
        NewMainWindow.lbPatchList.Height = 329
    End If
    
End Sub


'==========================================================================
' SUB: REFRESH VIEW
' Refresh main window components, depending on selected mode.
'==========================================================================
Sub RefreshView()

 Dim tmpCurrentPatch As Integer
 
    tmpCurrentPatch = NewMainWindow.lbPatchList.ListIndex
    
    PrevPatchNumber = tmpCurrentPatch
    
    Call LockWindowUpdate(NewMainWindow.hwnd)

        If DeveloperMode = 1 Then
        
         NewMainWindow.frmUserMode.Visible = False
         NewMainWindow.frmDevMode.Visible = True
         NewMainWindow.btnEditPatch.Caption = "View..."
         
         Call ShowPatchDeveloper(tmpCurrentPatch)
         
        Else
        
         NewMainWindow.frmUserMode.Visible = True
         NewMainWindow.frmDevMode.Visible = False
         NewMainWindow.btnEditPatch.Caption = "Edit..."
         
         Call ShowPatchUser(tmpCurrentPatch)
         
        End If
    
    Call LockWindowUpdate(0)

End Sub


'==========================================================================
' SUB: SHOW OR HIDE PARAMETERS
' Shows or hides parameter frame (USER MODE), depending on persistence of
' parameters in specified patch entry.
'==========================================================================
Sub ShowOrHideParams(SourceParamNumber As Integer)

On Error GoTo ErrorHandler

 Dim Counter As Byte
 Dim FirstEntry As Byte
 Dim CurParamType
 
 Dim ExistFlag As Byte
 
 Dim CurrentPatchNumber As Integer
 
    CurrentPatchNumber = PatchNameToNumber(NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex))
    CurParamType = PatchArray(CurrentPatchNumber).patchParams(SourceParamNumber).parType

    'Reset param frame look to default. Then, in each case, specific param types can modify looks of it.
    NewMainWindow.txtParamValue.Visible = True  'Text input box.
    NewMainWindow.txtParamValue.Width = 2040
    NewMainWindow.picColorTip2.Visible = False  'Color picker for RGB type.
    NewMainWindow.picBitSet8.Visible = False    'Bit set for Bit(8) type.
    NewMainWindow.picBitSet16.Visible = False   'Bit set for Bit(16) type.

    Select Case CurParamType
        Case 0: NewMainWindow.txtParamValue.ForeColor = &H40&: NewMainWindow.txtParamValue.MaxLength = 20       ' Float.
        
        Case 1: NewMainWindow.txtParamValue.ForeColor = vbBlack                                                 ' Bit set (8)
                NewMainWindow.txtParamValue.MaxLength = 4
                NewMainWindow.picBitSet8.Visible = True
                NewMainWindow.txtParamValue.Visible = False
                
        
        Case 2: NewMainWindow.txtParamValue.ForeColor = vbBlack                                                 ' Bit set (16)
                NewMainWindow.txtParamValue.MaxLength = 6
                NewMainWindow.picBitSet16.Visible = True
                NewMainWindow.txtParamValue.Visible = False
                
        
        Case 3: NewMainWindow.txtParamValue.ForeColor = &H40&: NewMainWindow.txtParamValue.MaxLength = 3        ' Byte (unsigned)
        Case 4: NewMainWindow.txtParamValue.ForeColor = &H40&: NewMainWindow.txtParamValue.MaxLength = 5        ' Integer (unsigned)
        Case 5: NewMainWindow.txtParamValue.ForeColor = &HC00000: NewMainWindow.txtParamValue.MaxLength = 4     ' Byte (signed)
        Case 6: NewMainWindow.txtParamValue.ForeColor = &HC00000: NewMainWindow.txtParamValue.MaxLength = 6     ' Integer (signed)
        Case 7: NewMainWindow.txtParamValue.ForeColor = &H40&: NewMainWindow.txtParamValue.MaxLength = 10       ' Long
        Case 8: NewMainWindow.txtParamValue.ForeColor = vbBlack: NewMainWindow.txtParamValue.MaxLength = 1      ' String (default length)
        Case 9: NewMainWindow.txtParamValue.ForeColor = &H40&: NewMainWindow.txtParamValue.MaxLength = 11       ' R,G,B
                NewMainWindow.picColorTip2.Visible = True
                NewMainWindow.txtParamValue.Width = 1760
        Case Else                                                                                               ' String (custom length)
                If CurParamType > 100 Then NewMainWindow.txtParamValue.ForeColor = vbBlack: NewMainWindow.txtParamValue.MaxLength = Fix(CurParamType / 100)
    End Select

    ExistFlag = 0
    FirstEntry = 255  'Reset them just in case...

    For Counter = LBound(PatchArray(CurrentPatchNumber).patchParams) To UBound(PatchArray(CurrentPatchNumber).patchParams)
        If PatchArray(CurrentPatchNumber).patchParams(Counter).parEnabled = 1 Then
        ExistFlag = ExistFlag + 1
        '
        If FirstEntry = 255 Then FirstEntry = Counter ': CurrentParamNumber = Counter
        '
        If SourceParamNumber > 0 Then FirstEntry = SourceParamNumber
        '
            If FirstEntry < 255 Then
                NewMainWindow.lblParamDesc.Caption = PatchArray(CurrentPatchNumber).patchParams(FirstEntry).parName & ":"
                NewMainWindow.txtParamValue.Text = PatchArray(CurrentPatchNumber).patchParams(FirstEntry).parModdedValue
            End If
        End If
    Next Counter
'
'
'
    If ExistFlag = 0 Then  ' If no parameters exist...
        GoTo ErrorHandler
    '
    ElseIf ExistFlag = 1 Then  ' If one parameter exists...
        NewMainWindow.frmParams.Visible = True
        
        NewMainWindow.lblParamDesc.Font.Underline = False
        NewMainWindow.lblParamDesc.ForeColor = &H505050
        NewMainWindow.lblParamDesc.MousePointer = 1
        NewMainWindow.txtPatchDesc.Height = 4660
    '
    ElseIf ExistFlag > 1 Then ' If many parameters exist...
        NewMainWindow.frmParams.Visible = True
        
        NewMainWindow.lblParamDesc.Font.Underline = True
        NewMainWindow.lblParamDesc.ForeColor = &H504040
        NewMainWindow.lblParamDesc.MousePointer = 99
        NewMainWindow.txtPatchDesc.Height = 4660
    '
    End If
    
'
Exit Sub
'
ErrorHandler:

    NewMainWindow.frmParams.Visible = False  ' Disable param frame completely.
    NewMainWindow.txtPatchDesc.Height = 5250 ' Make desc textbox height as whole right panel height.
    
    Exit Sub
'
End Sub


'==========================================================================
' SUB: LOCK DATA BLOCK
' Locks patch data block (DEV VIEW), if there are no patch offsets in list.
'==========================================================================
Sub LockDataBlock()

    Select Case NewMainWindow.lbOffsetList.ListCount
    
        Case Is <= 0:
            NewMainWindow.txtOriginalData.Enabled = False
            NewMainWindow.txtModifiedData.Enabled = False
            NewMainWindow.optPatchIgnore(0).Enabled = False
            NewMainWindow.optPatchIgnore(1).Enabled = False
            NewMainWindow.optPatchIgnore(0).Value = False
            NewMainWindow.optPatchIgnore(1).Value = False
            NewMainWindow.lblDataCondBehave.Enabled = False
            NewMainWindow.txtOriginalData.Text = vbNullString
            NewMainWindow.txtModifiedData.Text = vbNullString
        Case Is > 0:
            NewMainWindow.txtOriginalData.Enabled = True
            NewMainWindow.txtModifiedData.Enabled = True
            NewMainWindow.optPatchIgnore(0).Enabled = True
            NewMainWindow.optPatchIgnore(1).Enabled = True
            NewMainWindow.lblDataCondBehave.Enabled = True
    End Select

End Sub


'==========================================================================
' SUB: LOCK PARAM BLOCK
' Locks param block (DEV VIEW), if there are no params in list.
'==========================================================================
Sub LockParamBlock()

 Dim tmpOptionCounter As Byte
 
    Select Case NewMainWindow.cmbParSlot.ListCount
    
        Case Is <= 0:
            NewMainWindow.txtParOffset.Enabled = False
            NewMainWindow.txtParDefault.Enabled = False
            NewMainWindow.txtParTitle.Enabled = False
            NewMainWindow.txtParStringLength.Enabled = False
            NewMainWindow.lblParTitle.Enabled = False
            NewMainWindow.lblParType.Enabled = False
            NewMainWindow.lblParCondBehave.Enabled = False
            NewMainWindow.lblParOffset.Enabled = False
            NewMainWindow.lblParValue.Enabled = False
            NewMainWindow.txtParOffset.Text = vbNullString
            NewMainWindow.txtParDefault.Text = vbNullString
            NewMainWindow.txtParTitle.Text = vbNullString
            NewMainWindow.txtParStringLength.Text = vbNullString
            
            For tmpOptionCounter = 0 To 9
                NewMainWindow.optParType(tmpOptionCounter).Enabled = False
                NewMainWindow.optParType(tmpOptionCounter).Value = False
            Next
        
            For tmpOptionCounter = 0 To 1
                NewMainWindow.optParIgnore(tmpOptionCounter).Enabled = False
                NewMainWindow.optParIgnore(tmpOptionCounter).Value = False
            Next
            
        Case Is > 0:
            NewMainWindow.txtParOffset.Enabled = True
            NewMainWindow.txtParDefault.Enabled = True
            NewMainWindow.txtParTitle.Enabled = True
            NewMainWindow.txtParStringLength.Enabled = True
            NewMainWindow.lblParTitle.Enabled = True
            NewMainWindow.lblParType.Enabled = True
            NewMainWindow.lblParCondBehave.Enabled = True
            NewMainWindow.lblParOffset.Enabled = True
            NewMainWindow.lblParValue.Enabled = True
            
            For tmpOptionCounter = 0 To 9
                NewMainWindow.optParType(tmpOptionCounter).Enabled = True
            Next
        
            For tmpOptionCounter = 0 To 1
                NewMainWindow.optParIgnore(tmpOptionCounter).Enabled = True
            Next
            
            
    End Select

End Sub

' ----------------------- END OF COMMON UI CHAPTER ------------------------




' -------------- LOAD / SAVE PATCH INTO WINDOW CHAPTER --------------------


'==========================================================================
' SUB: SAVE PATCH SETTINGS
' Updates common patch settings before changing current patch focus.
'==========================================================================
Sub ApplyPatch(DesiredPatchNumber As Integer)

 Dim cntArrayCounter As Integer
 Dim tmpCurrentPatch As Integer
 
 If DesiredPatchNumber = -1 Then Exit Sub
 If DesiredPatchNumber > UBound(PatchArray) Then Exit Sub
 If DeveloperView = 0 Then Exit Sub
 If DeveloperMode = 0 Then Exit Sub
 
    If NewMainWindow.lbOffsetList.ListCount > 0 Then Call SaveOffset(DesiredPatchNumber, NewMainWindow.lbOffsetList.ListIndex)
    If NewMainWindow.cmbParSlot.ListCount > 0 Then Call SaveParam(DesiredPatchNumber, NewMainWindow.cmbParSlot.ListIndex)

    PatchArray(DesiredPatchNumber).patchName = NewMainWindow.txtEditPatchName.Text
    PatchArray(DesiredPatchNumber).patchDesc = CipherText(NewMainWindow.txtEditPatchDesc.Text)
    PatchArray(DesiredPatchNumber).patchFile = NewMainWindow.txtTargetFile.Text
    '
    PatchArray(DesiredPatchNumber).patchDependencies = NewMainWindow.txtEditPatchDependencies.Text
    '
    ' This is needed to prevent listbox from reacting on list change event.
    If NewMainWindow.lbPatchList.List(DesiredPatchNumber) <> NewMainWindow.txtEditPatchName.Text Then NewMainWindow.lbPatchList.List(DesiredPatchNumber) = NewMainWindow.txtEditPatchName.Text
    '
End Sub


'==========================================================================
' SUB: SHOW PATCH (DEVELOPER VIEW)
' Shows patch frame in developer view.
'==========================================================================
Sub ShowPatchDeveloper(PatchArrayIndex As Integer)

 Dim cntUnitCounter As Integer

    If PatchArrayIndex >= 32767 Or PatchArrayIndex <= -1 Then Exit Sub

    Call BackupPatch(PatchArrayIndex) ' Backup patch into temp entry to make Undo function work.
    
    NewMainWindow.lbOffsetList.Clear
    
    NewMainWindow.txtEditPatchName.Text = PatchArray(PatchArrayIndex).patchName
    NewMainWindow.txtEditPatchDesc.Text = DecipherText(PatchArray(PatchArrayIndex).patchDesc)
    NewMainWindow.txtEditPatchDependencies.Text = PatchArray(PatchArrayIndex).patchDependencies
    NewMainWindow.txtTargetFile.Text = PatchArray(PatchArrayIndex).patchFile
    
    Call ClearParamWindow   ' Reset param frame options to default.
    
    Call PopulateOffsetList ' Populate datas / params lists.
    Call PopulateParamList
    
    Call LockDataBlock   ' Lock / unlock data and param frames, depending on amount of datas / params.
    Call LockParamBlock

End Sub


'==========================================================================
' SUB: SHOW PATCH (USER VIEW)
' Shows patch frame in user view.
'==========================================================================
Sub ShowPatchUser(PatchArrayIndex As Integer)

 Dim cntDescLineCount As Integer
    
    If PatchArrayIndex = 32767 Or PatchArrayIndex = -1 Then Exit Sub
    
    '
    ' change labels
    NewMainWindow.lblPatchTitle.Caption = PatchArray(PatchArrayIndex).patchName
    NewMainWindow.txtPatchDesc.Text = DecipherText(PatchArray(PatchArrayIndex).patchDesc)
    '
    '
    CurrentParamNumber = 0
    
    Call ShowOrHideParams(0)
    Call UpdColor(NewMainWindow.txtParamValue.Text)
    Call ShowOrHideDependencies(PatchArrayIndex)

End Sub


'==========================================================================
' SUB: SHOW OR HIDE DEPENDENCIES INFOBOX
' If dependency test fails, shows dependency infobox and fixes other window
' components' positions.
'==========================================================================
Sub ShowOrHideDependencies(PatchArrayIndex As Integer)
    
        If DependencyTest(PatchArrayIndex) = False Then
    
        NewMainWindow.frmDependencies.Visible = True
        NewMainWindow.txtDependenciesList.Text = WarnDepString
        NewMainWindow.txtDependenciesList.Height = (GetLineCount(NewMainWindow.txtDependenciesList.hwnd)) * 205
        NewMainWindow.lblDepEnableNeeded.Top = NewMainWindow.txtDependenciesList.Top + NewMainWindow.txtDependenciesList.Height
        NewMainWindow.frmDependencies.Height = NewMainWindow.lblDepEnableNeeded.Top + NewMainWindow.lblDepEnableNeeded.Height + 90
        
        NewMainWindow.txtPatchDesc.Height = NewMainWindow.txtPatchDesc.Height - NewMainWindow.frmDependencies.Height - 20
        NewMainWindow.frmDependencies.Top = NewMainWindow.txtPatchDesc.Top + NewMainWindow.txtPatchDesc.Height + 20
        
    Else
        NewMainWindow.frmDependencies.Visible = False
    End If


End Sub

' ----------- END OF LOAD / SAVE PATCH INTO WINDOW CHAPTER ----------------




' ------------------- DEV MODE PARAM SETTINGS CHAPTER ---------------------

'==========================================================================
' SUB: ADD NEW PARAMETER
' Adds new parameter both to list and patch array.
'==========================================================================
Sub AddParam()
    
    NewMainWindow.cmbParSlot.AddItem k_ParamPrefix + CStr(NewMainWindow.cmbParSlot.ListCount), NewMainWindow.cmbParSlot.ListCount
    
    ReDim Preserve PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(NewMainWindow.cmbParSlot.ListCount - 1)
    PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(NewMainWindow.cmbParSlot.ListCount - 1).parEnabled = 1
    
    NewMainWindow.cmbParSlot.ListIndex = NewMainWindow.cmbParSlot.ListCount - 1
    
    ClearParamWindow

End Sub


'==========================================================================
' SUB: SAVE PARAMETER
' Saves parameter from window into patch array.
'==========================================================================
Sub SaveParam(PatchNumber As Integer, ParNumber As Integer)

On Error GoTo ErrorHandler

 Dim cntOptionCounter As Byte
 
    If ParNumber > UBound(PatchArray(PatchNumber).patchParams) Then Exit Sub
    If ParNumber < 0 Then Exit Sub

    PatchArray(PatchNumber).patchParams(ParNumber).parEnabled = 1
    PatchArray(PatchNumber).patchParams(ParNumber).parName = NewMainWindow.txtParTitle.Text
    PatchArray(PatchNumber).patchParams(ParNumber).parValue = NewMainWindow.txtParDefault.Text
    PatchArray(PatchNumber).patchParams(ParNumber).parOffset = NewMainWindow.txtParOffset.Text
    
    'Modded value also gets filled, if it's empty (for ex., new patch).
    If PatchArray(PatchNumber).patchParams(ParNumber).parModdedValue = vbNullString Then _
    PatchArray(PatchNumber).patchParams(ParNumber).parModdedValue = PatchArray(PatchNumber).patchParams(ParNumber).parValue
    
    For cntOptionCounter = 0 To 1

        If NewMainWindow.optParIgnore(cntOptionCounter).Value = True Then PatchArray(PatchNumber).patchParams(ParNumber).parCondBehave = cntOptionCounter
    
    Next cntOptionCounter
    
    
    For cntOptionCounter = 0 To 9
    
        If NewMainWindow.optParType(cntOptionCounter).Value = True Then
            
            Select Case cntOptionCounter
                              
                Case 8
                    PatchArray(PatchNumber).patchParams(ParNumber).parType = cntOptionCounter + Val(NewMainWindow.txtParStringLength.Text) * 100

                Case Else
                    PatchArray(PatchNumber).patchParams(ParNumber).parType = cntOptionCounter
                
            End Select
            
        End If
            
    Next cntOptionCounter
    
    Exit Sub
    
ErrorHandler:
    
    Exit Sub
    
End Sub



'==========================================================================
' SUB: LOAD PARAMETER
' Loads parameter from patch array's parameter sub-array.
'==========================================================================
Sub LoadParam(ParNumber As Integer)

On Error GoTo ErrorHandler

 Dim cntOptionCounter As Byte
 
    If ParNumber > UBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams) Then Exit Sub
    If ParNumber < 0 Then Exit Sub

    NewMainWindow.txtParTitle.Text = PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parName
    NewMainWindow.txtParDefault.Text = PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parValue
    NewMainWindow.txtParOffset.Text = PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parOffset
    
    For cntOptionCounter = 0 To 1

        If PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parCondBehave = cntOptionCounter Then NewMainWindow.optParIgnore(cntOptionCounter).Value = True
    
    Next cntOptionCounter
    
    
    Select Case PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parType
    
        Case 0 To 7
            NewMainWindow.optParType(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parType).Value = True
            NewMainWindow.txtParStringLength.Text = vbNullString
        Case 8
            NewMainWindow.optParType(8).Value = True
            NewMainWindow.txtParStringLength.Text = "1"
        Case 9:
            NewMainWindow.optParType(9).Value = True
        Case Is > 100
            NewMainWindow.optParType(8).Value = True
            NewMainWindow.txtParStringLength.Text = CStr(Fix(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(ParNumber).parType / 100))
            
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "something's wrong."

End Sub



'==========================================================================
' SUB: DESTROY PARAMETER
' Destroys selected parameter from patch array's parameter sub-array.
'==========================================================================
Sub DestroyParam(ParNumber As Integer)

On Error GoTo ErrorHandler

 Dim BackupIndex As Integer
 Dim cntArraySize As Integer
 Dim cntArrayCounter As Integer

    If NewMainWindow.cmbParSlot.ListCount = 0 Then Exit Sub
    If ParNumber < 0 Or ParNumber >= 32767 Then Exit Sub
    
    BackupIndex = NewMainWindow.cmbParSlot.ListIndex
    
    NewMainWindow.cmbParSlot.RemoveItem ParNumber
    
    cntArraySize = UBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams)
    
    For cntArrayCounter = ParNumber To cntArraySize - 1
        PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(cntArrayCounter) = PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(cntArrayCounter + 1)
    Next cntArrayCounter
    
    If cntArraySize > 0 Then
        ReDim Preserve PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams(cntArraySize - 1)
    Else
        Erase PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams
    End If
        
        
    RefreshParamList
        
    If BackupIndex < NewMainWindow.cmbParSlot.ListCount Then
        NewMainWindow.cmbParSlot.ListIndex = BackupIndex
    Else
        NewMainWindow.cmbParSlot.ListIndex = NewMainWindow.cmbParSlot.ListCount - 1
    End If
    
    
ErrorHandler:
    Exit Sub

End Sub


'==========================================================================
' SUB: POPULATE PARAMETER LIST
' Populates param list with corresponding array entries.
'==========================================================================
Sub PopulateParamList()

On Error GoTo ErrorHandler

 Dim cntUnitCounter As Integer
 
    NewMainWindow.cmbParSlot.Clear

    For cntUnitCounter = 0 To UBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchParams)
        NewMainWindow.cmbParSlot.AddItem "Parameter " & CStr(cntUnitCounter), cntUnitCounter
    Next cntUnitCounter
    
    NewMainWindow.cmbParSlot.ListIndex = 0
    Call LoadParam(0)
    
    Exit Sub

ErrorHandler:
    Exit Sub
    
End Sub


'==========================================================================
' SUB: REFRESH PARAM LIST
' Refreshes param list names.
'==========================================================================

Sub RefreshParamList()

 Dim cntListCounter As Integer
    
    If NewMainWindow.cmbParSlot.ListCount <= 0 Then Exit Sub
        
    For cntListCounter = 0 To NewMainWindow.cmbParSlot.ListCount - 1
    
        NewMainWindow.cmbParSlot.List(cntListCounter) = k_ParamPrefix & CStr(cntListCounter)
        
    Next cntListCounter
    
End Sub

' ---------------- END OF DEV MODE PARAM SETTINGS CHAPTER ------------------



' ------------------- DEV MODE DATAS SETTINGS CHAPTER ----------------------

Sub AddOffset()

 Dim tmpInputString As String

    tmpInputString = ShowInputWindow(k_IB_AddOffset, 8)
        
    If tmpInputString <> kNullStr Then
            
        NewMainWindow.lbOffsetList.AddItem tmpInputString
        ReDim Preserve PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(NewMainWindow.lbOffsetList.ListCount - 1)
        PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(NewMainWindow.lbOffsetList.ListCount - 1).dataOffset = tmpInputString
        NewMainWindow.lbOffsetList.ListIndex = NewMainWindow.lbOffsetList.ListCount - 1
    
    End If

End Sub



Sub EditOffset()

 Dim tmpInputString As String
 
    tmpInputString = ShowInputWindow(k_IB_EditOffset, 8, PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(NewMainWindow.lbOffsetList.ListIndex).dataOffset)
        
    If tmpInputString <> kNullStr Then
            
        NewMainWindow.lbOffsetList.List(NewMainWindow.lbOffsetList.ListIndex) = tmpInputString
        PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(NewMainWindow.lbOffsetList.ListIndex).dataOffset = tmpInputString
    
    End If

End Sub



Sub SaveOffset(PatchNumber As Integer, OffsetNumber As Integer)

 Dim cntOptionCounter As Byte

    If OffsetNumber > UBound(PatchArray(PatchNumber).patchDatas) Then Exit Sub
    If OffsetNumber < 0 Then Exit Sub

    PatchArray(PatchNumber).patchDatas(OffsetNumber).dataDefault = CipherText(NewMainWindow.txtOriginalData.Text)
    PatchArray(PatchNumber).patchDatas(OffsetNumber).dataModified = CipherText(NewMainWindow.txtModifiedData.Text)
    
    For cntOptionCounter = 0 To 1

        If NewMainWindow.optPatchIgnore(cntOptionCounter).Value = True Then PatchArray(PatchNumber).patchDatas(OffsetNumber).dataCondBehave = cntOptionCounter
    
    Next cntOptionCounter
      
End Sub



Sub LoadOffset(OffsetNumber As Integer)

 Dim cntOptionCounter As Byte

    If OffsetNumber > UBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas) Then Exit Sub
    If OffsetNumber < 0 Then Exit Sub

    NewMainWindow.txtOriginalData.Text = DecipherText(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(OffsetNumber).dataDefault)
    NewMainWindow.txtModifiedData.Text = DecipherText(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(OffsetNumber).dataModified)
    NewMainWindow.optPatchIgnore(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(OffsetNumber).dataCondBehave).Value = True

End Sub



Sub DestroyOffset(OffsetNumber As Integer)

On Error GoTo ErrorHandler

Dim BackupIndex As Integer
Dim cntArraySize As Integer
Dim cntArrayCounter As Integer

    If NewMainWindow.lbOffsetList.ListCount = 0 Then Exit Sub
    If OffsetNumber < 0 Or OffsetNumber >= 32767 Then Exit Sub
    
    BackupIndex = NewMainWindow.lbOffsetList.ListIndex
    
    NewMainWindow.lbOffsetList.RemoveItem OffsetNumber
    
    cntArraySize = UBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas)
    
    For cntArrayCounter = OffsetNumber To cntArraySize - 1
        PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(cntArrayCounter) = PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(cntArrayCounter + 1)
    Next cntArrayCounter
    
    
    If cntArraySize > 0 Then
        ReDim Preserve PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(cntArraySize - 1)
    Else
        Erase PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas
    End If

        
    If BackupIndex < NewMainWindow.lbOffsetList.ListCount Then
        NewMainWindow.lbOffsetList.ListIndex = BackupIndex
    Else
        NewMainWindow.lbOffsetList.ListIndex = NewMainWindow.lbOffsetList.ListCount - 1
    End If
    
    
ErrorHandler:
    Exit Sub
    
End Sub



Sub PopulateOffsetList()

On Error GoTo ErrorHandler

 Dim cntUnitCounter As Integer
 
    NewMainWindow.lbOffsetList.Clear

    For cntUnitCounter = LBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas) To UBound(PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas)
        NewMainWindow.lbOffsetList.AddItem PatchArray(NewMainWindow.lbPatchList.ListIndex).patchDatas(cntUnitCounter).dataOffset, cntUnitCounter
    Next cntUnitCounter

    NewMainWindow.lbOffsetList.ListIndex = 0
        
ErrorHandler:
    
    Exit Sub
    
End Sub

' ---------------- END OF DEV MODE DATAS SETTINGS CHAPTER -------------------




' ----------------------------- MISC CHAPTER --------------------------------

 'SEARCH FUNCTION
Function SearchForPatch(SearchString As String, SearchFromBeginning As Boolean) As Boolean
'
Dim Gobo As Integer  ' Got to correct these names, original TREP ones are confusing.
Dim Karus As Integer
Dim Reason As Integer
'
If SearchString = vbNullString Then Exit Function 'Kill if string is empty
'
SearchForPatch = True
NewMainWindow.txtSearchFailed.Visible = False: NewMainWindow.tmrTipTimer.Enabled = False
Reason = NewMainWindow.lbPatchList.ListIndex
'
If SearchFromBeginning = True Then
Reason = 0
'FromBeginning = False
End If
'
For Gobo = LBound(PatchArray) To UBound(PatchArray)
'
    If Len(PatchArray(Gobo).patchName) >= Len(SearchString) And SearchString <> vbNullString Then
        '
        For Karus = 1 To (Len(PatchArray(Gobo).patchName) - Len(SearchString)) + 1
            If UCase$(Mid$(PatchArray(Gobo).patchName, Karus, Len(SearchString))) = UCase$(SearchString) Then
            If Gobo > Reason Or SearchFromBeginning = True Then
                NewMainWindow.lbPatchList.ListIndex = Gobo
                Exit Function
            End If
            End If
        Next Karus
        '
    End If
'
Next Gobo
'
 SearchForPatch = False
 NewMainWindow.txtSearchFailed.Visible = True: NewMainWindow.tmrTipTimer.Enabled = True ' NOT FOUND effect...
'
End Function



Public Sub MovePatchDown()

 Dim TempCopy As typPatchEntry
 Dim TempName As String
 Dim TempFlag, TempFlag2 As Boolean
 
 Dim BackupIndex As Integer
'
    If NewMainWindow.lbPatchList.ListCount > 1 And NewMainWindow.lbPatchList.ListIndex < (NewMainWindow.lbPatchList.ListCount - 1) Then
      
        'copy previous patch data, name and selection into temp variables
        TempCopy = PatchArray(NewMainWindow.lbPatchList.ListIndex)
        TempName = NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex)
        TempFlag = NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex)
        TempFlag2 = NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex + 1)
        
        'back-up position
        BackupIndex = NewMainWindow.lbPatchList.ListIndex + 1
        
        'reset selection to prevent VB bug
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex) = False
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex + 1) = False
        
        'do data swap
        PatchArray(NewMainWindow.lbPatchList.ListIndex) = PatchArray(NewMainWindow.lbPatchList.ListIndex + 1)
        PatchArray(NewMainWindow.lbPatchList.ListIndex + 1) = TempCopy
        
        'do name swap
        NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex) = NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex + 1)
        NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex + 1) = TempName
        
        'update everything
        Call RefreshView

        'do selection swap
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex) = TempFlag2
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex + 1) = TempFlag
        
        'restore position
        NewMainWindow.lbPatchList.ListIndex = BackupIndex
                
    End If


End Sub



Public Sub MovePatchUp()
    
 Dim TempCopy As typPatchEntry
 Dim TempName As String
 Dim TempFlag, TempFlag2 As Boolean
 
 Dim BackupIndex As Integer
'
    If NewMainWindow.lbPatchList.ListCount > 1 And NewMainWindow.lbPatchList.ListIndex > 0 Then
        
        'copy previous patch data, name and selection into temp variables
        TempCopy = PatchArray(NewMainWindow.lbPatchList.ListIndex)
        TempName = NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex)
        TempFlag = NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex)
        TempFlag2 = NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex - 1)
        
        'back-up position
        BackupIndex = NewMainWindow.lbPatchList.ListIndex - 1
        
        'reset selection to prevent VB bug
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex) = False
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex - 1) = False
        
        'do data swap
        PatchArray(NewMainWindow.lbPatchList.ListIndex) = PatchArray(NewMainWindow.lbPatchList.ListIndex - 1)
        PatchArray(NewMainWindow.lbPatchList.ListIndex - 1) = TempCopy
        
        'do name swap
        NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex) = NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex - 1)
        NewMainWindow.lbPatchList.List(NewMainWindow.lbPatchList.ListIndex - 1) = TempName
        
        'update everything
        Call RefreshView

        'do selection swap
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex) = TempFlag2
        NewMainWindow.lbPatchList.Selected(NewMainWindow.lbPatchList.ListIndex - 1) = TempFlag
        
        'restore position
        NewMainWindow.lbPatchList.ListIndex = BackupIndex
                
    End If



End Sub



Public Sub DeletePatch(PatchNumber As Integer)

 Dim BackupIndex As Byte
 Dim cntArraySize, cntArrayCounter As Integer
 
    If PatchNumber < 0 Then Exit Sub

    If NewMainWindow.lbPatchList.ListIndex >= 0 Then BackupIndex = NewMainWindow.lbPatchList.ListIndex
    
    If NewMainWindow.lbPatchList.ListCount > 1 Then NewMainWindow.lbPatchList.RemoveItem PatchNumber
    
    cntArraySize = UBound(PatchArray)
    
    For cntArrayCounter = PatchNumber To cntArraySize - 1
        PatchArray(cntArrayCounter) = PatchArray(cntArrayCounter + 1)
    Next cntArrayCounter
    
    
    If cntArraySize > 0 Then
        ReDim Preserve PatchArray(cntArraySize - 1)
    Else
        ResetArrays
    End If
    
    RefreshView
    If NewMainWindow.lbPatchList.ListCount = 1 Then PopulatePatchList
    If NewMainWindow.lbPatchList.ListCount >= BackupIndex + 1 Then NewMainWindow.lbPatchList.ListIndex = BackupIndex ' Else List1.ListIndex = 0
    If NewMainWindow.lbPatchList.ListCount < BackupIndex + 1 And NewMainWindow.lbPatchList.ListCount >= 1 Then NewMainWindow.lbPatchList.ListIndex = NewMainWindow.lbPatchList.ListCount - 1

End Sub



Public Sub AddNewPatch(DoCopy As Boolean)

    Dim CurrMassiveAmount As Integer
    Dim tmpPatchName As String
    
    Dim cntNameCounter As Integer
    '
    
    CurrMassiveAmount = UBound(PatchArray) + 1
    '
    If CurrMassiveAmount < 32767 Then
    
        If DoCopy = False Then
    
            tmpPatchName = FindFreePatchName
            ReDim Preserve PatchArray(CurrMassiveAmount)
            PatchArray(CurrMassiveAmount).patchName = tmpPatchName
            PatchArray(CurrMassiveAmount).PatchEnabled = 1
            NewMainWindow.lbPatchList.AddItem tmpPatchName
        
        Else
        
            If NewMainWindow.lbPatchList.ListCount = 0 Or NewMainWindow.lbPatchList.ListIndex < 0 Then Exit Sub
        
            ReDim Preserve PatchArray(CurrMassiveAmount)
            PatchArray(CurrMassiveAmount) = PatchArray(NewMainWindow.lbPatchList.ListIndex)
            PatchArray(CurrMassiveAmount).patchName = PatchArray(CurrMassiveAmount).patchName & k_CopyPostfix
            PatchArray(CurrMassiveAmount).PatchEnabled = 1
            NewMainWindow.lbPatchList.AddItem PatchArray(CurrMassiveAmount).patchName
        
        End If
        
        NewMainWindow.lbPatchList.ListIndex = NewMainWindow.lbPatchList.ListCount - 1
    
        RefreshView

    End If

End Sub


Public Function FindFreePatchName() As String

Dim NameFound As Boolean
Dim NameIndex As Integer
Dim tmpString As String
Dim cntNameCount As Integer

NameFound = False

    Do Until NameFound = True

        NameIndex = NameIndex + 1
        tmpString = kNewPatchPrefix & NameIndex
            For cntNameCount = 0 To NewMainWindow.lbPatchList.ListCount
            If tmpString = NewMainWindow.lbPatchList.List(cntNameCount) Then NameFound = False: Exit For
            NameFound = True
            Next cntNameCount

    Loop
    
    FindFreePatchName = tmpString
    
End Function



Public Sub BackupPatch(PatchNumber As Integer)
    If PatchNumber = -1 Then Exit Sub
        tempBackupPatch = PatchArray(PatchNumber)
End Sub



Public Sub RestorePatch(PatchNumber As Integer)
    If PatchNumber = -1 Then Exit Sub
       PatchArray(PatchNumber) = tempBackupPatch
       RefreshView
End Sub


' CLICK PATCH LIST - CALLED EVERY TIME WE NEED TO INVOKE CLICK PATCH LIST EVENT.
Public Sub ClickPatchList()

If DeveloperMode = 1 And DeveloperView = 1 Then
    Call ApplyPatch(PrevPatchNumber)
    Call BackupPatch(NewMainWindow.lbPatchList.ListIndex)
End If

    Call RefreshView

End Sub


'LOAD CONFIG FILE - CALLED BEFORE COMMON INITIALIZATION
Public Sub LoadCfg()

    If LoadPreset(ConfigPath, True, False) = False Then
        Call RecreateConfigFile(ConfigPath)
        If LoadPreset(ConfigPath, True, False) = False Then MsgBox "Fatal error during config recreation.": Unload NewMainWindow
    End If

End Sub


'LOAD PATCH PRESETS FROM CONFIG FILE - CALLED AFTER COMMON INITIALIZATION.
Public Sub LoadPresetsFromCfg()

If LoadPreset(ConfigPath, False, True) = False Then Exit Sub ' ?

End Sub


'GENERAL LOAD CONFIG OR PRESET FILE FUNCTION.
Public Function LoadPreset(PresetName As String, LoadSystem As Boolean, LoadPresets As Boolean) As Boolean

On Error GoTo ErrorHandler

 Dim tmpString As String
 Dim tmpStringArray() As String
 
 Dim tmpCfg As typCfgParam
 
 Dim tmpSuccessFlag As Boolean
 
 Dim cntUnitCounter As Integer
 
 Dim HeaderFound As Boolean
 Dim PresetVersion As String
 

    LoadPreset = False  ' just in case
    HeaderFound = False
    
    If CheckExistence(PresetName) = False Then
        Exit Function
    End If
    
    
    Call LockWindowUpdate(NewMainWindow.hwnd)
 
    
        Open PresetName For Input As #1
        
        Do Until EOF(1)
        
            Line Input #1, tmpString
            tmpCfg = ConvCFG(tmpString)
            
            If LCase$(tmpCfg.Name) = kPresetVersion Then
            
                PresetVersion = tmpCfg.Value
                
                If Val(PresetVersion) <= Val(k_HeaderVersion) Then
                    HeaderFound = True
                Else
                    Call MsgBox("Warning: specified preset / config has higher version number." & vbCrLf & "Please update FLEP program!", vbCritical)
                    Call LockWindowUpdate(0)
                    Close #1
                    Exit Function
                End If
                
            End If
            
            If HeaderFound = True Then
            
                If LoadSystem = True Then
                
                    Select Case LCase$(tmpCfg.Name)
                    
                        Case kSilentPatch: MarkSilently = CByteL(tmpCfg.Value)
                        Case kDeveloperMode: DeveloperMode = CByteL(tmpCfg.Value)
                        Case kDeveloperView: DeveloperView = CByteL(tmpCfg.Value)
                        Case kDefaultPatchSet: CurrentSetName = tmpCfg.Value
                        Case kMaxParameters: MaxParameters = CInt(tmpCfg.Value)
                        Case kDefaultExeName: DefaultExeName = tmpCfg.Value
                        Case kLastPreset: LastPreset = tmpCfg.Value
                        Case kWindowTitle: WindowTitle = tmpCfg.Value
                        Case kAboutText1: AboutText(0) = DecipherText(tmpCfg.Value)
                        Case kAboutText2: AboutText(1) = DecipherText(tmpCfg.Value)
                        Case kAboutText3: AboutText(2) = DecipherText(tmpCfg.Value)
                        
                        Case kWindowPosX: WindowPositionX = CLng(tmpCfg.Value)
                        Case kWindowPosY: WindowPositionY = CLng(tmpCfg.Value)
                        
                        Case kLastFile ' checksum
                            tmpStringArray = Split(tmpCfg.Value, kDivider, 2)
                            ReDim Preserve tmpStringArray(1)
                            tmpSuccessFlag = UpdateExecList(tmpStringArray(0), HxVal(tmpStringArray(1)))
                            Erase tmpStringArray
                            
                    End Select
                    
                End If
                
                
                If LoadPresets = True And LCase$(tmpCfg.Name) = kPatchPrefix Then
                        
                        tmpStringArray = Split(tmpCfg.Value, kDivider, 3)
                        ReDim Preserve tmpStringArray(2)
                        tmpSuccessFlag = ConvertPatchPreset(tmpStringArray(2), PatchNameToNumber(tmpStringArray(0)))
                        tmpSuccessFlag = ScanAndSelect(tmpStringArray(0), CByteL(tmpStringArray(1)))
                        Erase tmpStringArray
    
                End If
    
            End If
    
        Loop
    
        Close #1
    
        RefreshListIndex

    Call LockWindowUpdate(0)

    LoadPreset = True
    Exit Function

ErrorHandler:
    Close #1
    Call LockWindowUpdate(0)
    MsgBox "Error loading config / preset file." & vbCrLf & "Reason: Error #" & CStr(Err.Number) & " (" & CStr(Err.Description) & ").", vbExclamation, "Cannot load config" 'alarm user"

End Function



'SAVE CONFIG OR PRESET FILE COMMON FUNCTION
Public Function SavePreset(PresetName As String, SaveSystem As Boolean, SavePresets As Boolean) As Boolean

On Error GoTo ErrorHandler

 Dim cntUnitCounter As Integer
 Dim tmpPresetString As String
 Dim tmpParamsString As String
 
    SavePreset = False  ' just in case
        
    Open PresetName For Output As #1
    
        Print #1, WriteSetting(kPresetVersion, k_HeaderVersion, 0, 2)
        
        If SaveSystem = True Then
        
            Print #1, WriteSetting(kSilentPatch, MarkSilently, 1)
            Print #1, WriteSetting(kDeveloperMode, DeveloperMode, 1)
            Print #1, WriteSetting(kDeveloperView, DeveloperView, 1)
            Print #1, WriteSetting(kDefaultPatchSet, CurrentSetName, 1)
            Print #1, WriteSetting(kMaxParameters, MaxParameters, 1)
            Print #1, WriteSetting(kDefaultExeName, DefaultExeName, 1)
            Print #1, WriteSetting(kLastPreset, LastPreset, 1, 1)
            
            Print #1, WriteSetting(kWindowPosX, NewMainWindow.Left, 1)
            Print #1, WriteSetting(kWindowPosY, NewMainWindow.Top, 1, 1)
            
            Print #1, WriteSetting(kWindowTitle, WindowTitle, 1)
            Print #1, WriteSetting(kAboutText1, CipherText(AboutText(0)), 1)
            Print #1, WriteSetting(kAboutText2, CipherText(AboutText(1)), 1)
            Print #1, WriteSetting(kAboutText3, CipherText(AboutText(2)), 1, 2)
        
        End If
        
        
        If SavePresets = True Then  ' If only system options are written (config file)...
        
            On Error Resume Next
        
            For cntUnitCounter = LBound(ExecList) To UBound(ExecList)
            If Err.Number = 9 Then Err.Clear: Exit For
            
                 Print #1, WriteSetting(kLastFile, ExecList(cntUnitCounter).execName + kDivider + CStr(Hex(ExecList(cntUnitCounter).execLastCRC)), 1, 1)
        
            Next cntUnitCounter
            
            
            For cntUnitCounter = LBound(PatchArray) To UBound(PatchArray)
            If Err.Number = 9 Then Err.Clear: Exit For
            
                 ' patch preset string. Format: Name,Enabled,[parameter values if exist,...,...]
                 
                 tmpPresetString = PatchArray(cntUnitCounter).patchName & kDivider & CStr(Abs(NewMainWindow.lbPatchList.Selected(cntUnitCounter)))
                 tmpParamsString = MergeModdedValues(cntUnitCounter)
                 
                 If LenB(tmpParamsString) > 0 Then tmpPresetString = tmpPresetString & kDivider & tmpParamsString
                 
                 Print #1, WriteSetting(kPatchPrefix, tmpPresetString, 1)
            
            Next cntUnitCounter
        
        End If
            
            
        On Error GoTo ErrorHandler
    
        Close #1
    
        SavePreset = True
        
        Exit Function
    
ErrorHandler:
    Close #1
    SavePreset = False
    MsgBox "Error saving config / preset file. Check integrity."

End Function


' COMMON CONFIG PARAMETERS APPLYING ROUTINE
Sub ApplyParameters()

If DeveloperView = 1 Then SwitchView

NewMainWindow.Caption = WindowTitle & (" v. " & App.Major & "." & App.Minor & "." & App.Revision)

If WindowPositionY > 0 Then NewMainWindow.Top = WindowPositionY
If WindowPositionX > 0 Then NewMainWindow.Left = WindowPositionX

End Sub


' RESET ALL PATCH SELECTIONS AND SETTINGS TO DEFAULTS.
Sub ResetList()

On Error Resume Next

 Dim cntPatches As Integer
 Dim cntParams As Integer


    For cntPatches = 0 To (NewMainWindow.lbPatchList.ListCount - 1)
    
        NewMainWindow.lbPatchList.Selected(cntPatches) = False
    
        For cntParams = 0 To UBound(PatchArray(cntPatches).patchParams)
        If Err.Number = 9 Then Err.Clear: Exit For
        
            PatchArray(cntPatches).patchParams(cntParams).parModdedValue = PatchArray(cntPatches).patchParams(cntParams).parValue
        
        Next cntParams
        
    Next cntPatches
    
RefreshView

End Sub



'==========================================================================
' FUNCTION: DEPENDENCY FIXER
' Automatically enables all needed dependencies.
'==========================================================================
Public Sub FixDependencies(ByVal PatchIndex As Integer)

 Dim NeededPatchNames() As String
 
 Dim cntPatchCounter As Integer
 Dim cntDepCounter As Integer
 
 
    If Trim$(PatchArray(PatchIndex).patchDependencies) = vbNullString Then Exit Sub
    
    NeededPatchNames = Split(PatchArray(PatchIndex).patchDependencies, kDivider)
    

    For cntPatchCounter = LBound(PatchArray) To UBound(PatchArray)
    
        For cntDepCounter = 0 To UBound(NeededPatchNames)
            
            If Trim$(UCase$(NeededPatchNames(cntDepCounter))) = Trim$(UCase$(PatchArray(cntPatchCounter).patchName)) Then
            
                NewMainWindow.lbPatchList.Selected(cntPatchCounter) = True
            
            End If
            
        Next cntDepCounter
    
    Next cntPatchCounter
    
    RefreshView
    NewMainWindow.lbPatchList.ListIndex = PatchIndex
    
End Sub



'==========================================================================
' FUNCTION: COLOR HELPER
' Shows color that is currently set in param box.
'==========================================================================
Public Sub UpdColor(ColorData As String)

If NewMainWindow.frmParams.Visible = False Then Exit Sub
If NewMainWindow.picColorTip2.Visible = False Then Exit Sub

 Dim tempString() As String
 Dim RedStep, BlueStep, GreenStep As Byte

    tempString = Split(ColorData, kDivider, 3)
    If UBound(tempString) < 2 Then ReDim Preserve tempString(2)
    
    RedStep = Fix(CByteL(Abs(Val(tempString(0)))))
    GreenStep = Fix(CByteL(Abs(Val(tempString(1)))))
    BlueStep = Fix(CByteL(Abs(Val(tempString(2)))))
    
    NewMainWindow.picColorTip2.BackColor = RGB(RedStep, GreenStep, BlueStep)
    
End Sub


'==========================================================================
' FUNCTION: UPDATE BITS
' Converts bits labels into decimal value and puts it into textbox.
'==========================================================================
Public Sub UpdBits()

Dim tmpBinStr As String
Dim cntBitCounter As Byte

    If NewMainWindow.picBitSet8.Visible = True Then
        For cntBitCounter = 0 To 7
            tmpBinStr = tmpBinStr & NewMainWindow.bits8(cntBitCounter).Caption
        Next cntBitCounter
        NewMainWindow.txtParamValue.Text = Bin2Dec(tmpBinStr)
        Exit Sub
    End If
    
    
    If NewMainWindow.picBitSet16.Visible = True Then
        For cntBitCounter = 0 To 15
            tmpBinStr = tmpBinStr & NewMainWindow.bits16(cntBitCounter).Caption
        Next cntBitCounter
        NewMainWindow.txtParamValue.Text = Bin2Dec(tmpBinStr)
        Exit Sub
    End If

End Sub


'==========================================================================
' FUNCTION: GET BITS
' Converts decimal textbox value into bits label captions.
'==========================================================================
Public Sub GetBits()

Dim tmpBinStr As String
Dim cntBitCounter As Byte

    If NewMainWindow.picBitSet8.Visible = True Then
        tmpBinStr = Dec2Bin8(CByte(Val(NewMainWindow.txtParamValue.Text)))
        For cntBitCounter = 0 To 7
            NewMainWindow.bits8(cntBitCounter).Caption = Mid$(tmpBinStr, cntBitCounter + 1, 1)
        Next cntBitCounter
        Exit Sub
    End If
    
    If NewMainWindow.picBitSet16.Visible = True Then
        tmpBinStr = Dec2Bin16(CIntL(Val(NewMainWindow.txtParamValue.Text)))
        For cntBitCounter = 0 To 15
            NewMainWindow.bits16(cntBitCounter).Caption = Mid$(tmpBinStr, cntBitCounter + 1, 1)
        Next cntBitCounter
        Exit Sub
    End If
        

End Sub


'==========================================================================
' FUNCTION: LONG TO R,G,B STRING
' Converts long color value into RGB string.
'==========================================================================
Public Function LongToRGB(Source As Long) As String

    LongToRGB = CStr(Source And &HFF) & kDivider & CStr((Source And &HFF00&) \ &H100&) & kDivider & CStr((Source And &HFF0000) \ &H10000)

End Function
