Attribute VB_Name = "constFLEPConstants"
'==========================================================================
' FLEP 1.0 GLOBAL CONSTANTS MODULE
'
' Main code by Pyuaumch. Refactoring and rewrites by Lwmte.
'
' Contains all constants used to identify / load / save various parameters
' and names across whole app.
'=========================================================================
'

Option Explicit


'=========================================================================
'
'=========================================================================
'Public Const k As String = ""
'=========================================================================



'=========================================================================
' INPUT BOX CONSTANTS
'=========================================================================
Public Const k_IB_AddOffset As String = "Add offset: "
Public Const k_IB_EditOffset As String = "Edit offset: "
'=========================================================================


'=========================================================================
' GLOBAL STRING CONSTANTS
'=========================================================================
Public Const k_HeaderVersion As String = "1.1"  ' MUST be increased with every major update affecting config / preset / patch set structure.
Public Const k_ConfigName As String = "flep.cfg"

Public Const k_FilterTitle_Preset As String = "FLEP Preset (*.fps)"
Public Const k_FilterExt_Preset As String = "fps"
Public Const k_FilterTitle_PatchSet As String = "FLEP Patch Set (*.flp)"
Public Const k_FilterExt_PatchSet As String = "flp"

Public Const k_OpenFileTitle As String = "Open"
Public Const k_SaveFileTitle As String = "Save"

Public Const kNewPatchPrefix As String = "New patch "
Public Const k_ParamPrefix As String = "Parameter "
Public Const k_EmptyPatchName As String = "Empty patch entry"
Public Const k_CopyPostfix As String = " - copy"

Public Const k_ModButton As String = "Modify"
Public Const k_ModButtonSuccess As String = "Done!"

Public Const kDivider As String = ","
Public Const kDivider2 As String = "|"
Public Const kTerminator As String = "//"
Public Const kEquals As String = "="
Public Const kCommentary As String = ";"
Public Const kMaskHex As String = "0123456789ABCDEF"
Public Const kMaskFloat As String = "0123456789.,"

Public Const kNullStr = "EMPTY"
Public Const k_IsCancelPressed As String = "CANCEL"

Public Const kVerTitle = "v. "
'=========================================================================


'=========================================================================
' COMMON CONFIG / PRESET NAME STRING CONSTANTS
'=========================================================================
Public Const kPresetVersion As String = "presetversion"    ' preset header

Public Const kSilentPatch As String = "silentpatch"
Public Const kDeveloperMode As String = "developermode"
Public Const kDeveloperView As String = "developerview"
Public Const kDefaultPatchSet As String = "defaultpatchset"
Public Const kMaxParameters As String = "maxparameters"

Public Const kDefaultExeName As String = "defaultexename"
Public Const kLastPreset As String = "lastpreset"
Public Const kLastFile As String = "lastfile"

Public Const kWindowPosX As String = "winposx"
Public Const kWindowPosY As String = "winposy"

Public Const kWindowTitle As String = "windowtitle"
Public Const kAboutText1 As String = "abouttext1"
Public Const kAboutText2 As String = "abouttext2"
Public Const kAboutText3 As String = "abouttext3"

Public Const kPatchPrefix = "patchpreset"
'=========================================================================


'=========================================================================
' COMMON PATCH SET NAME STRING CONSTANTS
'=========================================================================
Public Const kPatchSetVer As String = "patchsetversion"
Public Const kPatchHeader As String = "header"
Public Const kPatchFooter As String = "footer"
'=========================================================================


'=========================================================================
' COMMON PATCH NAME STRING CONSTANTS
'=========================================================================
Public Const kPatchEnabled As String = "enabled"
Public Const kPatchName As String = "name"
Public Const kPatchDesc As String = "description"
Public Const kPatchCategory As String = "category"
Public Const kPatchDependencies As String = "dependencies"
Public Const kPatchFile As String = "filename"
'=========================================================================


'=========================================================================
' PATCH DATAS NAME STRING CONSTANTS
'=========================================================================
Public Const kDataOffset As String = "dataoffset"
Public Const kDetaDefault As String = "datadefault"
Public Const kDataModified As String = "datamodified"
Public Const kDataCondBehave As String = "datacondbehave"
'=========================================================================


'=========================================================================
' PARAMETERS NAME STRING CONSTANTS
'=========================================================================
Public Const kParEnabled As String = "parenabled"
Public Const kParName As String = "parname"
Public Const kParOffset As String = "paroffset"
Public Const kParValue As String = "parvalue"
Public Const kParModdedValue As String = "parmoddedvalue"
Public Const kParType As String = "partype"
Public Const kParCondBehave As String = "parcondbehave"
'=========================================================================


'=========================================================================
' BINARY INTERPRETER COMMANDS
'=========================================================================
Public Const bi_kLen As String = "SETFILELENGTH"
Public Const bi_kFill As String = "FILL"

Public Const bi_MaxCommandLength As Integer = 30 'Multiplied by 2 for LenB function.
'=========================================================================



