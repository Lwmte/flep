VERSION 5.00
Begin VB.Form ParamList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parameter list"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000004&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblParamLink 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "This is parameter caption!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   0
      Left            =   30
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Width           =   2115
   End
End
Attribute VB_Name = "ParamList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Object_MouseMove Me, x, y, Button, 10
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub lblParamLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub lblParamLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub

Private Sub lblParamLink_Click(Index As Integer)
CurrentParamNumber = Index
Unload Me
End Sub
