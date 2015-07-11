VERSION 5.00
Begin VB.Form InputKor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Input window"
   ClientHeight    =   495
   ClientLeft      =   -5925
   ClientTop       =   -405
   ClientWidth     =   3735
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lblInputDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   450
   End
End
Attribute VB_Name = "InputKor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnOK_Click()
 Text1.SetFocus
 KillMeNow
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then Text1.Text = vbNullString: KillMeNow
End Sub

Private Sub Form_Unload(Cancel As Integer)
 KillMeNow
End Sub

Private Sub Text1_keypress(KeyASCII As Integer)
 If KeyASCII = 13 Then KillMeNow
 KeyASCII = Asc(UCase(ChrW(KeyASCII)))
End Sub

Private Sub KillMeNow()

Select Case Text1.Text
    Case "": InputCallbackString = kNullStr
    Case Else: InputCallbackString = Text1.Text
End Select
    
Unload Me

End Sub

Private Sub InputKor_LostFocus()
 KillMeNow
End Sub
