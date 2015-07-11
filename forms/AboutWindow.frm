VERSION 5.00
Begin VB.Form AboutWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2140
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   7485
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7480
      Begin VB.Timer tmrCredits 
         Interval        =   10
         Left            =   0
         Top             =   1680
      End
      Begin VB.Image imgTREPLogo 
         Height          =   2145
         Left            =   -10
         Picture         =   "AboutWindow.frx":0000
         Top             =   -10
         Width           =   7500
      End
   End
   Begin VB.TextBox txtCredits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "AboutWindow.frx":687C
      Top             =   1440
      Width           =   7215
   End
End
Attribute VB_Name = "AboutWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Marker As Integer



Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Form_Load()

 If LenB(AboutText(0)) > 0 Then txtCredits.Text = AboutText(0)

 Me.Left = NewMainWindow.Left + ((NewMainWindow.Width - Me.Width) / 2)
 Me.Top = NewMainWindow.Top + ((NewMainWindow.Height - Me.Height) / 2)
 
 Marker = 0
  
End Sub

Private Sub tmrCredits_Timer()

Select Case txtCredits.Top

    Case Is <= 140, Is > 155
    
        If txtCredits.Top > 250 Then
        
            Marker = Marker + 1
    
            Select Case Marker
            
                Case 1: txtCredits.Text = AboutText(1): txtCredits.Top = 96
                Case 2: txtCredits.Text = AboutText(2): txtCredits.Top = 96
                Case 3: txtCredits.Text = AboutText(0): txtCredits.Top = 96: Marker = 0
                
            End Select
            
        Else
        
            txtCredits.Top = txtCredits.Top + 2
            
        End If
        
    Case Else:
        txtCredits.Top = txtCredits.Top + 0.04

End Select

End Sub
