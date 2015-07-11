VERSION 5.00
Begin VB.Form NewMainWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEP (FLExible Patcher)"
   ClientHeight    =   6630
   ClientLeft      =   840
   ClientTop       =   3825
   ClientWidth     =   18090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NewMainWindow.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1206
   Begin VB.PictureBox picBitTooltip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9060
      Picture         =   "NewMainWindow.frx":0CCA
      ScaleHeight     =   315
      ScaleWidth      =   1065
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label lblBitTooltip 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bit "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   30
         TabIndex        =   108
         Top             =   30
         Width           =   990
      End
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3990
      Picture         =   "NewMainWindow.frx":0E39
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   106
      TabStop         =   0   'False
      ToolTipText     =   "Search"
      Top             =   165
      Width           =   300
   End
   Begin VB.TextBox txtSearchFailed 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Text            =   "[sequence not found]"
      Top             =   195
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.TextBox txtSearchEnterText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Text            =   "[enter search pattern here]"
      Top             =   195
      Width           =   3780
   End
   Begin VB.Timer tmrTipTimer 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   17520
      Top             =   6120
   End
   Begin VB.TextBox txtSearchBox 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   " [ Enter search pattern here... ]"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Frame frmDevModeControls 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   1.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   45
      Top             =   5520
      Width           =   4215
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   4155
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   30
         Width           =   4160
         Begin VB.CommandButton btnCopy 
            Caption         =   "Copy"
            Height          =   375
            Left            =   1400
            TabIndex        =   33
            Top             =   60
            Width           =   615
         End
         Begin VB.CommandButton btnEditPatch 
            Caption         =   "Edit..."
            Height          =   375
            Left            =   30
            TabIndex        =   31
            Top             =   60
            Width           =   690
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "Kill"
            Height          =   375
            Left            =   2040
            TabIndex        =   34
            Top             =   60
            Width           =   615
         End
         Begin VB.CommandButton btnDeleteAll 
            Caption         =   "Clear"
            Height          =   375
            Left            =   2690
            TabIndex        =   35
            Top             =   60
            Width           =   615
         End
         Begin VB.CommandButton btnNewPatch 
            Caption         =   "New"
            Height          =   375
            Left            =   750
            TabIndex        =   32
            Top             =   60
            Width           =   615
         End
         Begin VB.CommandButton btnMovePatchUp 
            Caption         =   "ñ"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3330
            TabIndex        =   36
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton btnMovePatchDn 
            Caption         =   "ò"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3735
            TabIndex        =   37
            Top             =   60
            Width           =   375
         End
      End
   End
   Begin VB.Frame frmDevMode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   1.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   11280
      TabIndex        =   44
      Top             =   105
      Visible         =   0   'False
      Width           =   6735
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   1480
         Left            =   30
         ScaleHeight     =   1485
         ScaleWidth      =   6645
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   60
         Width           =   6650
         Begin VB.TextBox txtEditPatchDesc 
            Height          =   765
            Left            =   1280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   113
            Top             =   380
            Width           =   5295
         End
         Begin VB.TextBox txtEditPatchName 
            Height          =   285
            Left            =   1280
            MaxLength       =   255
            TabIndex        =   112
            Top             =   60
            Width           =   2895
         End
         Begin VB.TextBox txtEditPatchDependencies 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1280
            MaxLength       =   255
            TabIndex        =   111
            Top             =   1160
            Width           =   5295
         End
         Begin VB.TextBox txtTargetFile 
            Height          =   285
            Left            =   5120
            MaxLength       =   255
            TabIndex        =   110
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label lblPatchDescEdit 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   195
            Left            =   80
            TabIndex        =   117
            Top             =   440
            Width           =   855
         End
         Begin VB.Label lblPatchTitleEdit 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            Height          =   195
            Left            =   80
            TabIndex        =   116
            Top             =   105
            Width           =   360
         End
         Begin VB.Label lblPatchDependenciesEdit 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependencies:"
            Height          =   195
            Left            =   80
            TabIndex        =   115
            Top             =   1205
            Width           =   1065
         End
         Begin VB.Label lblTargetFile 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Target file:"
            Height          =   195
            Left            =   4280
            TabIndex        =   114
            Top             =   105
            Width           =   795
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   30
         ScaleHeight     =   345
         ScaleWidth      =   6675
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   5570
         Width           =   6670
         Begin VB.CommandButton btnDiscard 
            Caption         =   "Undo"
            Height          =   290
            Left            =   5710
            TabIndex        =   28
            Top             =   30
            Width           =   855
         End
         Begin VB.CommandButton btnLoadPatchSet 
            Caption         =   "Load Patches..."
            Height          =   290
            Left            =   90
            TabIndex        =   29
            Top             =   30
            Width           =   1575
         End
         Begin VB.CommandButton btnSavePatchSet 
            Caption         =   "Save Patches..."
            Height          =   290
            Left            =   1715
            TabIndex        =   30
            Top             =   30
            Width           =   1575
         End
      End
      Begin VB.Frame frmParameters 
         Height          =   1890
         Left            =   120
         TabIndex        =   51
         Top             =   3670
         Width           =   6495
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1600
            Left            =   40
            ScaleHeight     =   1605
            ScaleWidth      =   6390
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   200
            Width           =   6390
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   3720
               ScaleHeight     =   255
               ScaleWidth      =   2655
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   1360
               Width           =   2655
               Begin VB.OptionButton optParIgnore 
                  Caption         =   "Patch default"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   0
                  Left            =   1330
                  TabIndex        =   27
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1300
               End
               Begin VB.OptionButton optParIgnore 
                  Caption         =   "Ignore patch"
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   26
                  Top             =   0
                  Width           =   1260
               End
            End
            Begin VB.CommandButton btnKillParam 
               Caption         =   "Kill"
               Height          =   300
               Left            =   3050
               TabIndex        =   11
               Top             =   60
               Width           =   615
            End
            Begin VB.CommandButton btnAddParam 
               Caption         =   "Add"
               Height          =   300
               Left            =   2400
               TabIndex        =   10
               Top             =   60
               Width           =   615
            End
            Begin VB.TextBox txtParStringLength 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   5960
               MaxLength       =   2
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "1"
               Top             =   1010
               Width           =   375
            End
            Begin VB.OptionButton optParType 
               Caption         =   "bytes R,G,B"
               Enabled         =   0   'False
               Height          =   255
               Index           =   9
               Left            =   4590
               TabIndex        =   17
               Top             =   800
               Width           =   1215
            End
            Begin VB.OptionButton optParType 
               Caption         =   "string, length:"
               Enabled         =   0   'False
               Height          =   255
               Index           =   8
               Left            =   4590
               TabIndex        =   22
               Top             =   1040
               Width           =   1330
            End
            Begin VB.OptionButton optParType 
               Caption         =   "float"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   1485
               TabIndex        =   19
               Top             =   1040
               Width           =   825
            End
            Begin VB.OptionButton optParType 
               Caption         =   "bits (8)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   2415
               TabIndex        =   20
               Top             =   1040
               Width           =   855
            End
            Begin VB.OptionButton optParType 
               Caption         =   "bits (16)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   3495
               TabIndex        =   21
               Top             =   1040
               Width           =   960
            End
            Begin VB.OptionButton optParType 
               Caption         =   "byte (u)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   540
               TabIndex        =   13
               Top             =   800
               Width           =   900
            End
            Begin VB.OptionButton optParType 
               Caption         =   "integer (u)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   2415
               TabIndex        =   15
               Top             =   800
               Width           =   1060
            End
            Begin VB.OptionButton optParType 
               Caption         =   "byte (s)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   5
               Left            =   1485
               TabIndex        =   14
               Top             =   800
               Width           =   920
            End
            Begin VB.OptionButton optParType 
               Caption         =   "integer (s)"
               Enabled         =   0   'False
               Height          =   255
               Index           =   6
               Left            =   3495
               TabIndex        =   16
               Top             =   800
               Width           =   1075
            End
            Begin VB.OptionButton optParType 
               Caption         =   "long"
               Enabled         =   0   'False
               Height          =   255
               Index           =   7
               Left            =   540
               TabIndex        =   18
               Top             =   1040
               Width           =   600
            End
            Begin VB.TextBox txtParOffset 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4560
               TabIndex        =   24
               Top             =   60
               Width           =   1815
            End
            Begin VB.TextBox txtParDefault 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               MaxLength       =   255
               TabIndex        =   25
               Top             =   420
               Width           =   1815
            End
            Begin VB.TextBox txtParTitle 
               Enabled         =   0   'False
               Height          =   285
               Left            =   540
               MaxLength       =   255
               TabIndex        =   12
               Top             =   420
               Width           =   3135
            End
            Begin VB.ComboBox cmbParSlot 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "NewMainWindow.frx":3CE4
               Left            =   540
               List            =   "NewMainWindow.frx":3CE6
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   60
               Width           =   1815
            End
            Begin VB.Label lblParType 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Type:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   60
               TabIndex        =   72
               Top             =   855
               Width           =   615
            End
            Begin VB.Label lblParOffset 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Offsets:"
               Height          =   195
               Left            =   3870
               TabIndex        =   71
               Top             =   120
               Width           =   600
            End
            Begin VB.Label lblParValue 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Default:"
               Height          =   195
               Left            =   3885
               TabIndex        =   70
               Top             =   465
               Width           =   585
            End
            Begin VB.Label lblParTitle 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Title:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   60
               TabIndex        =   69
               Top             =   465
               Width           =   360
            End
            Begin VB.Label lblParSlot 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Slot:"
               Height          =   195
               Left            =   60
               TabIndex        =   68
               Top             =   120
               Width           =   330
            End
            Begin VB.Label lblParCondBehave 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Conditional behaviour:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2020
               TabIndex        =   67
               Top             =   1380
               Width           =   1620
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   3750
               X2              =   3750
               Y1              =   0
               Y2              =   680
            End
         End
         Begin VB.Label lblDynamicPatches 
            AutoSize        =   -1  'True
            Caption         =   " Dynamic patches "
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.Frame frmStaticPatches 
         Height          =   2085
         Left            =   120
         TabIndex        =   50
         Top             =   1550
         Width           =   6495
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1785
            Left            =   30
            ScaleHeight     =   1785
            ScaleWidth      =   6420
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   240
            Width           =   6420
            Begin VB.CommandButton btnKillOffset 
               Caption         =   "Kill"
               Height          =   315
               Left            =   690
               TabIndex        =   4
               Top             =   1440
               Width           =   520
            End
            Begin VB.CommandButton btnAddOffset 
               Caption         =   "Add"
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Top             =   1440
               Width           =   540
            End
            Begin VB.TextBox txtModifiedData 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   2040
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   735
               Width           =   4335
            End
            Begin VB.TextBox txtOriginalData 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   2040
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   0
               Width           =   4335
            End
            Begin VB.ListBox lbOffsetList 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1410
               IntegralHeight  =   0   'False
               ItemData        =   "NewMainWindow.frx":3CE8
               Left            =   120
               List            =   "NewMainWindow.frx":3CEA
               MousePointer    =   99  'Custom
               TabIndex        =   2
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optPatchIgnore 
               Caption         =   "Patch original"
               Height          =   255
               Index           =   0
               Left            =   5080
               TabIndex        =   8
               Top             =   1520
               Value           =   -1  'True
               Width           =   1300
            End
            Begin VB.OptionButton optPatchIgnore 
               Caption         =   "Ignore patch"
               Height          =   255
               Index           =   1
               Left            =   3720
               TabIndex        =   7
               Top             =   1520
               Width           =   1335
            End
            Begin VB.Label lblEditTrim 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   $"NewMainWindow.frx":3CEC
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   390
               Left            =   1440
               MousePointer    =   99  'Custom
               TabIndex        =   63
               Top             =   480
               Width           =   420
            End
            Begin VB.Label lblDataModified 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Modified:"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   1320
               TabIndex        =   62
               Top             =   1200
               Width           =   660
            End
            Begin VB.Label lblDataOriginal 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Original:"
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   1365
               TabIndex        =   61
               Top             =   0
               Width           =   600
            End
            Begin VB.Label lblDataCondBehave 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Conditional behaviour:"
               Height          =   195
               Left            =   2025
               TabIndex        =   60
               Top             =   1535
               Width           =   1620
            End
         End
         Begin VB.Label lblStaticPatches 
            AutoSize        =   -1  'True
            Caption         =   " Static patches "
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   0
            Width           =   1110
         End
      End
   End
   Begin VB.ListBox lbPatchList 
      Height          =   4935
      IntegralHeight  =   0   'False
      ItemData        =   "NewMainWindow.frx":3CFA
      Left            =   120
      List            =   "NewMainWindow.frx":3CFC
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   510
      Width           =   4215
   End
   Begin VB.Frame frmUserMode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   1.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   4440
      TabIndex        =   43
      Top             =   105
      Width           =   6735
      Begin VB.Frame frmParams 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   2.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   40
         TabIndex        =   58
         Top             =   5350
         Width           =   6630
         Begin VB.TextBox txtParamValue 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Left            =   4090
            TabIndex        =   78
            Text            =   "0"
            Top             =   150
            Width           =   1760
         End
         Begin VB.PictureBox picParam 
            BorderStyle     =   0  'None
            Height          =   515
            Left            =   0
            ScaleHeight     =   510
            ScaleWidth      =   6615
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   0
            Width           =   6615
            Begin VB.PictureBox picBitSet8 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   320
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   975
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   240
               Visible         =   0   'False
               Width           =   975
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   0
                  TabIndex        =   105
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   120
                  TabIndex        =   104
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   2
                  Left            =   240
                  TabIndex        =   103
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   3
                  Left            =   360
                  TabIndex        =   102
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   4
                  Left            =   480
                  TabIndex        =   101
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   5
                  Left            =   600
                  TabIndex        =   100
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   6
                  Left            =   720
                  TabIndex        =   99
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits8 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   7
                  Left            =   840
                  TabIndex        =   98
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.PictureBox picBitSet16 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   320
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   1935
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1940
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   15
                  Left            =   1800
                  TabIndex        =   96
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   14
                  Left            =   1680
                  TabIndex        =   95
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   13
                  Left            =   1560
                  TabIndex        =   94
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   12
                  Left            =   1440
                  TabIndex        =   93
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   11
                  Left            =   1320
                  TabIndex        =   92
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   10
                  Left            =   1200
                  TabIndex        =   91
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   9
                  Left            =   1080
                  TabIndex        =   90
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   8
                  Left            =   960
                  TabIndex        =   89
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   7
                  Left            =   840
                  TabIndex        =   88
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   6
                  Left            =   720
                  TabIndex        =   87
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   5
                  Left            =   600
                  TabIndex        =   86
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   4
                  Left            =   480
                  TabIndex        =   85
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   3
                  Left            =   360
                  TabIndex        =   84
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   2
                  Left            =   240
                  TabIndex        =   83
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   120
                  TabIndex        =   82
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label bits16 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   0
                  TabIndex        =   81
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.PictureBox picColorTip2 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               ForeColor       =   &H00000000&
               Height          =   250
               Left            =   5880
               ScaleHeight     =   225
               ScaleWidth      =   225
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   160
               Visible         =   0   'False
               Width           =   260
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   410
               Left            =   4065
               Locked          =   -1  'True
               TabIndex        =   76
               Top             =   90
               Width           =   2145
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000010&
               X1              =   6580
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Label lblParamDesc 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "This is parameter desc:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   270
               Left            =   1685
               MousePointer    =   99  'Custom
               TabIndex        =   77
               ToolTipText     =   "Click to show parameters list"
               Top             =   140
               Width           =   2310
            End
            Begin VB.Image imgUndo 
               Height          =   360
               Left            =   6230
               Picture         =   "NewMainWindow.frx":3CFE
               ToolTipText     =   "Undo"
               Top             =   90
               Width           =   360
            End
         End
      End
      Begin VB.Frame frmDependencies 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   2.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1240
         Left            =   50
         TabIndex        =   54
         Top             =   3960
         Width           =   6645
         Begin VB.TextBox txtDependenciesList 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H006060CF&
            Height          =   615
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   56
            Text            =   "NewMainWindow.frx":3EB2
            Top             =   380
            Width           =   6255
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000010&
            X1              =   6590
            X2              =   10
            Y1              =   30
            Y2              =   30
         End
         Begin VB.Label lblDepEnableNeeded 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enable needed patches"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   4560
            TabIndex        =   57
            Top             =   1000
            Width           =   1980
         End
         Begin VB.Label lblDependenciesTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "You need to select other patches:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   105
            Width           =   2880
         End
      End
      Begin VB.TextBox txtPatchDesc 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3330
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   645
         Width           =   6495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   6650
         X2              =   60
         Y1              =   550
         Y2              =   550
      End
      Begin VB.Label lblPatchTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Here Be Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   80
         TabIndex        =   49
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton btnDefaults 
      Caption         =   "Defaults"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   40
      Top             =   6165
      Width           =   975
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   42
      Top             =   6165
      Width           =   855
   End
   Begin VB.CommandButton btnModify 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   41
      Top             =   6165
      Width           =   855
   End
   Begin VB.CommandButton btnLoadPreset 
      Caption         =   "Load Preset..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   38
      Top             =   6165
      Width           =   1455
   End
   Begin VB.CommandButton btnSavePreset 
      Caption         =   "Save Preset..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   39
      Top             =   6165
      Width           =   1455
   End
   Begin VB.Image imgDevModeSwitcher 
      Height          =   180
      Left            =   150
      MousePointer    =   99  'Custom
      Picture         =   "NewMainWindow.frx":3ECE
      ToolTipText     =   "Go developer mode!"
      Top             =   6270
      Width           =   180
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About FLEP..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   240
      Left            =   390
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   6225
      Width           =   1140
   End
End
Attribute VB_Name = "NewMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cntBlink As Integer ' ui smiley blink


Private Sub btnAddOffset_Click()
 Call AddOffset
 Call LockDataBlock
 If txtOriginalData.Enabled = True Then txtOriginalData.SetFocus Else lbOffsetList.SetFocus
End Sub

Private Sub btnCopy_Click()
 Call ClickPatchList
 Call AddNewPatch(True)
 
 If frmDevMode.Visible = True Then
    txtEditPatchName.SetFocus
    txtEditPatchName.SelLength = Len(txtEditPatchName.Text)
 End If

End Sub

Private Sub btnDefaults_Click()

Call LockWindowUpdate(NewMainWindow.hwnd)
    Call ClickPatchList
    Call ResetList
    Call RefreshView
Call LockWindowUpdate(0)

End Sub

Private Sub btnDelete_Click()
 Call ClickPatchList
 Call DeletePatch(NewMainWindow.lbPatchList.ListIndex)
 lbPatchList.SetFocus
End Sub

Private Sub btnDeleteAll_Click()
 Call ClickPatchList
 Call ResetArrays
 Call PopulatePatchList
 Call RefreshView
 lbPatchList.SetFocus
End Sub

Private Sub btnDiscard_Click()
 
RestorePatch (lbPatchList.ListIndex)
lbPatchList.SetFocus

End Sub

Private Sub btnEditPatch_Click()

If DeveloperView = 1 Then

    Call ClickPatchList

    If DeveloperMode = 0 Then
       DeveloperMode = 1
    Else
       DeveloperMode = 0
    End If
    
End If

 Call RefreshView
 lbPatchList.SetFocus
End Sub

Private Sub btnKillOffset_Click()
 Call DestroyOffset(lbOffsetList.ListIndex)
 Call LockDataBlock
 lbOffsetList.SetFocus
End Sub

Private Sub btnLoadPatchSet_Click()

 Dim tmpSuccessFlag As Boolean
 Dim tmpFileName As String
 
    

    tmpFileName = OpenFileDialog(NewMainWindow.hwnd, CurrentDirectory, k_FilterTitle_PatchSet, k_FilterExt_PatchSet, k_OpenFileTitle)

    If tmpFileName = k_IsCancelPressed Then Exit Sub
    If CheckExistence(tmpFileName) = False Then Exit Sub
    If LoadPatches(tmpFileName, False) = False Then Exit Sub
    
    Call RefreshView

    lbPatchList.SetFocus
    
End Sub

Private Sub btnLoadPreset_Click()
 Dim tmpSuccessFlag As Boolean
 Dim tmpFileName As String
 
    tmpFileName = OpenFileDialog(NewMainWindow.hwnd, CurrentDirectory, k_FilterTitle_Preset, k_FilterExt_Preset, k_OpenFileTitle)

    If tmpFileName = k_IsCancelPressed Then Exit Sub
    
    If CheckExistence(tmpFileName) = False Then Exit Sub
    
    Call ApplyPatch(PrevPatchNumber)
    
    
    If LoadPreset(tmpFileName, False, True) = True Then
        RefreshView
        lbPatchList.SetFocus
    End If

End Sub

Private Sub btnModify_Click()
 Call ApplyPatch(PrevPatchNumber)
 Call CreateExecList
 Call ModifyAllExecutables
End Sub

Private Sub btnMovePatchDn_Click()
 Call LockWindowUpdate(NewMainWindow.hwnd)
    Call ClickPatchList
    Call MovePatchDown
    lbPatchList.SetFocus
 Call LockWindowUpdate(0)
End Sub

Private Sub btnMovePatchUp_Click()
 Call LockWindowUpdate(NewMainWindow.hwnd)
    Call ClickPatchList
    Call MovePatchUp
    lbPatchList.SetFocus
 Call LockWindowUpdate(0)
End Sub

Private Sub btnNewPatch_Click()
 Call ClickPatchList
 Call AddNewPatch(False)
 
 If frmDevMode.Visible = True Then
    txtEditPatchName.SetFocus
    txtEditPatchName.SelLength = Len(txtEditPatchName.Text)
    txtTargetFile.Text = DefaultExeName
 End If
End Sub

Private Sub btnQuit_Click()
 Call KillApp
End Sub

Private Sub btnSavePatchSet_Click()
If SavePatches = True Then Exit Sub
MsgBox "Error while saving patch set. Something went wrong.", vbExclamation
End Sub

Private Sub btnSavePreset_Click()
 Dim tmpFileName As String
 
    tmpFileName = SaveFileDialog(NewMainWindow.hwnd, CurrentDirectory, k_FilterTitle_Preset, k_FilterExt_Preset, k_SaveFileTitle)

    If tmpFileName = k_IsCancelPressed Then Exit Sub
    
    Call ApplyPatch(PrevPatchNumber)
    If SavePreset(tmpFileName, False, True) = True Then Exit Sub
    
    MsgBox "Errors while saving preset. Something went wrong."

End Sub

Private Sub btnSearch_Click()

End Sub


Private Sub btnAddParam_Click()
 If cmbParSlot.ListCount > 0 Then Call SaveParam(NewMainWindow.lbPatchList.ListIndex, cmbParSlot.ListIndex)
 
 Call AddParam
 txtParTitle.SetFocus
 
End Sub

Private Sub cmbParSlot_GotFocus()
 If cmbParSlot.ListIndex >= 0 Then Call SaveParam(NewMainWindow.lbPatchList.ListIndex, cmbParSlot.ListIndex)
End Sub

Private Sub cmbParSlot_Click()
    
    If cmbParSlot.ListIndex >= 0 Then Call LoadParam(cmbParSlot.ListIndex)
    Call LockParamBlock
    
End Sub

Private Sub btnKillParam_Click()

    Call DestroyParam(cmbParSlot.ListIndex)
    Call LockParamBlock
    cmbParSlot.SetFocus
End Sub


Private Sub Form_Load()

    Call InitializeWindow
    Call SwitchView
    Call LoadProgram
    Call LoadCfg
    Call LoadDefaultPatchSet
    Call LoadPresetsFromCfg
    Call ApplyParameters
    Call CheckCRC
    Call RefreshView
    
    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tmpSuccess As Boolean
  Call SetCurrentDirectory(CurrentDirectory)
  tmpSuccess = SavePreset((k_ConfigName), True, True)
End Sub


Private Sub imgUndo_Click()
txtParamValue.Text = PatchArray(PatchNameToNumber(lbPatchList.List(lbPatchList.ListIndex))).patchParams(CurrentParamNumber).parValue
PatchArray(PatchNameToNumber(lbPatchList.List(lbPatchList.ListIndex))).patchParams(CurrentParamNumber).parModdedValue = txtParamValue.Text
End Sub

Private Sub imgDevModeSwitcher_Click()
 
Call LockWindowUpdate(NewMainWindow.hwnd)

 Call ClickPatchList
 
 If DeveloperView = 1 Then
    DeveloperMode = 0
    DeveloperView = 0
    Call RefreshView
 Else
    DeveloperView = 1
 End If
 
 Call SwitchView
 
Call LockWindowUpdate(0)
 
End Sub
Private Sub imgDevModeSwitcher_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub imgDevModeSwitcher_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub



Private Sub bits8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
picBitTooltip.Visible = True
tmrTipTimer.Enabled = True
lblBitTooltip.Caption = "Bit " & CStr(Abs(Index - 7))
End Sub
Private Sub bits8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub bits8_Click(Index As Integer)
bits8(Index).Caption = (Not (bits8(Index).Caption)) And 1
Call UpdBits
End Sub


Private Sub bits16_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
picBitTooltip.Visible = True
tmrTipTimer.Enabled = True
lblBitTooltip.Caption = "Bit " & CStr(Abs(Index - 15))
End Sub
Private Sub bits16_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub bits16_Click(Index As Integer)
bits16(Index).Caption = (Not (bits16(Index).Caption)) And 1
Call UpdBits
End Sub



Private Sub lblAbout_Click()
FadeIn AboutWindow, 1, 230
End Sub
Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub



Private Sub lblDepEnableNeeded_Click()

    Call LockWindowUpdate(NewMainWindow.hwnd)
    Call FixDependencies(lbPatchList.ListIndex)
    Call LockWindowUpdate(0)

End Sub
Private Sub lblDepEnableNeeded_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub lblDepEnableNeeded_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub

Private Sub lblEditTrim_Click()
 
 If txtOriginalData.Enabled = False Or txtModifiedData.Enabled = False Then Exit Sub
 
 txtOriginalData.Text = StripOut(txtOriginalData.Text, " " & vbCrLf)
 txtModifiedData.Text = StripOut(txtModifiedData.Text, " " & vbCrLf)
 Call SaveOffset(NewMainWindow.lbPatchList.ListIndex, NewMainWindow.lbOffsetList.ListIndex)
End Sub
Private Sub lblEditTrim_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub lblEditTrim_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub

Private Sub lblParamDesc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) And UBound(PatchArray(lbPatchList.ListIndex).patchParams) > 0 Then SetCursor lHandCursorHandle
End Sub
Private Sub lblParamDesc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) And UBound(PatchArray(lbPatchList.ListIndex).patchParams) > 0 Then SetCursor lHandCursorHandle
End Sub

Private Sub lblParamDesc_Click()
If UBound(PatchArray(lbPatchList.ListIndex).patchParams) > 0 Then Call ShowParamListWindow(NewMainWindow.lbPatchList.ListIndex)
End Sub

Private Sub lblPatchTitle_Click()
 If LenB(WarnDepString) > 0 Then Call WarnDependencies
End Sub

Private Sub imgUndo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub imgUndo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub

Private Sub lbOffsetList_Click()
    
    If lbOffsetList.ListIndex >= 0 Then
        Call LoadOffset(lbOffsetList.ListIndex)
    End If
    
    Call LockDataBlock

End Sub

Private Sub lbOffsetList_DblClick()
    
    If lbOffsetList.ListIndex >= 0 Then
        Call EditOffset
    End If

End Sub


Private Sub lbPatchList_Click()
    Call ClickPatchList
End Sub

Private Sub optParIgnore_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If frmDevMode.Visible = True And cmbParSlot.Visible = True Then cmbParSlot.SetFocus
End Sub

Private Sub optParType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If frmDevMode.Visible = True And cmbParSlot.Visible = True Then cmbParSlot.SetFocus
End Sub

Private Sub optPatchIgnore_Click(Index As Integer)
 Call SaveOffset(NewMainWindow.lbPatchList.ListIndex, NewMainWindow.lbOffsetList.ListIndex)
End Sub



Private Sub picColorTip2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub picColorTip2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub picColorTip2_Click()

 Dim tmpColor As Long
 
    tmpColor = ShowColorDialog(NewMainWindow.hwnd, True, picColorTip2.BackColor)
    If tmpColor <> 1 Then NewMainWindow.txtParamValue.Text = LongToRGB(tmpColor)

End Sub


Private Sub picSearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub
Private Sub picSearch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (lHandCursorHandle > 0) Then SetCursor lHandCursorHandle
End Sub

Private Sub picSearch_Click()

 Dim tmpResult As Boolean
 
    If SearchForPatch(NewMainWindow.txtSearchEnterText.Text, False) = True Then
        Exit Sub
    Else
        tmpResult = SearchForPatch(NewMainWindow.txtSearchEnterText.Text, True)
    End If
    
    If tmpResult = False Then txtSearchFailed.Visible = True: tmrTipTimer.Enabled = True
    
    txtSearchEnterText.SetFocus

End Sub


Private Sub tmrTipTimer_Timer()
 txtSearchFailed.Visible = False
 picBitTooltip.Visible = False
 btnModify.Caption = k_ModButton
 tmrTipTimer.Enabled = False
End Sub

Private Sub txtEditPatchDependencies_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtEditPatchDependencies, KeyCode, Shift)
End Sub

Private Sub txtEditPatchDesc_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtEditPatchDesc, KeyCode, Shift)
End Sub

Private Sub txtEditPatchName_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtEditPatchName, KeyCode, Shift)
End Sub
Private Sub txtEditPatchName_LostFocus()
If txtEditPatchName.Text = vbNullString Then txtEditPatchName.Text = FindFreePatchName
End Sub

Private Sub txtParamValue_Change()
Call UpdateCurrentParam
Call UpdColor(txtParamValue.Text)
Call GetBits
End Sub



Private Sub txtParDefault_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtParDefault, KeyCode, Shift)
End Sub

Private Sub txtParOffset_KeyPress(KeyASCII As Integer)
KeyASCII = Asc(UCase(ChrW(KeyASCII)))
End Sub
Private Sub txtParOffset_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtParOffset, KeyCode, Shift)
End Sub

Private Sub txtParTitle_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtParTitle, KeyCode, Shift)
End Sub

Private Sub txtSearchEnterText_KeyPress(KeyASCII As Integer)
    If KeyASCII = 13 Then picSearch_Click
End Sub

Private Sub txtSearchEnterText_GotFocus()

    If txtSearchEnterText.ForeColor = &HC0C0C0 Then
        txtSearchEnterText.Text = vbNullString
        txtSearchEnterText.ForeColor = &H404040
    End If
    
End Sub


Private Sub txtModifiedData_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 4 Then NewMainWindow.txtModifiedData.Text = Clipboard.GetText()
 Call SaveOffset(NewMainWindow.lbPatchList.ListIndex, NewMainWindow.lbOffsetList.ListIndex)
End Sub
Private Sub txtModifiedData_KeyUp(KeyCode As Integer, Shift As Integer)
 Call SaveOffset(NewMainWindow.lbPatchList.ListIndex, NewMainWindow.lbOffsetList.ListIndex)
End Sub
Private Sub txtModifiedData_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtModifiedData, KeyCode, Shift)
End Sub


Private Sub txtOriginalData_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 4 Then NewMainWindow.txtOriginalData.Text = Clipboard.GetText()
 Call SaveOffset(NewMainWindow.lbPatchList.ListIndex, NewMainWindow.lbOffsetList.ListIndex)
End Sub
Private Sub txtOriginalData_KeyUp(KeyCode As Integer, Shift As Integer)
 Call SaveOffset(NewMainWindow.lbPatchList.ListIndex, NewMainWindow.lbOffsetList.ListIndex)
End Sub
Private Sub txtOriginalData_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtOriginalData, KeyCode, Shift)
End Sub


Private Sub txtParStringLength_GotFocus()
    optParType(8).Value = True
End Sub
Private Sub txtParStringLength_KeyPress(KeyASCII As Integer)
    KeyASCII = KeyFilterNumUnsigned(KeyASCII)
End Sub
Private Sub txtParStringLength_lostfocus()
    txtParStringLength.Text = CStr(FilterMinValue(txtParStringLength.Text, 1))
    txtParStringLength.Text = CStr(FilterMaxValue(txtParStringLength.Text, 64))
End Sub


Private Sub InitializeWindow()
    NewMainWindow.Width = 11370
    frmDevMode.Left = 296
    frmDevMode.Top = 7
    
    picBitSet8.Left = 4670
    picBitSet8.Top = 150
    
    picBitSet16.Left = 4170
    picBitSet16.Top = 150
    
    BorderWidth = (GetSystemMetrics(SM_CXDLGFRAME) * Screen.TwipsPerPixelX) * 2
    BorderHeight = (GetSystemMetrics(SM_CYDLGFRAME) * Screen.TwipsPerPixelY) * 2
    CaptionHeight = (GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY)
End Sub

Private Sub txtSearchFailed_GotFocus()
Call ClickPatchList
End Sub

Private Sub txtTargetFile_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectAll(txtTargetFile, KeyCode, Shift)
End Sub
