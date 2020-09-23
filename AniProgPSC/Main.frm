VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " AniProg  by RR"
   ClientHeight    =   8040
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   9825
   DrawStyle       =   2  'Dot
   DrawWidth       =   2
   FillStyle       =   6  'Cross
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRoughPAL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "<"
      Height          =   315
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   " Extract Rough Palette (24 BPP only) "
      Top             =   3645
      Width           =   180
   End
   Begin VB.Timer TimerSR 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   900
      Top             =   7575
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H80000013&
      Caption         =   "Play"
      Height          =   270
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6390
      Width           =   600
   End
   Begin VB.PictureBox picPlay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   465
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   65
      Top             =   6780
      Width           =   510
   End
   Begin VB.PictureBox PICT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   60
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   64
      Top             =   7515
      Width           =   480
   End
   Begin VB.CommandButton cmdEffects 
      BackColor       =   &H80000013&
      Caption         =   "&Effects"
      Height          =   405
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3540
      Width           =   780
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Index           =   2
      Left            =   4935
      MultiLine       =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Text            =   "Main.frx":0442
      Top             =   1635
      Width           =   4125
   End
   Begin VB.CommandButton cmdTestCursor 
      BackColor       =   &H80000013&
      Caption         =   "Test &cursor"
      Height          =   630
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5490
      Width           =   765
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   1
      Left            =   3420
      MultiLine       =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Text            =   "Main.frx":044F
      Top             =   1185
      Width           =   3015
   End
   Begin VB.CommandButton cmdTestAni 
      BackColor       =   &H80000013&
      Caption         =   "Test &ani"
      Height          =   630
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4635
      Width           =   780
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   0
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Main.frx":045C
      Top             =   1008
      Width           =   3975
   End
   Begin VB.Frame fraTools 
      BackColor       =   &H80000013&
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4410
      Left            =   1335
      TabIndex        =   19
      Top             =   3390
      Width           =   1800
      Begin VB.CommandButton cmdUndoRedo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "UA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   1080
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   " Undo All "
         Top             =   3990
         Width           =   435
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   20
         Left            =   1155
         Picture         =   "Main.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   " Mirror L/R "
         Top             =   3030
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   19
         Left            =   630
         Picture         =   "Main.frx":057B
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   " Roll / "
         Top             =   3030
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   18
         Left            =   135
         Picture         =   "Main.frx":0785
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   " Roll \ "
         Top             =   3030
         Width           =   390
      End
      Begin VB.CommandButton cmdUndoRedo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   " Clear backups "
         Top             =   3990
         Width           =   405
      End
      Begin VB.CommandButton cmdUndoRedo 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1080
         MaskColor       =   &H00E0E0E0&
         Picture         =   "Main.frx":098F
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   " Redo "
         Top             =   3555
         Width           =   420
      End
      Begin VB.CommandButton cmdUndoRedo 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         Picture         =   "Main.frx":0B99
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   " Undo "
         Top             =   3555
         Width           =   405
      End
      Begin VB.CommandButton cmdUndoRedo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   " Deselect "
         Top             =   3990
         Width           =   330
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SR"
         Height          =   390
         Index           =   9
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   " Select rectangle "
         Top             =   1680
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   17
         Left            =   1155
         Picture         =   "Main.frx":0DA3
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   " Flip UD "
         Top             =   2580
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   16
         Left            =   645
         Picture         =   "Main.frx":0EB5
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   " Flip LR "
         Top             =   2580
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   15
         Left            =   135
         Picture         =   "Main.frx":0FC7
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   " Rotate 90 "
         Top             =   2580
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   14
         Left            =   1155
         Picture         =   "Main.frx":10D9
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   " Shift U/D "
         Top             =   2145
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   13
         Left            =   645
         Picture         =   "Main.frx":12E3
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   " Shift L/R "
         Top             =   2130
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   12
         Left            =   1155
         Picture         =   "Main.frx":13F5
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   " Roll U/D "
         Top             =   1680
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   11
         Left            =   645
         Picture         =   "Main.frx":15FF
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   " Roll L/R "
         Top             =   1680
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   10
         Left            =   135
         Picture         =   "Main.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   " Replace L by R color "
         Top             =   2130
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   8
         Left            =   1155
         Picture         =   "Main.frx":1A13
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   " PickA Color "
         Top             =   1155
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   7
         Left            =   135
         Picture         =   "Main.frx":1C1D
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   " Erase to TC "
         Top             =   1155
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   6
         Left            =   660
         Picture         =   "Main.frx":1E27
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   " Fill "
         Top             =   1155
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   5
         Left            =   1155
         Picture         =   "Main.frx":2031
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   " Filled Oval "
         Top             =   705
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   4
         Left            =   645
         Picture         =   "Main.frx":223B
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   " Filled Rectangle "
         Top             =   705
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   3
         Left            =   135
         Picture         =   "Main.frx":2445
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   " Oval "
         Top             =   705
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   2
         Left            =   1155
         Picture         =   "Main.frx":264F
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " Rectangle "
         Top             =   255
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   1
         Left            =   645
         Picture         =   "Main.frx":2859
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   " Line "
         Top             =   255
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   0
         Left            =   135
         Picture         =   "Main.frx":2A63
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   " Dot "
         Top             =   255
         Width           =   390
      End
      Begin VB.Label LabQ1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   1575
         TabIndex        =   59
         Top             =   3555
         Width           =   165
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   0
         X2              =   1770
         Y1              =   3465
         Y2              =   3465
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   0
         X2              =   1755
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label LabQ1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   1572
         TabIndex        =   56
         Top             =   1692
         Width           =   168
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000013&
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   52
         Top             =   3795
         Width           =   420
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000013&
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   51
         Top             =   3570
         Width           =   390
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   1755
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   15
         X2              =   1785
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.Frame fraColors 
      BackColor       =   &H80000013&
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4410
      Left            =   3285
      TabIndex        =   5
      Top             =   3390
      Width           =   1845
      Begin VB.PictureBox picGD 
         AutoRedraw      =   -1  'True
         Height          =   285
         Left            =   105
         ScaleHeight     =   225
         ScaleWidth      =   1590
         TabIndex        =   54
         Top             =   3945
         Width           =   1650
      End
      Begin VB.PictureBox PICM 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1155
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   45
         Top             =   3150
         Width           =   510
      End
      Begin VB.PictureBox PICI 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   330
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   44
         Top             =   3150
         Width           =   510
      End
      Begin VB.PictureBox picPAL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   240
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   6
         Top             =   240
         Width           =   1470
      End
      Begin VB.Shape shpGD 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   90
         Left            =   300
         Top             =   4260
         Width           =   60
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000013&
         Caption         =   "GD Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   3720
         Width           =   780
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   "Im"
         Height          =   210
         Index           =   3
         Left            =   75
         TabIndex        =   47
         Top             =   3330
         Width           =   225
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   "Ma"
         Height          =   210
         Index           =   4
         Left            =   870
         TabIndex        =   46
         Top             =   3345
         Width           =   285
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1320
         TabIndex        =   36
         Top             =   2895
         Width           =   330
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   2
         Left            =   840
         TabIndex        =   35
         Top             =   2895
         Width           =   330
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   1
         Left            =   480
         TabIndex        =   34
         Top             =   2895
         Width           =   330
      End
      Begin VB.Label LabQ1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   18
         Top             =   2610
         Width           =   165
      End
      Begin VB.Label Lab224 
         BackColor       =   &H80000013&
         Caption         =   "231, 227, 231"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   915
         TabIndex        =   16
         Top             =   2295
         Width           =   825
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   2895
         Width           =   330
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   "TC"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   2250
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   "R"
         Height          =   210
         Index           =   1
         Left            =   1470
         TabIndex        =   13
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         Caption         =   "L"
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   12
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label LabTCul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Top             =   2205
         Width           =   420
      End
      Begin VB.Label LabCulLR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   9
         Top             =   2625
         Width           =   1035
      End
      Begin VB.Label LabCulLR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   990
         TabIndex        =   8
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label LabCulLR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   7
         Top             =   1800
         Width           =   420
      End
   End
   Begin VB.PictureBox aPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   390
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   420
      Width           =   480
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   5505
      MousePointer    =   99  'Custom
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
      Begin VB.Shape STool 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         DrawMode        =   7  'Invert
         Height          =   435
         Index           =   0
         Left            =   300
         Top             =   570
         Width           =   450
      End
      Begin VB.Shape STool 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         DrawMode        =   7  'Invert
         Height          =   435
         Index           =   1
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   450
      End
      Begin VB.Line LineTool 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         DrawMode        =   7  'Invert
         X1              =   16
         X2              =   46
         Y1              =   16
         Y2              =   33
      End
      Begin VB.Shape shpSelRect 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         DrawMode        =   7  'Invert
         Height          =   660
         Left            =   105
         Top             =   1725
         Width           =   1110
      End
      Begin VB.Shape shpSelRect2 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   16  'Merge Pen
         Height          =   660
         Left            =   720
         Top             =   1560
         Width           =   1110
      End
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   645
      Shape           =   3  'Circle
      Top             =   7500
      Width           =   165
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   570
      Shape           =   3  'Circle
      Top             =   7425
      Width           =   315
   End
   Begin VB.Label LabCenFrm 
      BackColor       =   &H80000013&
      Caption         =   "Label1"
      Height          =   180
      Left            =   660
      TabIndex        =   63
      ToolTipText     =   " Center form "
      Top             =   7500
      Width           =   165
   End
   Begin VB.Label LabBPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "LabBPP"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   62
      Top             =   4170
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   4050
      Index           =   1
      Left            =   5400
      Top             =   3405
      Width           =   4050
   End
   Begin VB.Label LabFileName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   43
      Top             =   75
      Width           =   255
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Frames.  LC select,  RC Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   30
      Width           =   2385
   End
   Begin VB.Label LabTool 
      BackColor       =   &H80000013&
      Caption         =   "LabTool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7125
      TabIndex        =   41
      Top             =   7560
      Width           =   1725
   End
   Begin VB.Shape Shape3 
      Height          =   4050
      Index           =   0
      Left            =   5430
      Top             =   3420
      Width           =   4050
   End
   Begin VB.Label LabXY 
      BackColor       =   &H80000013&
      Caption         =   "31, 31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6060
      TabIndex        =   10
      Top             =   7530
      Width           =   870
   End
   Begin VB.Label LabPicNum 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "64"
      Height          =   255
      Left            =   5565
      TabIndex        =   3
      Top             =   7500
      Width           =   285
   End
   Begin VB.Label LabFN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   375
      TabIndex        =   2
      Top             =   930
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000013&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00B2B2B2&
      FillStyle       =   0  'Solid
      Height          =   2970
      Left            =   300
      Top             =   315
      Width           =   9225
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      FillColor       =   &H00C0FFFF&
      Height          =   3060
      Left            =   240
      Top             =   270
      Width           =   9330
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuFileSet 
         Caption         =   "&New ani cursor"
         Index           =   0
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "&Open ani cursor"
         Index           =   1
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "&Save ani cursor"
         Index           =   2
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "&Load JASC PAL"
         Index           =   4
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "Save &JASC PAL"
         Index           =   6
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "Open cur cursor"
         Index           =   8
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "Save cur cursor"
         Index           =   9
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "Open ico"
         Index           =   11
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFileSet 
         Caption         =   "E&xit"
         Index           =   13
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuCP 
         Caption         =   "&Copy"
         Index           =   0
      End
      Begin VB.Menu mnuCP 
         Caption         =   "&Paste"
         Index           =   1
      End
      Begin VB.Menu mnuCP 
         Caption         =   "C&lear to TC"
         Index           =   2
      End
      Begin VB.Menu mnuCP 
         Caption         =   "C&apture"
         Index           =   3
      End
      Begin VB.Menu mnuCP 
         Caption         =   "&Remove frame"
         Index           =   4
      End
      Begin VB.Menu mnuCP 
         Caption         =   "&Swap"
         Index           =   5
         Begin VB.Menu mnuCPSub 
            Caption         =   "Swap with Next"
            Index           =   0
         End
         Begin VB.Menu mnuCPSub 
            Caption         =   "Swap with Previous"
            Index           =   1
         End
         Begin VB.Menu mnuCPSub 
            Caption         =   "Swap with First"
            Index           =   2
         End
         Begin VB.Menu mnuCPSub 
            Caption         =   "Swap with Last"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuDetails 
      Caption         =   "&Details"
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "&Effects"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Main.frm

' AniProg by Robert Rayment 3 Oct 2004
' Update to calculate total duration in frmDetails

' Simplifications:-

' 1,4,8 or 24 BPP icons
' 32 x 32 size only
' Sequencing ignored
' All palettes the same as the first one in ani cursor
' Only up to 64 frames catered for
' All frames cycled thru (ie cFrames=cSteps, see Ani.txt)


Option Explicit
Option Base 1

Private ORGFrmWid As Long
Private ORGFrmHit As Long

' Public aniFSpec$         ' Temp filespec for Testing anicursor
Private curFSpec$         ' Temp filespec for Testing cursor

' Rectangle & Oval (STool(0) & STool(1)) fixed start coords
Private XS0 As Single
Private YS0 As Single

Private zSSF As Single  ' Scale factor for PIC Shape border width

'Public zSF As Single   ' Scale factor for Text help locs
                        ' applied to develop IDE locs for
                        ' Large fonts
                        
Private SRCount As Long ' SeleRect SRTimer count

Private CommonDialog1 As OSDialog

Private Sub cmdRoughPAL_Click()
If BPP = 24 Then
   ExtractRoughPalette
   ShowPalette
End If
End Sub


Private Sub Form_Load()
Dim dev_mode As DEVMODE

   aPlay = False
   TimerSR.Enabled = False
   
    ' Get the current mode.
    If EnumDisplaySettings(ByVal vbNullString, _
       ENUM_CURRENT_SETTINGS, dev_mode) = 0 Then
       ' Assume 16bit color
       ScreenBits = 16
    Else
      ScreenBits = dev_mode.dmBitsPerPel
    End If
    If ScreenBits < 16 Then
      aPlay = False
      MsgBox "Color setting < 16 bits", vbCritical, "Screen setting"
      Unload frmCapture
      Unload frmHelp
      Unload frmEffects
      Unload Me
      End
   End If

   PICT.Visible = False
   Tool = 99 '0
   
   LabTool = ""
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   Form1.Width = 9915
   Form1.Height = 9000 '9120
   ORGFrmWid = Form1.Width
   ORGFrmHit = Form1.Height
   
   AniTest = False
   LCul = 0
   RCul = vbWhite
   picNum = 1
   
   LabCulLR(0).BackColor = LCul
   LabCulLR(1).BackColor = RCul
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   kspace = 8   ' Edit grid spacing
   ReDim RateValues(64)
   
   ' Position textHelps
   zSF = aPic(1).Width / 32   ' eg @ 120 dpi = 40/32
   txtHelp(0).Left = 237 * zSF
   txtHelp(0).Top = 323 * zSF
   txtHelp(0).Visible = False
   txtHelp(1).Left = 206 * zSF
   txtHelp(1).Top = 313 * zSF
   txtHelp(1).Visible = False
   txtHelp(2).Left = 206 * zSF
   txtHelp(2).Top = 413 * zSF
   txtHelp(2).Visible = False
   txtHelpTEXT
     
   'TransparentColor = RGB(224, 224, 224)
   TransparentColor = RGB(231, 227, 231)
   
   frmHelpLeft = 6000 '(Form1.Left + Form1.Width - frmHelp.Width) / STX
   frmHelpTop = 60 * STX
   
   frmEffectsLeft = 1875 '(Form1.Left + Form1.Width - frmHelp.Width) / STX
   frmEffectsTop = 5385
   
   Reduction = 0
   Rotation = 0
   
   NumFrames = 2
   SizeAll
   
   ' Edit picbox PIC  kspace=8
   zSSF = 271 / 257   ' Shape3(0).Width/PIC.Width
   With PIC
      .Left = Shape3(0).Left + 3 '369
      .Top = Shape3(0).Top + 7   '233
      .Width = kspace * 32 + 1
      .Height = kspace * 32 + 1
      PIC.BackColor = TransparentColor
   End With
   Shape3(0).Width = PIC.Width * zSSF
   Shape3(0).Height = PIC.Height * zSSF
   Shape3(1).Width = PIC.Width * zSSF
   Shape3(1).Height = PIC.Height * zSSF
   
   PICM.BackColor = vbWhite
   
   GridCul = RGB(146, 146, 146) ' & X=760 on picGD
   DrawGridOnPIC GridCul, TransparentColor
   picNumHighlighted = 1         ' For LRates() & LabFN()
   
   InitNEW
   
   Me.Show
   
   With frmDetails
      .txtHelp.Left = 288 * zSF
      .txtHelp.Top = 58 * zSF
      .txtHelp.Visible = False
      .Show vbModeless
      .WindowState = vbMinimized
   End With
   
   LineTool.Visible = False
   STool(0).Visible = False  ' Rectangle
   STool(1).Visible = False  ' Oval
   
   Unload frmCapture
   Unload frmHelp
   Unload frmEffects
   
   DoEvents
End Sub


Private Sub InitNEW()
   TimerSR.Enabled = False
   aSelect = False
   Set_SelRect
   
   optTools(9).ForeColor = 0
   LabFileName = " New "
   aNew = True
   NumFrames = 2
   picNum = 1
   aCopy = False     ' Copy/Paste
   mnuCP(1).Enabled = False
   aMouseDown = False            ' No PIC action yet
   LabFN(picNumHighlighted).BackColor = vbWhite
   picNumHighlighted = 1         ' For LRates() & LabFN()
   LabFN(1).BackColor = vbYellow
   GridCul = RGB(146, 146, 146) ' & X=760 on picGD
   shpGD.Left = 675
   picGD_MouseUp 1, 0, 675, 90
   
   AniTitle$ = "Untitled"
   AniAuthor$ = "NoName"
   BPP = 8     ' Default to 8BPP
   HotX = 0
   HotY = 0
   
   ReDim XORA(32, 32, NumFrames)
   ReDim ANDA(32, 32, NumFrames)
   picNum = 2
   ClearXORA_ANDA
   picNum = 1
   ClearXORA_ANDA
   RestartBackUps
   cmdUndoRedo(4).Enabled = False
   
   ReDim RateValues(64)
   NColors = 256
   DefaultPalette
   RateValues(1) = 3 ' Defaults
   RateValues(2) = 3
   Fill_fraINFO
   LabTCul.BackColor = TransparentColor
   ShowAllIcons    ' Does aPic_MouseUp Cint(picNum), 0, 0, 0, 0
   
   picNumHighlighted = 1
   With frmDetails
      .fraINFO.Caption = "NEW"
      .LRates(picNumHighlighted).BackColor = vbWhite
      .LRates(1).BackColor = vbYellow
   End With
End Sub

Public Sub Fill_fraINFO()
On Error Resume Next
   With frmDetails
      .txtTitle.Text = AniTitle$
      .txtAuthor.Text = AniAuthor$
      avscr = False
      .vscrNumFrames.Value = NumFrames
      avscr = True
      .LabNumFrames = Str$(NumFrames)
      Select Case BPP
      Case 1: .optBPP(0).Value = True
      Case 4: .optBPP(1).Value = True
      Case 8: .optBPP(2).Value = True
      Case 24: .optBPP(3).Value = True
      Case Else: Exit Sub
      End Select
      If HotX < 0 Or HotX > 31 Then HotX = 0
      If HotY < 0 Or HotY > 31 Then HotY = 0
      .LabHotX = Str$(HotX)
      .LabHotY = Str$(HotY)
      .hscrHotX.Value = HotX
      .vscrHotY.Value = HotY
      .ShowHotXY vbRed
      Visibility
      CalcDuration
      frmDetails.LRates_MouseUp CInt(picNum), 0, 0, 0, 0
      If NumFrames > 1 Then
         mnuCP(4).Enabled = True   ' Remove
      Else
         mnuCP(4).Enabled = False   ' Remove
      End If
   End With
   CheckSwap
On Error GoTo 0
End Sub

Private Sub LabCenFrm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If WindowState <> vbMaximized Then
      Shape4.BackColor = RGB(&HFF, &HC0, &HC0)
      Shape5.BackColor = vbRed
      Me.Left = (Screen.Width - Me.Width) / 2
      Me.Top = (Screen.Height - Me.Height) / 2
      Sleep 60
   End If
End Sub

Private Sub LabCenFrm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Shape4.BackColor = RGB(&HFF, &HE0, &HC0)
      Shape5.BackColor = RGB(&HFF, &HC0, &HC0)
End Sub

Private Sub LabQ1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtHelp(Index).Visible = True
End Sub

Private Sub LabQ1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtHelp(Index).Visible = False
End Sub

Private Sub mnuDetails_Click()
   aPlay = False
   If AniTest Then
      RestoreOldCursor
      AniTest = False
      On Error Resume Next
      Kill aniFSpec$
      DoEvents
   End If

 If frmDetails.WindowState = vbNormal Then
   frmDetails.WindowState = vbMinimized
 Else
   frmDetails.WindowState = vbNormal
 End If
End Sub

Private Sub cmdEffects_Click()
   mnuEffects_Click
End Sub

Private Sub mnuEffects_Click()
   If aMouseDown Then Exit Sub
   aEffects = True
   frmEffects.Show vbModal 'vbModeless ', Me
   ShowPalette
   ShowAllIcons
End Sub

Private Sub mnuHelp_Click()
Dim A$
   If aMouseDown Then Exit Sub
   A$ = PathSpec$ & "APHelp.txt"
   If Len(Dir$(A$)) = 0 Then
      MsgBox "APHelp.txt missing ", , "AniProg - Help"
      Exit Sub
   Else
      frmHelp.Show vbModeless ', Me
   End If
End Sub


'#############################################################

Private Sub cmdUndoRedo_Click(Index As Integer)
   If NumBackUps > 0 Then
      Select Case Index
      Case 0   ' Undo
        If BackUpNumber > 1 Then
            BackUpNumber = BackUpNumber - 1
            CopyMemory XORA(1, 1, picNum), XORABU(1, 1, BackUpNumber), 32 * 32 * 4
            CopyMemory ANDA(1, 1, picNum), ANDABU(1, 1, BackUpNumber), 32 * 32
            ShowAllIcons
            cmdUndoRedo(1).Enabled = True      ' Redo
            If BackUpNumber = 1 Then
               cmdUndoRedo(0).Enabled = False  ' Undo
            End If
         End If
      Case 1  ' Redo
         If BackUpNumber < NumBackUps Then
            BackUpNumber = BackUpNumber + 1
            CopyMemory XORA(1, 1, picNum), XORABU(1, 1, BackUpNumber), 32 * 32 * 4
            CopyMemory ANDA(1, 1, picNum), ANDABU(1, 1, BackUpNumber), 32 * 32
            ShowAllIcons
            cmdUndoRedo(0).Enabled = True      ' Undo
            If BackUpNumber = NumBackUps Then
               cmdUndoRedo(1).Enabled = False  ' Redo
            End If
         End If
      Case 2 ' Clear backups
         RestartBackUps
         cmdUndoRedo(4).Enabled = False
      Case 4 ' Undo All
         BackUpNumber = 1
         NumBackUps = 1
         CopyMemory XORA(1, 1, picNum), XORABU(1, 1, 1), 32 * 32 * 4
         CopyMemory ANDA(1, 1, picNum), ANDABU(1, 1, 1), 32 * 32
         cmdUndoRedo(0).Enabled = False    'Undo
         cmdUndoRedo(1).Enabled = False    'Undo
         Label12(0) = "T" & Str$(NumBackUps)
         Label12(1) = "N" & Str$(BackUpNumber)
         ShowAllIcons
         ReDim Preserve XORABU(32, 32, 1)
         ReDim Preserve ANDABU(32, 32, 1)
         cmdUndoRedo(4).Enabled = False
      End Select
   End If
   
   If Index = 3 Then ' Deselect
     SelRect_OFF
     optTools(9).ForeColor = 0
     If Tool <= 9 Then
        optTools(Tool).Value = False
     End If
     Tool = 99
     LabTool = "Deselected"
   End If
   Label12(0) = "T" & Str$(NumBackUps)
   Label12(1) = "N" & Str$(BackUpNumber)
End Sub

 Public Sub RestartBackUps()
   BackUpNumber = 1
   NumBackUps = 1
   ReDim XORABU(32, 32, 1)
   ReDim ANDABU(32, 32, 1)
   CopyBU   ' Initial state of aPic()
   
'   CopyMemory XORABU(1, 1, NumBackUps), XORA(1, 1, picNum), 32 * 32 * 4
'   CopyMemory ANDABU(1, 1, NumBackUps), ANDA(1, 1, picNum), 32 * 32
   
   cmdUndoRedo(0).Enabled = False    'Undo
   cmdUndoRedo(1).Enabled = False    'Undo
   Label12(0) = "T" & Str$(NumBackUps)
   Label12(1) = "N" & Str$(BackUpNumber)
 End Sub

Private Sub BackUp()
   If NumBackUps > 0 Then
      NumBackUps = NumBackUps + 1
      BackUpNumber = NumBackUps
      ReDim Preserve XORABU(32, 32, NumBackUps)
      ReDim Preserve ANDABU(32, 32, NumBackUps)
      CopyBU
      cmdUndoRedo(0).Enabled = True     'Undo
      cmdUndoRedo(1).Enabled = False    'Redo
      Label12(0) = "T" & Str$(NumBackUps)
      Label12(1) = "N" & Str$(BackUpNumber)
   End If
End Sub

'#############################################################

Public Sub CalcDuration()
Dim k As Long
   zDuration = 0
      For k = 1 To NumFrames
         zDuration = zDuration + RateValues(k)
      Next k
      If zDuration = 0 Then
         zDuration = (NumFrames * anih.JifRate) / 60
      Else
         zDuration = zDuration / 60
      End If
   With frmDetails
      For k = 1 To NumFrames
         .LRates(k) = RateValues(k)
      Next k
      .LabTotTime = Format(Str$(zDuration), "##.00")
   End With
End Sub

Public Sub DefaultPalette()
Dim N As Long
   Select Case BPP
   Case 1
   ' Black & White
   ReDim BGRA(0 To 1)
      BGRA(0).B = 0
      BGRA(0).G = 0
      BGRA(0).R = 0
      BGRA(1).B = 255
      BGRA(1).G = 255
      BGRA(1).R = 255
      NColors = 2
   Case 4
   ' QBColors
   ReDim BGRA(0 To 15)
   For N = 0 To 15
      LngToRGB QBColor(N)
      BGRA(N).B = bred
      BGRA(N).G = bgreen
      BGRA(N).R = bblue
      NColors = 16
   Next N
   Case 8, 24
      CenteredPal
   End Select
   ShowPalette
   ShowAllIcons
End Sub

Private Sub LabCulLR_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
   If Button = vbRightButton Then
      Cul = LCul
      LCul = RCul
      RCul = Cul
      LabCulLR(0).BackColor = LCul
      LabCulLR(1).BackColor = RCul
   End If
End Sub

Private Sub LabCulLR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
   Cul = LabCulLR(Index).BackColor
   If Cul >= 0 Then
      LabCulLR(2).BackColor = Cul
      LngToRGB Cul
      LngToRGB Cul
      LabRGB(0) = Str$(bred)
      LabRGB(1) = Str$(bgreen)
      LabRGB(2) = Str$(bblue)
      LabRGB(3) = ""
   End If
End Sub

Private Sub LabTCul_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      LCul = LabTCul.BackColor
      LabCulLR(0).BackColor = LCul
   Else
      RCul = LabTCul.BackColor
      LabCulLR(1).BackColor = RCul
   End If
End Sub

'#### RATES #####################################################

Private Sub CheckRateValues()
' Essential to check that RateValues are not zero
Dim N As Long
   For N = 1 To NumFrames
      If RateValues(N) = 0 Then RateValues(N) = 1
   Next N
   Fill_fraINFO
End Sub

'### Tools ################################################

Private Sub optTools_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Tool = Index
   optTools(Index).Value = True
   If Index <= 9 Then   ' Drawing Tools on PIC
      SelRect_OFF
      optTools(9).ForeColor = 0
      LabTool = optTools(Index).ToolTipText
   ElseIf Index <= 20 Then  ' Index >= 9
      LabTool = ""
      Select Case Index
      Case 10
         ChangeLCulforRCul
      Case 11  ' Roll L/R
         Roll_LR Button
      Case 12  ' Roll U/D
         Roll_UD Button
      Case 13  ' Shift L/R
         Shift_LR Button
      Case 14  ' Shift U/D
         Shift_UD Button
      Case 15  ' Rot 90 R Clockwise/L Anti-clockwise
         Rotate90 Button
      Case 16  ' Flip LR
         Flip_LR
      Case 17  ' Flip UD
         Flip_UD
      Case 18  ' \ Roll
         Roll_LR Button
         Roll_UD Button
      Case 19  '/ Roll
         Roll_LR 3 - Button
         Roll_UD Button
      Case 20
         MirrorLR Button
      End Select
      BackUp
      ShowAllIcons
      cmdUndoRedo(4).Enabled = True

      ' Does transfer XORA() & ANDA() to PIC,PICI & PICM
      ' ie  aPic_MouseUp CInt(picNum), 0, 0, 0, 0
   End If
   PIC.SetFocus
End Sub

Private Sub optTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index >= 10 Then
      optTools(Index).Value = False ' Lift opt button
      PIC.SetFocus
   End If
End Sub


Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Have aPic(picNum)

Dim Cul As Long
Dim k As Long
   
   If Tool <= 17 Then
   
      aMouseDown = True
      aNew = False
      frmDetails.fraINFO.Caption = "STARTED"
      LabFileName = "STARTED "
      
      Select Case Tool
      Case 0         ' Pencil
         PIC_MouseMove Button, Shift, X, Y
      Case Is = 1    ' Line
         LineTool.Visible = True
         LineTool.X1 = (X \ 8) * 8 + 4
         LineTool.Y1 = (Y \ 8) * 8 + 4
         LineTool.X2 = (X \ 8) * 8 + 4
         LineTool.Y2 = (Y \ 8) * 8 + 4
         PIC_MouseMove Button, Shift, X, Y
      Case 2, 4, 3, 5  ' Rectangle & Filled Rectangle, ' Oval & Filled Oval
         k = 0
         If Tool = 3 Or Tool = 5 Then k = 1
         With STool(k)
            .Visible = True
            .Left = (X \ 8) * 8 + 4
            .Top = (Y \ 8) * 8 + 4
            .Width = 2
            .Height = 2
            XS0 = .Left
            YS0 = .Top
         End With
      Case 6   ' Flood Fill
         If X >= 4 And X <= 259 Then
         If Y >= 2 And Y <= 257 Then
            If Button = vbLeftButton Then
               Cul = LCul
            ElseIf Button = vbRightButton Then
               Cul = RCul
            Else
               Exit Sub
            End If
            FloodFill X, Y, Cul
         End If
         End If
      Case 7   ' Erase to TC
         PIC_MouseMove Button, Shift, X, Y
      Case 8   ' PickA color
         Cul = PIC.Point(X, Y)
         If Cul > -1 Then
            If Button = vbLeftButton Then
               LCul = Cul
               LabCulLR(0).BackColor = LCul
            ElseIf Button = vbRightButton Then
               RCul = Cul
               LabCulLR(1).BackColor = RCul
            End If
         End If
      Case 9   ' SelRect
         If Button = vbLeftButton Then
            shpSelRect2.Visible = False
            TimerSR.Enabled = False
            With shpSelRect
               .Left = 8 * (X \ 8) + 3
               .Top = 8 * (Y \ 8) + 3
               .Width = 2
               .Height = 2
               XS0 = .Left
               YS0 = .Top
               .Visible = True
            End With
            aSelect = True
            shpSelRect.Visible = True
            optTools(9).ForeColor = vbRed
         End If
      Case 10  ' Change Left for Right Color
      Case 11  ' Roll L/R
      Case 12  ' Roll U/D
      Case 13  ' Shift LR
      Case 14  ' Shift UD
      Case 15  ' Rotate 90
      Case 16  ' Flip LR
      Case 17  ' Flip UD
      
      End Select
   End If
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim k As Long
Dim wh As Long
Dim ix As Long, iy As Long
   If X > 255 Then X = 255
   If Y > 255 Then Y = 255
   LabXY = Str$(X \ 8) & "," & Str$(Y \ 8)
   ix = X
   iy = Y
   
   If aMouseDown Then
      If ix >= 4 And ix <= 259 Then
      If iy >= 2 And iy <= 257 Then
         If Button = vbLeftButton Then
            Cul = LCul
         ElseIf Button = vbRightButton Then
            Cul = RCul
         Else
            Exit Sub
         End If
         
         Select Case Tool
         Case 0   ' Pencil
            X = (X \ 8) * 8 + 1
            Y = (Y \ 8) * 8 + 1
            SingleDot X, Y, Cul
         Case 1   ' LineTool
            LineTool.X2 = (X \ 8) * 8 + 4
            LineTool.Y2 = (Y \ 8) * 8 + 4
         Case 2, 4, 3, 5 ' Rectangle & Filled Rectangle, ' Oval & Filled Oval
            k = 0
            If Tool = 3 Or Tool = 5 Then k = 1
            With STool(k)
            
            X = (X \ 8) * 8 + 4
            Y = (Y \ 8) * 8 + 4
               If X >= XS0 Then
                  wh = X - XS0 + 1
                  If wh < 0 Then wh = 0
                  .Width = wh
               Else  ' X < XS0
                  .Left = X
                  wh = XS0 - X + 2
                  If wh < 0 Then wh = 0
                  .Width = wh
               End If
               If Y >= YS0 Then
                  wh = Y - YS0 + 1
                  If wh < 0 Then wh = 0
                  .Height = wh
               Else  ' Y < YS0
                  .Top = Y
                  wh = YS0 - Y + 2
                  If wh < 0 Then wh = 0
                  .Height = wh
               End If
            End With
         Case 6   ' Solid Fill
         Case 7   ' Erase to TC
            X = (X \ 8) * 8 + 1
            Y = (Y \ 8) * 8 + 1
            SingleDot X, Y, TransparentColor
         Case 8   ' PickA color
         Case 9   ' SelRect
            If Button = vbLeftButton Then
               With shpSelRect
               X = (X \ 8) * 8 + 4
               Y = (Y \ 8) * 8 + 4
                  If X >= XS0 Then
                     wh = X - XS0 + 1
                     If wh < 0 Then wh = 0
                     .Width = wh
                  Else  ' X < XS0
                     .Left = X
                     wh = XS0 - X + 2
                     If wh < 0 Then wh = 0
                     .Width = wh
                  End If
                  If Y >= YS0 Then
                     wh = Y - YS0 + 1
                     If wh < 0 Then wh = 0
                     .Height = wh
                  Else  ' Y < YS0
                     .Top = Y
                     wh = YS0 - Y + 2
                     If wh < 0 Then wh = 0
                     .Height = wh
                  End If
               End With
            End If
         Case 10  ' Change Left for Right Color
         Case 11  ' Roll L/R
         Case 12  ' Roll U/D
         Case 13  ' Shift LR
         Case 14  ' Shift UD
         Case 15  ' Rotate 90
         Case 16  ' Flip LR
         Case 17  ' Flip UD
         End Select
      End If
      End If
   Else  ' Not MouseDown
      If Tool = 8 Then ' PickA color
         Cul = PIC.Point(X, Y)
         If Cul >= 0 Then
            'N = X \ 6 + (Y \ 6) * 16
            LabCulLR(2).BackColor = Cul
            LngToRGB Cul
            LabRGB(0) = Str$(bred)
            LabRGB(1) = Str$(bgreen)
            LabRGB(2) = Str$(bblue)
            LabRGB(3) = "" 'Str$(N)
         End If
      End If
   End If
End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Have aPic(picNum)
Dim Cul As Long

      If X > 255 Then X = 255
      If Y > 255 Then Y = 255
   
   If aMouseDown Then
   
      If Button = vbLeftButton Then
         Cul = LCul
      ElseIf Button = vbRightButton Then
         Cul = RCul
      Else
         aMouseDown = False
         Exit Sub
      End If
      Select Case Tool
      Case 0   ' Pencil
      Case 1   ' LineTool
         With LineTool
            LineTool.X2 = (X \ 8) * 8 + 4
            LineTool.Y2 = (Y \ 8) * 8 + 4
            DrawLine .X1, .Y1, .X2, .Y2, Cul
            .Visible = False
         End With
      Case 2, 4  ' Rectangle & Filled Rectangle
         With STool(0)
            DrawRectangle .Left, .Top, .Left + .Width - 1, .Top + .Height - 1, Cul
           .Visible = False
         End With
      Case 3, 5  ' Oval & Filled Oval
         With STool(1)
            DrawOval .Left, .Top, .Left + .Width - 1, .Top + .Height - 1, Cul
           .Visible = False
         End With
      Case 6   ' Solid Fill
      Case 7   ' Erase to TC
      Case 8   ' PickA color
      Case 9   ' SelRect X,Y
         If Button = vbLeftButton Then
            Place_SelRect X, Y
            shpSelRect2.Visible = True
            TimerSR.Enabled = True
         End If
      Case 10  ' Change Left for Right Color
      Case 11  ' Roll L/R
      Case 12  ' Roll U/D
      Case 13  ' Shift LR
      Case 14  ' Shift UD
      Case 15  ' Rotate 90
      Case 16  ' Flip LR
      Case 17  ' Flip UD
      End Select
      
      ' Transfer XORA() & ANDA() to PIC, PICI, PICM
      ' Transfer_aPic_to_XORA_ANDA
      ' &  Transfers XORA() & ANDA() to PIC, PICI, PICM
      aPic_MouseUp CInt(picNum), 0, 0, 0, 0
      
      If Tool <= 7 Then
         BackUp   ' ie Only for Drawing Tools
         ShowAllIcons
         cmdUndoRedo(4).Enabled = True
      End If
      
      PIC.Refresh
      LabTool = optTools(CInt(Tool)).ToolTipText & Str$(Tool)
   End If
   
   aMouseDown = False
   
End Sub

Private Sub Place_SelRect(X As Single, Y As Single)
Dim ixL As Long
Dim iyT As Long
Dim ixR As Long
Dim iyB As Long
Dim ixdiff As Long
Dim iydiff As Long
Dim W As Long, H As Long

   If X < 0 Then
      shpSelRect.Width = shpSelRect.Width - 3
      X = 0
   End If
   If Y < 0 Then
      shpSelRect.Height = shpSelRect.Height - 3
      Y = 0
   End If

   If X > 255 Then
      shpSelRect.Width = shpSelRect.Width - 3
      X = 254 '5
   End If
   If Y > 255 Then
      shpSelRect.Height = shpSelRect.Height - 3
      Y = 254 '5
   End If
   
   ixL = 8 * (XS0 \ 8)
   ixdiff = XS0 - ixL
   If X < XS0 Then
      If XS0 > 255 Then XS0 = 255
      ixL = 8 * (X \ 8)
      ixdiff = X - ixL
   End If
   iyT = 8 * (YS0 \ 8)
   iydiff = YS0 - iyT
   If Y < YS0 Then
      iyT = 8 * (Y \ 8)
      iydiff = Y - iyT
   End If
   
   W = shpSelRect.Width + ixdiff
   W = 8 * (W \ 8)
   If W < 2 Then W = 2
   H = shpSelRect.Height + iydiff
   H = 8 * (H \ 8)
   If H < 2 Then H = 2
   ixR = ixL + W
   iyB = iyT + H
   
   ' Rect coords for XORA .. are
   ' (ixL \ 8)+1,(iyT \ 8)+1
   '
   '                 (ixR \ 8)+1, (iyB \ 8)+1
   ' Selection coords
   'Public ixs1 As Long, iys1 As Long
   'Public ixs2 As Long, iys2 As Long
   ixs1 = (ixL \ 8) + 1
   iys1 = (iyT \ 8) + 1
   ixs2 = (ixR \ 8) + 1
   If ixs2 > 32 Then ixs2 = 32
   iys2 = (iyB \ 8) + 1
   If iys2 > 32 Then iys2 = 32
   
   iys2 = 33 - ((iyT \ 8) + 1)
   iys1 = 33 - ((iyB \ 8) + 1)
   
   ' Check 1-32, 1-32
   LabTool = Str$(ixs1) & "," & Str$(iys1) & "---" & Str$(ixs2) & "," & Str$(iys2)
         
   With shpSelRect
      .Left = ixL
      .Top = iyT
      .Width = W + 9
      .Height = H + 9
   End With
   ' Match shpSelRect2, background white
   With shpSelRect2
      .Left = shpSelRect.Left
      .Top = shpSelRect.Top
      .Width = shpSelRect.Width
      .Height = shpSelRect.Height
   End With
End Sub

Private Sub SingleDot(ByVal X As Single, ByVal Y As Single, Cul As Long)
' Have aPic(picNum)
Dim ixx As Long, iyy As Long
   
   PIC.Line (X, Y)-(X + 6, Y + 6), Cul, BF
   
   ixx = X \ 8
   iyy = Y \ 8
   aPic(picNum).PSet (ixx, iyy), Cul
End Sub

Private Sub DrawLine(ByVal ix1 As Long, ByVal iy1 As Long, ByVal ix2 As Long, ByVal iy2 As Long, ByVal Cul As Long)
   SingleDot ix1, iy1, Cul
   If (ix2 = ix1) And (iy2 = iy1) Then Exit Sub
   ix1 = (ix1) \ 8
   iy1 = (iy1) \ 8
   ix2 = (ix2) \ 8
   iy2 = (iy2) \ 8
   PICT.Cls
   PICT.PSet (ix1, iy1), 0
   PICT.Line (ix1, iy1)-(ix2, iy2), 0
   PICT.PSet (ix2, iy2), 0
   PICT.Refresh
   Scan_PICT Cul
End Sub

Private Sub DrawRectangle(ByVal ix1 As Long, ByVal iy1 As Long, ByVal ix2 As Long, ByVal iy2 As Long, ByVal Cul As Long)
   SingleDot ix1, iy1, Cul
   If (ix2 = ix1) And (iy2 = iy1) Then Exit Sub
   ix1 = (ix1) \ 8
   iy1 = (iy1) \ 8
   ix2 = (ix2) \ 8
   iy2 = (iy2) \ 8
   PICT.Cls
   If Tool = 2 Then ' Rectangle
      PICT.Line (ix1, iy1)-(ix2, iy2), 0, B
   Else  'Filled Rectangle
      PICT.Line (ix1, iy1)-(ix2, iy2), 0, BF
   End If
   PICT.Refresh
   Scan_PICT Cul
End Sub

Private Sub DrawOval(ByVal ix1 As Long, ByVal iy1 As Long, ByVal ix2 As Long, ByVal iy2 As Long, ByVal Cul As Long)
Dim xc As Single, yc As Single
Dim za As Single, zB As Single
Dim zrad As Single, zasp As Single
Dim zradx As Single, zrady As Single
   If (ix1 = ix2) And (iy1 = iy2) Then
      SingleDot ix1, iy1, Cul
      Exit Sub
   End If
   ix1 = (ix1) \ 8
   iy1 = (iy1) \ 8
   ix2 = (ix2) \ 8
   iy2 = (iy2) \ 8
   ' 0-31
   If (ix1 = ix2) And (iy1 = iy1) Then
      SingleDot ix1 * 8, iy1 * 8, Cul
      Exit Sub
   End If
   xc = (ix1 + ix2) / 2
   yc = (iy1 + iy2) / 2
   za = Abs(xc - ix2)
   zB = Abs(yc - iy2)
   zradx = Abs(xc - ix2)
   zrady = Abs(yc - iy2)
   If zradx = 0 Then
      zrad = zrady
      zasp = 100
   ElseIf zradx >= zrady Then
      zrad = zradx
      zasp = zrady / zradx
   Else
      zrad = zrady
      zasp = zrady / zradx
   End If
   PICT.Cls
   If Tool = 5 Then  ' Oval
      PICT.FillColor = 0
      PICT.FillStyle = vbFSSolid
   End If
   PICT.Circle (xc, yc), zrad, 0, , , zasp
   PICT.FillStyle = vbFSTransparent  'Default (Transparent)
   PICT.Refresh
   Scan_PICT Cul
End Sub

Private Sub Scan_PICT(Cul As Long)
' Have picNum
Dim ix As Long, iy As Long
   For iy = 0 To 31
   For ix = 0 To 31
      If PICT.Point(ix, iy) = 0 Then
         aPic(picNum).PSet (ix, iy), Cul
      End If
   Next ix
   Next iy
End Sub

Private Sub FloodFill(X As Single, Y As Single, Cul As Long)
' Have aPic(picNum)
Dim ixx As Long
Dim iyy As Long
   aPic(picNum).DrawStyle = vbSolid
   aPic(picNum).DrawMode = 13
   aPic(picNum).FillColor = Cul
   aPic(picNum).FillStyle = vbFSSolid
   ixx = X \ 8: iyy = Y \ 8 ' aPic() x,y
   ' FLOODFILLSURFACE = 1
   ' Fills with FillColor so long as point surrounded by
   ' Color = aPIC(picNum).Point(X, Y)
   ExtFloodFill aPic(picNum).hDC, ixx, iyy, aPic(picNum).Point(ixx, iyy), FLOODFILLSURFACE
   aPic(picNum).FillStyle = vbFSTransparent  'Default (Transparent)
   aPic(picNum).Refresh
   
   'Transfer_aPic_to_XORA_ANDA
End Sub

Public Sub ShowAllIcons()
Dim ix As Long
Dim iy As Long
Dim N As Long
   For N = 1 To NumFrames
      aPic(N).Cls
      For iy = 1 To 32
      For ix = 1 To 32
         If ANDA(ix, iy, N) = 0 Then   ' Show image
            aPic(N).PSet (ix - 1, 32 - iy), XORA(ix, iy, N)
         End If
      Next ix
      Next iy
      aPic(N).Refresh
   Next N
   aPic_MouseUp CInt(picNum), 0, 0, 0, 0
End Sub


'#### Color #############################################################

Private Sub picPAL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim N As Long
   Cul = picPAL.Point(X, Y)
   If Cul >= 0 Then
      N = X \ 6 + (Y \ 6) * 16
      LabCulLR(2).BackColor = Cul
      LngToRGB Cul
      LabRGB(0) = Str$(bred)
      LabRGB(1) = Str$(bgreen)
      LabRGB(2) = Str$(bblue)
      LabRGB(3) = Str$(N)
   End If
End Sub

Private Sub picPAL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
   Cul = picPAL.Point(X, Y)
   If Cul >= 0 Then
      If Button = vbLeftButton Then
         LCul = Cul
         LabCulLR(0).BackColor = LCul
      ElseIf Button = vbRightButton Then
         RCul = Cul
         LabCulLR(1).BackColor = RCul
      End If
   End If
End Sub

Public Sub ShowPalette()
   picPAL.Top = PIC.Top - 3
   picPAL.Left = 160
   Select Case BPP
   Case 1
      picPAL.Width = 14 * STX
      picPAL.Height = 8 * STY
   Case 4
      picPAL.Width = 98 * STX
      picPAL.Height = 8 * STY
   Case 8, 24
      picPAL.Width = 98 * STX
      picPAL.Height = 98 * STY
   End Select
   DisplayPalette
End Sub

Public Sub ShowBPP24Palette()
   picPAL.Top = PIC.Top - 3
   picPAL.Left = 160
   picPAL.Width = 98 * STX
   picPAL.Height = 98 * STY
   DisplayPalette
End Sub

Public Sub DisplayPalette()
Dim N As Long
Dim px As Long
Dim py As Long
   picPAL.Cls
   N = 0
   For py = 0 To 91 Step 6
   For px = 0 To 91 Step 6
      With BGRA(N)
      picPAL.Line (px, py)-(px + 6, py + 6), RGB(.R, .G, .B), BF
      N = N + 1
      If N > NColors - 1 Then Exit For
      End With
   Next px
      If N > NColors - 1 Then Exit For
   Next py
End Sub

Private Sub picGD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'LabXY = Str$(X) & Str$(Y)
End Sub

Private Sub picGD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' On fraBG
Dim Cul As Long
Dim ix As Long, iy As Long
   Cul = picGD.Point(X, Y)
   If Cul <> -1 Then
      shpGD.Left = X + picGD.Left
      GridCul = Cul
      For ix = 0 To kspace * 33 + 1 Step kspace
         PIC.Line (ix, 0)-(ix, PIC.Height), GridCul
      Next ix
      For iy = 0 To kspace * 33 + kspace Step kspace
         PIC.Line (0, iy)-(PIC.Width, iy), GridCul
      Next iy
      Shape2.FillColor = GridCul
   End If
End Sub

Private Sub PICI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show Image color
Dim Cul As Long
   Cul = PICI.Point(X, Y)
   If Cul >= 0 Then
      LngToRGB Cul
      LabCulLR(2).BackColor = Cul
      LabRGB(0) = Str$(bred)
      LabRGB(1) = Str$(bgreen)
      LabRGB(2) = Str$(bblue)
      LabRGB(3) = ""
   End If
End Sub


'####  Frames ##############################################################

Private Sub aPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show image in frame color
Dim Cul As Long
   If Tool = 8 Then ' PickA color
      Cul = aPic(Index).Point(X, Y)
      If Cul >= 0 Then
         LabCulLR(2).BackColor = Cul
         LngToRGB Cul
         LabRGB(0) = Str$(bred)
         LabRGB(1) = Str$(bgreen)
         LabRGB(2) = Str$(bblue)
         LabRGB(3) = "" 'Str$(N)
      End If
   End If
End Sub

Private Sub aPic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' XORA(ix,iy,Index) 32x32  PIC(IX,IY) 32(8x8)32
Dim ix As Long
Dim iy As Long
Dim kx As Long
Dim ky As Long
Dim Cul As Long
   ' Color rate buttons
   picNum = Index
   frmDetails.LRates_MouseUp Index, Button, 0, 0, 0
   
   PIC.Cls  ' to RGB(224,224,224)
   PICM.Cls
   
   ' aPic() color -> XORA()
   ' IF aPic() color = TransparentColor Then
   ' ANDA() = 1 Else ANDA() = 0
   Transfer_aPic_to_XORA_ANDA
   
   For iy = 1 To 32
   For ix = 1 To 32
      Cul = XORA(ix, iy, picNum)
      kx = 8 * (ix - 1) + 1 ' 1-249
      ky = 8 * (iy - 1) + 1 ' 1-249
      ky = 250 - ky

'      If ANDA(ix, iy, picNum) = 0 Or _
'      XORA(ix, iy, picNum) <> ImageTransparentColor Then
'      'Then   ' Show image      ??
         
      PIC.Line (kx, ky)-(kx + 6, ky + 6), Cul, BF
      
      'If XORA(ix, iy, picNum) <> ImageTransparentColor Then
      If ANDA(ix, iy, picNum) = 0 Then
         'ANDA(ix, iy, picNum) = 1
         'PIC.Line (kx, ky)-(kx + 6, ky + 6), Cul, BF
         PICI.PSet (ix - 1, 32 - iy), Cul
         PICM.PSet (ix - 1, 32 - iy), vbBlack
      Else  ' Transparent
         'ANDA(ix, iy, picNum) = 0
         PICI.PSet (ix - 1, 32 - iy), ImageTransparentColor
         PICM.PSet (ix - 1, 32 - iy), vbWhite
      End If
   Next ix
   Next iy
   LabPicNum = LTrim$(Str$(picNum))
   ' Re-draw Grid
   For ix = 0 To kspace * 33 + 1 Step kspace
      PIC.Line (ix, 0)-(ix, PIC.Height), GridCul
   Next ix
   For iy = 0 To kspace * 33 + kspace Step kspace
      PIC.Line (0, iy)-(PIC.Width, iy), GridCul
   Next iy
   
   CheckSwap
   
   If Button = vbLeftButton Then
      RestartBackUps
   End If
   If Button = vbRightButton Then
      PopupMenu Me.mnuPopUp
   End If
End Sub

Private Sub Transfer_aPic_to_XORA_ANDA()
' Have picNum
Dim ixx As Long
Dim iyy As Long
Dim PCul As Long
   ' Enter values in XORA(),ANDA() from aPic() whence
   ' PIC.MouseUp() will transfer aPic() to PIC, PICI & PICM
   For iyy = 0 To 31
   For ixx = 0 To 31
      PCul = aPic(picNum).Point(ixx, iyy)
      XORA(ixx + 1, 32 - iyy, picNum) = PCul
      If PCul = TransparentColor Then
         ANDA((ixx + 1), 32 - iyy, picNum) = 1
      Else
         ANDA((ixx + 1), 32 - iyy, picNum) = 0
      End If
   Next ixx
   Next iyy
End Sub


'#### PopUp menus #############################################

Private Sub mnuCP_Click(Index As Integer)
' Have picNum
   Select Case Index
   Case 0  ' Copy
      CopyXORA_ANDA
      mnuCP(1).Enabled = True ' Paste
   Case 1  'Paste
      PasteXORA_ANDA
   Case 2  ' Clear to TC
      ClearXORA_ANDA
   Case 3   ' Capture
      Open_Image
   Case 4   ' Remove frame
      If NumFrames > 1 Then
         Delete_picNum
         Fill_fraINFO
      End If
   Case 5   ' Swap
   End Select
   RestartBackUps
   ShowAllIcons
End Sub

Private Sub mnuCPSub_Click(Index As Integer)
   Swap Index
   ShowAllIcons
End Sub

Private Sub CheckSwap()
   If NumFrames > 1 Then
      mnuCP(5).Enabled = True
      
      'Case 0,3   ' Swap with Next, Swap with Last
      If picNum < NumFrames Then
         mnuCPSub(0).Enabled = True
         mnuCPSub(3).Enabled = True
      Else
         mnuCPSub(0).Enabled = False
         mnuCPSub(3).Enabled = False
      End If
      'Case 1,2   ' Swap with Previous, Swap with First
      If picNum > 1 Then
         mnuCPSub(1).Enabled = True
         mnuCPSub(2).Enabled = True
      Else
         mnuCPSub(1).Enabled = False
         mnuCPSub(2).Enabled = False
      End If
   Else
      mnuCP(5).Enabled = False
   End If
End Sub



'#### FILE MENUS ############################################

Private Sub mnuFileSet_Click(Index As Integer)
   
   aPlay = False
   
   If AniTest Then
      RestoreOldCursor
      AniTest = False
   End If
   
   Select Case Index
   Case 0   ' New
      InitNEW
   Case 1   ' Open
      Open_ani
   Case 2   ' Save
      Save_ani
   Case 3   ' Break
   Case 4   ' Load JASC PAL
      Load_Palette
   Case 6   ' Capture
      Save_Palette
   Case 7   ' Break
   Case 8   ' Open cur cursor picNum
      Open_cur 2
   Case 9   ' Save cur cursor picNum
      Save_cur
   Case 10   ' Break
   Case 11   ' Open ico picNum
      Open_cur 1
   Case 12   ' Break
   Case 13  ' Exit
      Form_QueryUnload 0, 0
   End Select
End Sub

Private Sub Open_ani()
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   Set CommonDialog1 = New OSDialog
   Title$ = "Open Ani Cursor"
   Filt$ = "Open ani (*.ani)|*.ani"
   If FileSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(FileSpec$, "\")
      InDir$ = Left$(FileSpec$, p)
   End If
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   ' FIndex = 1 *.ani file
   Set CommonDialog1 = Nothing
   
   If LenB(FileSpec$) <> 0 Then
      
      READ_ANI_FILE FileSpec$
      If FileSpec$ = "" Then
         MsgBox "Problem with reading ani file", vbCritical, "Reading ani"
         Exit Sub
      End If
      LabTCul.BackColor = TransparentColor
      aNew = False
      LabFileName = FileSpec$ & " "
      icontype = 0
   Else
      Exit Sub
   End If
   ' Have:
   ' ANI$  - whole file
   ' AniTitle$
   ' AniAuthor$
   ' NumFrames
   ' RateValues(1 to NumFrames)
   ' HotX = .iHotX
   ' HotY = .iHotY
   ' BPP = .NBPP
   ' tWidth = .NWidth
   ' tHeight = .NHeight \ 2
   
   SelRect_OFF
   optTools(9).ForeColor = 0
   aCopy = False     ' Copy/Paste
   mnuCP(1).Enabled = False   ' Paste disabled
   picNum = 1
   RestartBackUps
   
   aMouseDown = False            ' No PIC action yet
   LabFN(picNumHighlighted).BackColor = vbWhite
   picNumHighlighted = 1         ' For LRates() & LabFN()
   LabFN(1).BackColor = vbYellow
   
   frmDetails.fraINFO.Caption = "INFO"
   
   Fill_fraINFO
   
   If BPP = 24 Then
      ExtractRoughPalette
   End If
   
   ShowPalette
   ShowAllIcons   ' Does aPic_MouseUp Cint(picNum), 0, 0, 0, 0
End Sub

Private Sub Save_ani()
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long

   Set CommonDialog1 = New OSDialog
   Title$ = "Save Ani Cursor"
   Filt$ = "Save ani (*.ani)|*.ani"
   If FileSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(FileSpec$, "\")
      InDir$ = Left$(FileSpec$, p)
   End If
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   ' FIndex = 1 *.ani file
   Set CommonDialog1 = Nothing
   
   If LenB(FileSpec$) <> 0 Then
      aNew = False
      AniTitle$ = frmDetails.txtTitle.Text
      AniAuthor$ = frmDetails.txtAuthor.Text
      CheckRateValues
      FixExtension FileSpec$, ".ani"
      SAVE_ANI_FILE FileSpec$
   End If
End Sub

Private Sub cmdTestAni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   AniTest = Not AniTest
   If AniTest Then
      AniTitle$ = frmDetails.txtTitle.Text
      AniAuthor$ = frmDetails.txtAuthor.Text
      aniFSpec$ = PathSpec$ & "~~ana~~.ani"
      CheckRateValues
      SAVE_ANI_FILE aniFSpec$
      DoEvents
      ShowNewCursor aniFSpec$
      DoEvents
   Else
      RestoreOldCursor
      DoEvents
      On Error Resume Next
      Kill aniFSpec$
      DoEvents
   End If

End Sub

Private Sub Open_cur(icontype As Long)
Dim ix As Long
Dim iy As Long
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long

   Set CommonDialog1 = New OSDialog
   If icontype = 1 Then
      Title$ = "Open 32x32 icon"
      Filt$ = "Open ico (*.ico)|*.ico"
   Else
      Title$ = "Open 32x32 Cursor"
      Filt$ = "Open cur (*.cur)|*.cur"
   End If
   If FileSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(FileSpec$, "\")
      InDir$ = Left$(FileSpec$, p)
   End If
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   Set CommonDialog1 = Nothing
   
   If LenB(FileSpec$) <> 0 Then
      aNew = False
      READ_CUR_FILE FileSpec$, icontype
      
      SelRect_OFF
      
      optTools(9).ForeColor = 0
      aCopy = False     ' Copy/Paste
      mnuCP(1).Enabled = False   ' Paste disabled
      
      RestartBackUps
      
      aMouseDown = False            ' No PIC action yet
      frmDetails.fraINFO.Caption = "INFO"
      
      Fill_fraINFO
      ShowPalette
      
      ' Show cursor image
      aPic(picNum).Cls
      
      ' Test if all ANDA()s are 1
      ' Some weird MS cursors !!
      p = 1
      For iy = 1 To 32
      For ix = 1 To 32
         If ANDA(ix, iy, picNum) = 0 Then
            p = 0
            Exit For
         End If
      Next ix
      If p = 0 Then Exit For
      Next iy
      
      If p = 1 Then  ' All ANDA()s are 1
         ' Change ANDA() as XORA() is <> 0
         For iy = 1 To 32
         For ix = 1 To 32
            If XORA(ix, iy, picNum) <> 0 Then
               ANDA(ix, iy, picNum) = 0
            End If
         Next ix
         Next iy
      End If
      
      For iy = 1 To 32
      For ix = 1 To 32
         If ANDA(ix, iy, picNum) = 0 Then
            aPic(picNum).PSet (ix - 1, 32 - iy), XORA(ix, iy, picNum)
         End If
      Next ix
      Next iy
         
      aPic(picNum).Refresh
      ' Transfer to PIC, PICI & PICM
      aPic_MouseUp CInt(picNum), 0, 0, 0, 0
      
      If BPP = 24 Then
         ExtractRoughPalette
      End If
   End If

End Sub

Private Sub Save_cur()
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long

   Set CommonDialog1 = New OSDialog
   Title$ = "Save Cursor"
   Filt$ = "Save cur (*.cur)|*.cur"
   If FileSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(FileSpec$, "\")
      InDir$ = Left$(FileSpec$, p)
   End If
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   ' FIndex = 1 *.cur file
   Set CommonDialog1 = Nothing
   
   If LenB(FileSpec$) <> 0 Then
      aNew = False
      FixExtension FileSpec$, ".cur"
      SAVE_CUR_FILE FileSpec$
   End If
End Sub

Private Sub cmdPlay_Click()
Dim N As Long
Dim ix As Long, iy As Long
   If NumFrames > 1 Then
      aPlay = Not aPlay
      If aPlay Then
         cmdPlay.Caption = "Stop"
         picPlay.SetFocus
         Do
            For N = 1 To NumFrames
               Form1.picPlay.Cls
               BitBlt picPlay.hDC, 0, 0, 32, 32, aPic(N).hDC, 0, 0, vbSrcCopy
               Sleep RateValues(N) * 15.48 '16.7
               DoEvents
               If Not aPlay Then Exit For
            Next N
         Loop Until Not aPlay
         cmdPlay.Caption = "Play"
         picPlay.SetFocus
         picPlay.Cls
      Else
         cmdPlay.Caption = "Play"
         picPlay.SetFocus
         picPlay.Cls
      End If
   End If
End Sub
Private Sub picPlay_Click()
      aPlay = Not aPlay
End Sub


Private Sub cmdTestCursor_Click()
   AniTest = Not AniTest
   If AniTest Then
      curFSpec$ = PathSpec$ & "~~cur~~.cur"
      SAVE_CUR_FILE curFSpec$
      DoEvents
      ShowNewCursor curFSpec$
      MousePointer = 99
      DoEvents
   Else
      RestoreOldCursor
      MousePointer = 0
      DoEvents
      On Error Resume Next
      Kill curFSpec$
      DoEvents
   End If
End Sub

' VB method
' but gives black cursor in IDE
'Private Sub cmdTestCursor_Click()
''Change MouseCursor
'If CurFSpec$ <> "" And LCase$(Right(CurFSpec$, 4)) = ".cur" Then
'   MouseIcon = LoadPicture(CurFSpec$)
'   MousePointer = 99
'End If
'End Sub

'Private Sub cmdResetCursor_Click()
'MousePointer = 0
'End Sub

Private Sub ExtractRoughPalette()
Dim N As Long
Dim ix As Long
Dim iy As Long
Dim Cul As Long
   NColors = 256
   ReDim Preserve BGRA(0 To 255)
   N = 0
   For iy = 1 To 32 Step 2
   For ix = 1 To 32 Step 2
         Cul = XORA(ix, iy, picNum)
         If Cul <> TransparentColor Then
            LngToRGB Cul
            BGRA(N).B = bblue
            BGRA(N).G = bgreen
            BGRA(N).R = bred
            N = N + 1
         End If
   Next ix
   Next iy
End Sub

Private Sub Load_Palette()
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim ix As Long, iy As Long
Dim cul0 As Long, cul1 As Long
' For re-mapping palettes
Dim N As Long
Dim bR As Byte, bG As Byte, BB As Byte
   Set CommonDialog1 = New OSDialog
   Title$ = "Load JASC_PAL"
   Filt$ = "Open pal (*.pal)|*.pal"
   If PalSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(PalSpec$, "\")
      InDir$ = Left$(PalSpec$, p)
   End If
   CommonDialog1.ShowOpen PalSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   ' FIndex = 1 *.pal file
   Set CommonDialog1 = Nothing
   
   If LenB(PalSpec$) <> 0 Then
      If BPP < 24 Then
         ' Get indexes and then re-map for BPP<24
         ReDim Indexes(32, 32, NumFrames)
         For N = 1 To NumFrames
            Select Case BPP
            Case 1
               cul0 = RGB(BGRA(0).B, BGRA(0).G, BGRA(0).R)
               cul1 = RGB(BGRA(1).B, BGRA(1).G, BGRA(1).R)
               For iy = 1 To 32
               For ix = 1 To 32
                  If XORA(ix, iy, N) = cul0 Then
                     Indexes(ix, iy, N) = 0
                  Else
                     Indexes(ix, iy, N) = 1
                  End If
               Next ix
               Next iy
            Case 4
            Dim bIndex As Byte
               For iy = 1 To 32
               For ix = 1 To 32 Step 2
                  cul0 = XORA(ix, iy, N)
                  cul1 = XORA(ix + 1, iy, N)
                  bIndex = GetIndex2Nybbles(cul0, cul1)
                  Indexes(ix, iy, N) = bIndex And &HF
                  Indexes(ix + 1, iy, N) = (bIndex And &HF0) \ 16
               Next ix
               Next iy
            Case 8
               For iy = 1 To 32
               For ix = 1 To 32
                  cul0 = XORA(ix, iy, N)
                  Indexes(ix, iy, N) = GetPalIndexByte(cul0)
               Next ix
               Next iy
            End Select
         Next N
      End If
      
      READ_JASC_PAL PalSpec$
      
      If PalSpec$ = "" Then
         MsgBox "Problem reading Pal file", vbCritical, "Reading JASC_PAL"
         Exit Sub
      End If
      ' Returns with BGRA() filled as BPP
      If BPP = 24 Then
         ShowBPP24Palette
      Else
         ShowPalette
      End If
      ' Re-map palette
      If BPP < 24 Then
         For N = 1 To NumFrames
            For iy = 1 To 32
            For ix = 1 To 32
               bR = BGRA(Indexes(ix, iy, N)).R
               bG = BGRA(Indexes(ix, iy, N)).G
               BB = BGRA(Indexes(ix, iy, N)).B
               XORA(ix, iy, N) = RGB(bR, bG, BB)
            Next ix
            Next iy
         Next N
      End If
      ShowAllIcons
      Erase Indexes()
   End If
End Sub

Private Sub Save_Palette()
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   Set CommonDialog1 = New OSDialog
   Title$ = "Save JASC_PAL"
   Filt$ = "Save pal (*.pal)|*.pal"
   If PalSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(PalSpec$, "\")
      InDir$ = Left$(PalSpec$, p)
   End If
   CommonDialog1.ShowSave PalSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   ' FIndex = 1 *.pal file
   Set CommonDialog1 = Nothing
   
   If LenB(PalSpec$) <> 0 Then
      FixExtension PalSpec$, ".pal"
      SAVE_JASC_PAL PalSpec$
   End If
End Sub

Private Sub Open_Image()
' Capturing
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long

   Set CommonDialog1 = New OSDialog
   Title$ = "Extract from image"
   '*' Filt$ = "Open BMP (*.bmp)|*.bmp|Open GIF (*.gif)|*.gif|Open JPG (*.jpg)|*.jpg"
   Filt$ = "Open Pic |*.bmp;*.gif;*.jpg"
   If ImageSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(ImageSpec$, "\")
      InDir$ = Left$(ImageSpec$, p)
   End If
   CommonDialog1.ShowOpen ImageSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   '*' FIndex = 1 *.bmp file
   '*' FIndex = 2 *.gif file
   '*' FIndex = 3 *.jpg file
   Set CommonDialog1 = Nothing
   
   If LenB(ImageSpec$) <> 0 Then
      ImageType = FIndex
      frmCapture.Show vbModeless, Me
      aNew = False
      'If BPP = 24 Then ExtractRoughPalette  ' Done in Capture
      icontype = 0   ' possible ani
   End If
End Sub

'############################################################

Public Sub Visibility()
' Public NumFrames
Dim k As Long
On Error Resume Next
      For k = 1 To NumFrames
         aPic(k).Visible = True
         LabFN(k).Visible = True
      Next k
      For k = NumFrames + 1 To 64
         aPic(k).Visible = False
         LabFN(k).Visible = False
      Next k
   
   With frmDetails
      For k = 1 To NumFrames
         .LRates(k).Visible = True
      Next k
      For k = NumFrames + 1 To 64
         .LRates(k).Visible = False
      Next k
   End With
End Sub

Private Sub SizeAll()
Dim k As Long
Dim T As Long, L As Long
Dim G As Long, D As Long, E As Long
Dim ix As Long
   Me.Cls
   aPic(1).BackColor = TransparentColor
   
   aPic(1).Width = 32
   aPic(1).Height = 32
   PICM.Width = 32 * STX
   PICM.Height = 32 * STY
   PICI.Width = 32 * STX
   PICI.Height = 32 * STY
   
   'Load aPic frames & position
   For k = 2 To 64 ' Max number of frames - 1
      Load aPic(k)
      Load LabFN(k)
      LabFN(k).ZOrder
      LabFN(k).Visible = True
   Next k
   LabFN(1).Top = aPic(1).Top + 32
   ' Position aPic()s & print frame numbers
   T = aPic(1).Top
   L = aPic(1).Left
   G = 38
   D = G - 4
   E = 10
   
   CurrentY = T + D
   CurrentX = L + 8
   Print "0";
      
   For k = 2 To 64
      With aPic(k)
         Select Case k
         Case Is <= 16
            L = L + G
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         Case 17
            L = aPic(1).Left
            T = aPic(1).Top + G + E
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         Case Is <= 32
            L = L + G
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         Case 33
            L = aPic(1).Left
            T = aPic(1).Top + 2 * (G + E)
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         Case Is <= 48
            L = L + G
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         Case 49
            L = aPic(1).Left
            T = aPic(1).Top + 3 * (G + E)
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         Case Is <= 64
            L = L + G
            .Left = L
            .Top = T
            With LabFN(k)
               .Left = L
               .Top = T + 32
               .Caption = k
               .Visible = True
            End With
         End Select
         .Visible = True
         .Refresh
      End With
   Next k
   
   ' Load & postion LRates()
   With frmDetails
      T = .LRates(1).Top
      L = .LRates(1).Left
      E = 300 ' LRate spacing
      .LRates(1).Caption = ""
      For k = 2 To 16
            Load .LRates(k)
            .LRates(k).Top = T
            .LRates(k).Left = L + ((k - 1) * E)
            .LRates(k).Visible = True
      Next k
      T = T + 225
      For k = 17 To 32
            Load .LRates(k)
            .LRates(k).Top = T
            .LRates(k).Left = L + (k - 17) * E
            .LRates(k).Visible = True
      Next k
      T = T + 225
      For k = 33 To 48
            Load .LRates(k)
            .LRates(k).Top = T
            .LRates(k).Left = L + (k - 33) * E
            .LRates(k).Visible = True
      Next k
      T = T + 225
      For k = 49 To 64
            Load .LRates(k)
            .LRates(k).Top = T
            .LRates(k).Left = L + (k - 49) * E
            .LRates(k).Visible = True
      Next k
   End With
   
   For k = 1 To 64
      RateValues(k) = 1
   Next k
   
   ' Grid & back-display color frame fraBG
   E = 50
   For ix = 0 To 1500 Step 200
      picGD.Line (ix, 0)-(ix + 200, picGD.Height), RGB(E, E, E), BF
      E = E + 32
      If E > 255 Then E = 255
   Next ix
End Sub

Private Sub DrawGridOnPIC(Cul As Long, BacCul As Long)
' Public kspace
Dim iy As Long
Dim ix As Long
   PIC.Cls
   PIC.BackColor = BacCul
   PIC.Left = PIC.Left + 4
   For ix = 0 To kspace * 33 + 1 Step kspace
      PIC.Line (ix, 0)-(ix, PIC.Height), Cul
   Next ix
   For iy = 0 To kspace * 33 + kspace Step kspace
      PIC.Line (0, iy)-(PIC.Width, iy), Cul
   Next iy
   Shape2.BackColor = BacCul
   LineTool.BorderColor = BacCul
   STool(0).BorderColor = BacCul
   STool(1).BorderColor = BacCul
End Sub

Private Sub txtHelpTEXT()
Dim A$
   A$ = ""
   A$ = A$ & " Im  = Image" & vbCrLf
   A$ = A$ & " Ma  = Mask" & vbCrLf
   A$ = A$ & " L   = Left color" & vbCrLf
   A$ = A$ & " R   = Right color" & vbCrLf
   A$ = A$ & "       L/R-Click on palette to set." & vbCrLf
   A$ = A$ & "       R-Click on L or R to swap." & vbCrLf
   A$ = A$ & " TC  = Fixed transparent color." & vbCrLf
   A$ = A$ & "       L/R-Click to set." & vbCrLf
   A$ = A$ & " GD  = Grid & Frame-display" & vbCrLf
   A$ = A$ & "       backcolor." & vbCrLf
   txtHelp(0).Text = A$
   
   
   A$ = ""
   A$ = A$ & " SR = Select Rectangle." & vbCrLf
   A$ = A$ & " These tools will also" & vbCrLf
   A$ = A$ & " act on a selection." & vbCrLf
   txtHelp(1).Text = A$
   
   A$ = ""
   A$ = A$ & " D = De-select any Selection Rectangle" & vbCrLf
   A$ = A$ & "     & Tools." & vbCrLf
   A$ = A$ & " CB = Clear Backups." & vbCrLf
   A$ = A$ & " UA = Undo All." & vbCrLf
   A$ = A$ & " T  = Total number of backups" & vbCrLf
   A$ = A$ & " N  = Number of current backup." & vbCrLf
   txtHelp(2).Text = A$
   
   A$ = ""
   A$ = A$ & " BPP can" & vbCrLf
   A$ = A$ & " only be" & vbCrLf
   A$ = A$ & " changed" & vbCrLf
   A$ = A$ & " directly" & vbCrLf
   A$ = A$ & " after New." & vbCrLf
   frmDetails.txtHelp.Text = A$

End Sub


'#### QUIT ##################################################

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form
Dim res As Long
   aPlay = False
   If AniTest Then
      RestoreOldCursor
      AniTest = False
   End If
   
' This would work if the system cursor filespec is known
'    aniFSpec$ = "C:\WINDOWS\Cursors\appstart.ani"
'    Call SetSystemCursor(LoadCursorFromFile(aniFSpec$), OCR_NORMAL)
   
   On Error Resume Next
   Kill curFSpec$
   Kill aniFSpec$
   
   If UnloadMode = 0 Then    'Close on Form1 pressed
      res = MsgBox("", vbQuestion + vbYesNo + vbSystemModal, "Quit ?")
      If res = vbNo Then
         Cancel = True
      Else
         Cancel = False
         'Screen.MousePointer = vbDefault
         TimerSR.Enabled = False
         ' Make sure all forms cleared
         For Each Form In Forms
            Unload Form
            Set Form = Nothing
         Next Form
         End
      End If
   End If
End Sub
'############################################################

' Select Rectangle
' See also PIC &mPrivate Sub Place_SelRect
Private Sub Set_SelRect()
   With shpSelRect
      .Visible = False
      .BorderWidth = 1
      .BorderStyle = 3
      .DrawMode = 7
      .BorderColor = vbWhite
   End With
   With shpSelRect2
      .Visible = False
      .BorderWidth = 1
      .BorderStyle = 1
      .DrawMode = 13
      .BorderColor = vbWhite
   End With
End Sub

Private Sub TimerSR_Timer()
   If aSelect Then
      SRCount = 1 - SRCount
      If SRCount = 1 Then
         shpSelRect2.BorderColor = vbWhite
      Else
         shpSelRect2.BorderColor = 0
      End If
   End If
End Sub

Private Sub SelRect_OFF()
   aSelect = False
   shpSelRect.Visible = False
   shpSelRect.BorderStyle = 3
   shpSelRect2.Visible = False
   TimerSR.Enabled = False
End Sub
