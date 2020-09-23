VERSION 5.00
Begin VB.Form frmEffects 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Effects"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   735
      Index           =   3
      Left            =   4140
      MultiLine       =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "frmEffects.frx":0000
      Top             =   3330
      Width           =   3390
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
      Height          =   1635
      Index           =   2
      Left            =   4155
      MultiLine       =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "frmEffects.frx":000D
      Top             =   1440
      Width           =   3750
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
      Height          =   1125
      Index           =   1
      Left            =   4140
      MultiLine       =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "frmEffects.frx":001A
      Top             =   135
      Width           =   3570
   End
   Begin VB.Frame fraActions 
      BackColor       =   &H80000013&
      Caption         =   "Gradate"
      Height          =   3990
      Left            =   105
      TabIndex        =   20
      Top             =   90
      Width           =   1650
      Begin VB.OptionButton optRoller 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Roller ->  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3240
         Width           =   1095
      End
      Begin VB.HScrollBar HSRot 
         Height          =   165
         LargeChange     =   10
         Left            =   105
         Max             =   360
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   1080
      End
      Begin VB.HScrollBar HSReduce 
         Height          =   165
         LargeChange     =   10
         Left            =   105
         Max             =   100
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   885
         Width           =   1080
      End
      Begin VB.OptionButton optGO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GO  L/R"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1125
         Width           =   795
      End
      Begin VB.OptionButton optPepper 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pepper  L/R"
         Height          =   255
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1470
         Width           =   1110
      End
      Begin VB.OptionButton optWave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Wave  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1755
         Width           =   1110
      End
      Begin VB.OptionButton optReverse 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reverse  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2940
         Width           =   1095
      End
      Begin VB.OptionButton optSwirl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swirl  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2025
         Width           =   1110
      End
      Begin VB.OptionButton optSwivel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Squash  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2295
         Width           =   1110
      End
      Begin VB.OptionButton optCopy 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copy  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " Copy frame to Left or Right "
         Top             =   3645
         Width           =   1095
      End
      Begin VB.OptionButton optXFade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "XFade  L/R"
         Height          =   240
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2565
         Width           =   1110
      End
      Begin VB.Label LabRot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "360"
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
         Left            =   1215
         TabIndex        =   37
         Top             =   465
         Width           =   330
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000013&
         Caption         =   "Rotate deg"
         Height          =   210
         Index           =   0
         Left            =   390
         TabIndex        =   36
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000013&
         Caption         =   "Reduce  %"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   35
         Top             =   690
         Width           =   795
      End
      Begin VB.Label LabReduce 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
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
         Left            =   1215
         TabIndex        =   34
         Top             =   870
         Width           =   330
      End
      Begin VB.Label LabQ1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   33
         Top             =   195
         Width           =   165
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   15
         X2              =   1635
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   15
         X2              =   1605
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Label LabQ1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   75
         TabIndex        =   32
         Top             =   1485
         Width           =   165
      End
      Begin VB.Label LabQ1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   75
         TabIndex        =   31
         Top             =   2910
         Width           =   165
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   15
         X2              =   1605
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   0
         X2              =   1635
         Y1              =   2865
         Y2              =   2865
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   15
         X2              =   1650
         Y1              =   3555
         Y2              =   3555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   30
         X2              =   1620
         Y1              =   3570
         Y2              =   3570
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "For 24 bpp only"
      Height          =   3990
      Left            =   1815
      TabIndex        =   3
      Top             =   90
      Width           =   2115
      Begin VB.HScrollBar scrParam 
         Height          =   180
         Index           =   1
         LargeChange     =   10
         Left            =   225
         Max             =   100
         Min             =   -50
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1080
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000013&
         Caption         =   "Selected to Last Frame"
         Height          =   2820
         Left            =   105
         TabIndex        =   4
         Top             =   255
         Width           =   1890
         Begin VB.HScrollBar scrParam 
            Height          =   180
            Index           =   5
            Left            =   120
            Max             =   16
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1080
         End
         Begin VB.HScrollBar scrParam 
            Height          =   165
            Index           =   2
            Left            =   120
            Max             =   3
            Min             =   -3
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1020
            Width           =   1080
         End
         Begin VB.HScrollBar scrParam 
            Height          =   165
            Index           =   0
            LargeChange     =   10
            Left            =   105
            Max             =   100
            Min             =   -50
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   495
            Value           =   1
            Width           =   1080
         End
         Begin VB.HScrollBar scrParam 
            Height          =   165
            Index           =   3
            Left            =   120
            Max             =   0
            Min             =   255
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1515
            Value           =   128
            Width           =   1080
         End
         Begin VB.HScrollBar scrParam 
            Height          =   165
            Index           =   4
            Left            =   120
            Max             =   48
            Min             =   16
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2025
            Value           =   16
            Width           =   1080
         End
         Begin VB.Label LabParam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Index           =   5
            Left            =   1230
            TabIndex        =   44
            Top             =   2520
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "Diffuse"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label LabParam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Index           =   2
            Left            =   1230
            TabIndex        =   16
            Top             =   1005
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "Darken/Brighten"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   15
            Top             =   255
            Width           =   1185
         End
         Begin VB.Label LabParam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
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
            Index           =   0
            Left            =   1215
            TabIndex        =   14
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "Smooth/Sharpen"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label LabParam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "128"
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
            Index           =   3
            Left            =   1230
            TabIndex        =   12
            Top             =   1500
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "Black && White"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "Black && White  Dither"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   1755
            Width           =   1515
         End
         Begin VB.Label LabParam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "16"
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
            Index           =   4
            Left            =   1230
            TabIndex        =   9
            Top             =   2010
            Width           =   330
         End
      End
      Begin VB.Label LabParam 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
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
         Index           =   1
         Left            =   1335
         TabIndex        =   19
         Top             =   3540
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Darken/Brighten Palette"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdAC 
      BackColor       =   &H80000013&
      Caption         =   "&Reset"
      Height          =   360
      Index           =   2
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4170
      Width           =   720
   End
   Begin VB.CommandButton cmdAC 
      BackColor       =   &H80000013&
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4170
      Width           =   705
   End
   Begin VB.CommandButton cmdAC 
      BackColor       =   &H80000013&
      Caption         =   "&Accept"
      Height          =   360
      Index           =   0
      Left            =   465
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4170
      Width           =   705
   End
End
Attribute VB_Name = "frmEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmEffects.frm

' Gradates Acts L/R
' 24 bpp acts to right, picNum to NumFrames

Option Explicit
Option Base 1

'Private DBGRA() As PAL4
Private XORACPY() As Long
Private ANDACPY() As Byte
Private XORATEMP() As Long
Private ANDATEMP() As Byte
Private Intens() As Long

Private Param As Long
Private zParam As Single
Private k As Long
Private Cul As Long
Private N As Long
Private ix As Long
Private iy As Long
Private RR As Long
Private GG As Long
Private BB As Long
Private SR As Long, SG As Long, SB As Long
Private aBlock As Boolean


Private Sub cmdAC_Click(Index As Integer)
Dim i As Integer
   Select Case Index
   Case 0   ' Accept
      Form1.ShowPalette
      Form1.ShowAllIcons
      frmEffectsLeft = frmEffects.Left
      frmEffectsTop = frmEffects.Top
      aEffects = False
      Unload Me
   Case 1   ' Cancel
      CopyMemory BGRA(0), DBGRA(0), NColors * 4
      CopyMemory XORA(1, 1, 1), XORACPY(1, 1, 1), 32 * 32 * NumFrames * 4
      CopyMemory ANDA(1, 1, 1), ANDACPY(1, 1, 1), 32 * 32 * NumFrames
      Erase DBGRA()
      Erase XORACPY()
      Erase ANDACPY()
      Form1.ShowPalette
      Form1.ShowAllIcons
      frmEffectsLeft = frmEffects.Left
      frmEffectsTop = frmEffects.Top
      aEffects = False
      Unload Me
   Case 2  ' Reset
      CopyMemory BGRA(0), DBGRA(0), NColors * 4
      CopyMemory XORA(1, 1, 1), XORACPY(1, 1, 1), 32 * 32 * NumFrames * 4
      CopyMemory ANDA(1, 1, 1), ANDACPY(1, 1, 1), 32 * 32 * NumFrames
      For i = 0 To 5
         SetParams i
      Next i
      Form1.ShowPalette
      Form1.ShowAllIcons
   End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
   
   aPlay = False
   If AniTest Then
      RestoreOldCursor
      AniTest = False
      On Error Resume Next
      Kill aniFSpec$
      DoEvents
   End If
   
   txtHelpStuff
   
   LabReduce = Trim$(Str$(Reduction))
   LabRot = Trim$(Str$(Rotation))
   
   ReDim DBGRA(0 To NColors - 1)
   ReDim XORACPY(32, 32, NumFrames)
   ReDim ANDACPY(32, 32, NumFrames)
   CopyMemory DBGRA(0), BGRA(0), NColors * 4
   CopyMemory XORACPY(1, 1, 1), XORA(1, 1, 1), 32 * 32 * NumFrames * 4
   CopyMemory ANDACPY(1, 1, 1), ANDA(1, 1, 1), 32 * 32 * NumFrames
   Left = frmEffectsLeft
   Top = frmEffectsTop
   
   If BPP = 24 Then
      optXFade.Enabled = True
      Frame2.Enabled = True
   Else
      optXFade.Enabled = False
      Frame2.Enabled = False
   End If
   
   HSRot.Value = Rotation
   HSReduce.Value = Reduction
   For i = 0 To 5
      SetParams i
   Next i
      
End Sub


Private Sub HSRot_Change()
   Rotation = HSRot.Value
   LabRot = Trim$(Str$(Rotation))
End Sub

Private Sub HSReduce_Change()
   Reduction = HSReduce.Value
   LabReduce = Trim$(Str$(Reduction))
End Sub


Private Sub LabQ1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtHelp(Index).Visible = True
End Sub

Private Sub LabQ1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtHelp(Index).Visible = False
End Sub

Private Sub SetParams(Index As Integer)
   aBlock = True
   Select Case Index
   Case 0: scrParam(Index).Value = 1
   Case 1: scrParam(Index).Value = 1
   Case 2: scrParam(Index).Value = 0
   Case 3: scrParam(Index).Value = 128
   Case 4: scrParam(Index).Value = 16
   Case 5: scrParam(Index).Value = 0
   End Select
   Param = scrParam(Index).Value
   zParam = 1 + Param / 100
   Select Case Index
   Case 0: LabParam(Index) = Trim$(Str$(zParam))
   Case 1: LabParam(Index) = Trim$(Str$(zParam))
   Case 2: LabParam(Index) = Trim$(Str$(Param))
   Case 3: LabParam(Index) = Trim$(Str$(Param))
   Case 4: LabParam(Index) = Trim$(Str$(Param))
   Case 5: LabParam(Index) = Trim$(Str$(Param))
   End Select
   aBlock = False
End Sub

Private Sub scrParam_Change(Index As Integer)
   If aBlock Then Exit Sub
   Param = scrParam(Index).Value
   zParam = 1 + Param / 100
   Select Case Index
   Case 0: LabParam(Index) = Trim$(Str$(zParam))
      DarkenBrighten
   Case 1: LabParam(Index) = Trim$(Str$(zParam))
      DO_PAL
   Case 2: LabParam(Index) = Trim$(Str$(Param))
      Smooth_Sharpen
   Case 3: LabParam(Index) = Trim$(Str$(Param))
      Black_White
   Case 4: LabParam(Index) = Trim$(Str$(Param))
      GreyDither
   Case 5
      LabParam(Index) = Trim$(Str$(Param))
      Diffuse
   End Select
End Sub

Private Sub scrParam_Scroll(Index As Integer)
   If aBlock Then Exit Sub
   Param = scrParam(Index).Value
   zParam = 1 + Param / 100
   LabParam(Index) = Trim$(Str$(zParam))
   Select Case Index
   Case 0: LabParam(Index) = Trim$(Str$(zParam))
      DarkenBrighten
   Case 1: LabParam(Index) = Trim$(Str$(zParam))
      DO_PAL
   Case 2: LabParam(Index) = Trim$(Str$(Param))
      Smooth_Sharpen
   Case 3: LabParam(Index) = Trim$(Str$(Param))
      Black_White
   Case 4: LabParam(Index) = Trim$(Str$(Param))
      GreyDither
   Case 5
      LabParam(Index) = Trim$(Str$(Param))
      Diffuse
   End Select
End Sub

Private Sub DarkenBrighten()
' zParam = .5 to 2
   CopyMemory XORA(1, 1, 1), XORACPY(1, 1, 1), 32 * 32 * NumFrames * 4
   For N = picNum To NumFrames
      For iy = 1 To 32
      For ix = 1 To 32
         Cul = XORA(ix, iy, N)
         If Cul <> 0 Then
            LngToRGB Cul
            RR = 1& * bred * zParam
            If RR > 255 Then RR = 255
            GG = 1& * bgreen * zParam
            If GG > 255 Then GG = 255
            BB = 1& * bblue * zParam
            If BB > 255 Then BB = 255
            XORA(ix, iy, N) = RGB(RR, GG, BB)
         End If
      Next ix
      Next iy
   Next N
   Form1.ShowAllIcons
End Sub

Private Sub Smooth_Sharpen()
'Param

Dim NP As Long
Dim cn As Long
Dim j As Long
Dim i As Long
   CopyMemory XORA(1, 1, 1), XORACPY(1, 1, 1), 32 * 32 * NumFrames * 4
   For N = picNum To NumFrames
      If Param > 0 Then
         'SHARPEN
'>0     -2  -2  -2  Param times
'       -2  +26 -2
'       -2  -2  -2
         For NP = 1 To Param
            For iy = 1 To 32
            For ix = 1 To 32
               SR = 0: SG = 0: SB = 0
               cn = 0
               For j = iy - 1 To iy + 1
               For i = ix - 1 To ix + 1
                  If i > 0 Then
                  If i < 33 Then
                  If j > 0 Then
                  If j < 33 Then
                     Cul = XORA(i, j, N)
                     If Cul <> 0 Then
                        LngToRGB Cul
                        If j <> iy Or i <> ix Then
                           SR = SR - 2& * bred
                           SG = SG - 2& * bgreen
                           SB = SB - 2& * bblue
                           cn = cn - 2
                        Else
                           SR = SR + 26& * bred
                           SG = SG + 26& * bgreen
                           SB = SB + 26& * bblue
                           cn = cn + 26
                        End If
                     End If
                  End If
                  End If
                  End If
                  End If
               Next i
               Next j
               If cn <> 0 Then
                  SR = SR \ cn
                  SG = SG \ cn
                  SB = SB \ cn
                  CheckRGB
                  XORA(ix, iy, N) = RGB(SR, SG, SB)
               End If
            Next ix
            Next iy
         Next NP
      ElseIf Param < 0 Then
      
      Select Case Param
      Case -1
      ' SMOOTH 1
'-1     0 1 0
'       1 4 1
'       0 1 0
         For iy = 1 To 32
         For ix = 1 To 32
            SR = 0: SG = 0: SB = 0
            cn = 0
            
            Cul = XORA(ix, iy, N)
            If Cul <> 0 Then
               LngToRGB Cul
               SR = SR + 4 * bred: SG = SG + 4 * bgreen: SB = SB + 4 * bblue
               cn = cn + 4
            End If
            
            
            j = iy - 1
            If j > 0 Then
               i = ix
               Cul = XORA(i, j, N)
               If Cul <> 0 Then
                  LngToRGB Cul
                  SR = SR + bred: SG = SG + bgreen: SB = SB + bblue
                  cn = cn + 1
               End If
            End If
            
            j = iy
            i = ix - 1
            If i > 0 Then
               Cul = XORA(i, j, N)
               If Cul <> 0 Then
                  LngToRGB Cul
                  SR = SR + bred: SG = SG + bgreen: SB = SB + bblue
                  cn = cn + 1
               End If
            End If
            
            j = iy
            i = ix + 1
            If i < 33 Then
               Cul = XORA(i, j, N)
               If Cul <> 0 Then
                  LngToRGB Cul
                  SR = SR + bred: SG = SG + bgreen: SB = SB + bblue
                  cn = cn + 1
               End If
            End If
            
            j = iy + 1
            If j < 33 Then
               i = ix
               Cul = XORA(i, j, N)
               If Cul <> 0 Then
                  LngToRGB Cul
                  SR = SR + bred: SG = SG + bgreen: SB = SB + bblue
                  cn = cn + 1
               End If
            End If
            
            If cn <> 0 Then
               SR = SR \ cn
               SG = SG \ cn
               SB = SB \ cn
               CheckRGB
               XORA(ix, iy, N) = RGB(SR, SG, SB)
            End If
         
         
         Next ix
         Next iy
      
      Case -2
      ' SMOOTH 2
'-2     1 1 1
'       1 0 1
'       1 1 1
         For iy = 1 To 32
         For ix = 1 To 32
            SR = 0: SG = 0: SB = 0
            cn = 0
            For j = iy - 1 To iy + 1
            For i = ix - 1 To ix + 1
               If i > 0 Then
               If i < 33 Then
               If j > 0 Then
               If j < 33 Then
                  Cul = XORA(i, j, N)
                  If Cul <> 0 Then
                     LngToRGB Cul
                     SR = SR + bred: SG = SG + bgreen: SB = SB + bblue
                     cn = cn + 1
                  End If
               End If
               End If
               End If
               End If
            Next i
            Next j
            If cn <> 0 Then
               SR = SR \ cn
               SG = SG \ cn
               SB = SB \ cn
               CheckRGB
               XORA(ix, iy, N) = RGB(SR, SG, SB)
            End If
         Next ix
         Next iy
      Case -3
      ' SMOOTH 3
'-3     1 1 1 1 1
'       1 1 1 1 1
'       1 1 0 1 1
'       1 1 1 1 1
'       1 1 1 1 1
         For iy = 1 To 32
         For ix = 1 To 32
            SR = 0: SG = 0: SB = 0
            cn = 0
            For j = iy - 2 To iy + 2
            For i = ix - 2 To ix + 2
               If i > 0 Then
               If i < 33 Then
               If j > 0 Then
               If j < 33 Then
                  Cul = XORA(i, j, N)
                  If Cul <> 0 Then
                     LngToRGB Cul
                     SR = SR + bred: SG = SG + bgreen: SB = SB + bblue
                     cn = cn + 1
                  End If
               End If
               End If
               End If
               End If
            Next i
            Next j
            If cn <> 0 Then
               SR = SR \ cn
               SG = SG \ cn
               SB = SB \ cn
               CheckRGB
               XORA(ix, iy, N) = RGB(SR, SG, SB)
            End If
         Next ix
         Next iy
      End Select
      End If
   Next N
   Form1.ShowAllIcons
End Sub

Private Sub CheckRGB()
   If SR < 0 Then SR = 0
   If SR > 255 Then SR = 255
   If SG < 0 Then SG = 0
   If SG > 255 Then SG = 255
   If SB < 0 Then SB = 0
   If SB > 255 Then SB = 255
End Sub

Private Sub Black_White()
' Param = 0 - 255
   CopyMemory XORA(1, 1, 1), XORACPY(1, 1, 1), 32 * 32 * NumFrames * 4
   For N = picNum To NumFrames
      For iy = 1 To 32
      For ix = 1 To 32
         Cul = XORACPY(ix, iy, N)
         LngToRGB Cul
         SG = (1& * bred + bgreen + bblue) \ 3
         If SG <= Param Then
            XORA(ix, iy, N) = 0
         Else
            XORA(ix, iy, N) = vbWhite
         End If
      Next ix
      Next iy
   Next N
   Form1.ShowAllIcons
End Sub

Private Sub GreyDither()
' Floyd-Steinberg B
' spreader
' 0 7 0
' 3 5 1 / 16
' Param = 16 - 48

Dim greysum As Long
Dim greycount As Long
Dim zDiv As Single
Dim zMul As Single
Dim zErr As Single
   CopyMemory XORA(1, 1, 1), XORACPY(1, 1, 1), 32 * 32 * NumFrames * 4
   ReDim Intens(-1 To 34, -1 To 34)
   greysum = 0
   greycount = 32 * 32
   For N = picNum To NumFrames
      For iy = 1 To 32
      For ix = 1 To 32
         Cul = XORACPY(ix, iy, N)
         LngToRGB Cul
         Intens(ix, iy) = (1& * bred + bgreen + bblue) \ 3
         greysum = greysum + Intens(ix, iy)
      Next ix
      Next iy
      greysum = greysum \ greycount
      zDiv = Param
      zMul = 1 / zDiv
      For iy = 2 To 31
      For ix = 2 To 31
         If Intens(ix, iy) > greysum Then
            XORA(ix, iy, N) = vbWhite
            zErr = (Intens(ix, iy) - 255) * zMul
         Else
            XORA(ix, iy, N) = 0
            zErr = Intens(ix, iy) * zMul
         End If
         ' Spread error
         Intens(ix - 1, iy + 1) = Intens(ix - 1, iy + 1) + 3 * zErr
         Intens(ix, iy + 1) = Intens(ix, iy + 1) + 5 * zErr
         Intens(ix + 1, iy + 1) = Intens(ix + 1, iy + 1) + zErr
         Intens(ix + 1, iy) = Intens(ix + 1, iy) + 7 * zErr
      Next ix
      Next iy
   Next N
   Erase Intens()
   Form1.ShowAllIcons
End Sub

Private Sub Diffuse()
Dim ix As Long, iy As Long
Dim i As Long, j As Long
Dim N As Long
   If Param = 0 Then
      For N = picNum To NumFrames
         CopyMemory XORA(1, 1, N), XORACPY(1, 1, N), 32 * 32 * 4
         CopyMemory ANDA(1, 1, N), ANDACPY(1, 1, N), 32 * 32
      Next N
   Else
      For N = picNum To NumFrames
         For iy = 1 To 32
         For ix = 1 To 32
            j = Rnd * Param - Param \ 2
            i = Rnd * Param - Param \ 2
            If ix + i < 1 Then i = 0
            If ix + i > 32 Then i = 0
            If iy + j < 1 Then j = 0
            If iy + j > 32 Then j = 0
            XORA(ix, iy, N) = XORA(ix + i, iy + j, N)
            ANDA(ix, iy, N) = ANDA(ix + i, iy + j, N)
         Next ix
         Next iy
      Next N
   End If
   Form1.ShowAllIcons
End Sub

Private Sub DO_PAL()
' zParam = .5 to 2
   CopyMemory BGRA(0), DBGRA(0), NColors * 4
   For k = 0 To NColors - 1
      With BGRA(k)
         Cul = (1& * BGRA(k).B * zParam)
         If Cul > 255 Then Cul = 255
         .B = Cul
         Cul = (1& * BGRA(k).G * zParam)
         If Cul > 255 Then Cul = 255
         .G = Cul
         Cul = (1& * BGRA(k).R * zParam)
         If Cul > 255 Then Cul = 255
         .R = Cul
      End With
   Next k
   Form1.ShowPalette
End Sub

Private Sub optCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optCopy.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      CopyLR Button
      Form1.ShowAllIcons
   End If
End Sub

Private Sub optCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optCopy.Value = False
   Me.SetFocus
End Sub

Private Sub optGO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' LC to left, RC to right
' Rotation deg,  Reduction %
' NumFrames, picNum (XORA etc picNum
   optGO.SetFocus
   If NumFrames < 2 Then Exit Sub
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      Gradates Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optGO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optGO.Value = False
   Me.SetFocus
End Sub

Private Sub optPepper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optPepper.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      Pepper Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optPepper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optPepper.Value = False
   Me.SetFocus
End Sub

Private Sub optReverse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optReverse.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      ReverseXORA_ANDA Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optReverse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optReverse.Value = False
   Me.SetFocus
End Sub

Private Sub optRoller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optRoller.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      RollerXORA_ANDA Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optRoller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optRoller.Value = False
   Me.SetFocus
End Sub

Private Sub optSwivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optSwivel.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      Swivel Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optSwivel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optSwivel.Value = False
   Me.SetFocus
End Sub

Private Sub optWave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optWave.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      Wave Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optWave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optWave.Value = False
   Me.SetFocus
End Sub

Private Sub optSwirl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optSwirl.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      Swirl Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optSwirl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optSwirl.Value = False
   Me.SetFocus
End Sub

Private Sub optXFade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optXFade.SetFocus
   If (Button = vbLeftButton And picNum) > 1 Or _
      (Button = vbRightButton And picNum < NumFrames) Then
      XFade Button
      Form1.ShowAllIcons
   End If
End Sub
Private Sub optXFade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   optXFade.Value = False
   Me.SetFocus
End Sub

Private Sub txtHelpStuff()
Dim A$
   A$ = ""
   A$ = A$ & " Using ONLY the selected frame," & vbCrLf
   A$ = A$ & " Rotates/Reduces it to the first" & vbCrLf
   A$ = A$ & " or last frame as Left- or" & vbCrLf
   A$ = A$ & " Right-click on the GO button." & vbCrLf
   txtHelp(1).Text = A$
   A$ = ""
   A$ = A$ & " Acts on all frames beyond the" & vbCrLf
   A$ = A$ & " selected one to the first or last" & vbCrLf
   A$ = A$ & " frame as Left- or Right-clicked." & vbCrLf
   A$ = A$ & " Apart from XFade which fades the" & vbCrLf
   A$ = A$ & " selected frame into the last" & vbCrLf
   A$ = A$ & " frame & only for 24 bpp images." & vbCrLf
   A$ = A$ & " Will also operate on a selection." & vbCrLf
   txtHelp(2).Text = A$
   A$ = ""
   A$ = A$ & " Acts on all frames including" & vbCrLf
   A$ = A$ & " that selected, to the Left or" & vbCrLf
   A$ = A$ & " to the Right." & vbCrLf
   txtHelp(3).Text = A$
   
   txtHelp(1).Left = 26 * zSF
   txtHelp(1).Top = 0 * zSF
   txtHelp(1).Visible = False
   txtHelp(2).Left = 24 * zSF
   txtHelp(2).Top = 67 * zSF
   txtHelp(2).Visible = False
   txtHelp(3).Left = 23 * zSF
   txtHelp(3).Top = 181 * zSF
   txtHelp(3).Visible = False
End Sub
