VERSION 5.00
Begin VB.Form frmDetails 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "Details"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   4500
   ClientWidth     =   6165
   ControlBox      =   0   'False
   Icon            =   "frmDetails.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      ScaleHeight     =   345
      ScaleWidth      =   6000
      TabIndex        =   42
      Top             =   60
      Width           =   6030
      Begin VB.OptionButton optMinimize 
         BackColor       =   &H80000003&
         Caption         =   "&M"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   5610
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   " Minimize "
         Top             =   15
         Width           =   375
      End
      Begin VB.Shape Shape1 
         Height          =   30
         Left            =   -15
         Top             =   345
         Width           =   4665
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000002&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   43
         Top             =   15
         Width           =   930
      End
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
      Height          =   1260
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "frmDetails.frx":000C
      Top             =   870
      Width           =   1260
   End
   Begin VB.Frame fraRates 
      BackColor       =   &H80000013&
      Caption         =   "Rate Table"
      Height          =   1665
      Left            =   615
      TabIndex        =   30
      Top             =   3975
      Width           =   4950
      Begin VB.VScrollBar VRates 
         Height          =   300
         Index           =   0
         Left            =   1155
         Max             =   1
         Min             =   600
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1245
         Value           =   1
         Width           =   315
      End
      Begin VB.VScrollBar VRates 
         Height          =   300
         Index           =   1
         Left            =   1485
         Max             =   1
         Min             =   600
         SmallChange     =   10
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1245
         Value           =   1
         Width           =   315
      End
      Begin VB.VScrollBar VRates 
         Height          =   300
         Index           =   2
         Left            =   3570
         Max             =   1
         Min             =   600
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1245
         Value           =   1
         Width           =   315
      End
      Begin VB.VScrollBar VRates 
         Height          =   300
         Index           =   3
         Left            =   3900
         Max             =   1
         Min             =   600
         SmallChange     =   10
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1245
         Value           =   1
         Width           =   315
      End
      Begin VB.Label LRates 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "64"
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
         Left            =   75
         TabIndex        =   40
         Top             =   255
         Width           =   285
      End
      Begin VB.Label LabTheRate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   39
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label LabRNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "64"
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
         Height          =   210
         Left            =   135
         TabIndex        =   38
         Top             =   1290
         Width           =   300
      End
      Begin VB.Label LabTheRate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3030
         TabIndex        =   37
         Top             =   1245
         Width           =   510
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000013&
         Caption         =   "Single"
         Height          =   195
         Index           =   0
         Left            =   1860
         TabIndex        =   36
         Top             =   1305
         Width           =   555
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000013&
         Caption         =   "All"
         Height          =   195
         Index           =   1
         Left            =   4320
         TabIndex        =   35
         Top             =   1305
         Width           =   255
      End
   End
   Begin VB.Frame fraINFO 
      BackColor       =   &H80000013&
      Caption         =   "NEW"
      Height          =   3495
      Left            =   105
      TabIndex        =   0
      Top             =   465
      Width           =   5925
      Begin VB.TextBox txtTitle 
         Height          =   330
         Left            =   150
         TabIndex        =   24
         Text            =   "Untitled"
         Top             =   435
         Width           =   2670
      End
      Begin VB.TextBox txtAuthor 
         Height          =   345
         Left            =   150
         TabIndex        =   23
         Text            =   "Noname"
         Top             =   1035
         Width           =   2670
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Color"
         Height          =   1530
         Left            =   2910
         TabIndex        =   17
         Top             =   240
         Width           =   2880
         Begin VB.OptionButton optBPP 
            BackColor       =   &H80000013&
            Caption         =   "1 bpp Black && white"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   21
            Top             =   360
            Width           =   2280
         End
         Begin VB.OptionButton optBPP 
            BackColor       =   &H80000013&
            Caption         =   "4 bpp 16 colors"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   20
            Top             =   615
            Width           =   2340
         End
         Begin VB.OptionButton optBPP 
            BackColor       =   &H80000013&
            Caption         =   "8 bpp 256 colors"
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   19
            Top             =   885
            Width           =   2175
         End
         Begin VB.OptionButton optBPP 
            BackColor       =   &H80000013&
            Caption         =   "24 bpp High color"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   18
            Top             =   1155
            Width           =   2325
         End
         Begin VB.Label LabQ2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "?"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2580
            TabIndex        =   22
            Top             =   180
            Width           =   165
         End
      End
      Begin VB.VScrollBar vscrNumFrames 
         Height          =   330
         Left            =   2235
         Max             =   1
         Min             =   64
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1590
         Value           =   2
         Width           =   330
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000013&
         Height          =   990
         Left            =   135
         TabIndex        =   6
         Top             =   2295
         Width           =   2700
         Begin VB.VScrollBar vscrHotY 
            Height          =   255
            Left            =   1305
            Max             =   31
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   525
            Value           =   31
            Width           =   300
         End
         Begin VB.PictureBox picHotXY 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   1950
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   8
            Top             =   285
            Width           =   510
         End
         Begin VB.HScrollBar hscrHotX 
            Height          =   210
            Left            =   1290
            Max             =   31
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000013&
            Caption         =   "HotX"
            Height          =   270
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   225
            Width           =   480
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000013&
            Caption         =   "HotY"
            Height          =   270
            Index           =   1
            Left            =   150
            TabIndex        =   14
            Top             =   540
            Width           =   480
         End
         Begin VB.Line Line1 
            X1              =   1845
            X2              =   2115
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line2 
            X1              =   1830
            X2              =   1830
            Y1              =   180
            Y2              =   435
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000013&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   0
            Left            =   2190
            TabIndex        =   13
            Top             =   120
            Width           =   180
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000013&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   1
            Left            =   1740
            TabIndex        =   12
            Top             =   495
            Width           =   180
         End
         Begin VB.Label LabHotX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   660
            TabIndex        =   11
            Top             =   225
            Width           =   480
         End
         Begin VB.Label LabHotY 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   660
            TabIndex        =   10
            Top             =   525
            Width           =   480
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000013&
         Height          =   1560
         Left            =   2895
         TabIndex        =   1
         Top             =   1740
         Width           =   2880
         Begin VB.CommandButton cmdRateTable 
            BackColor       =   &H80000013&
            Caption         =   "Rate Table"
            Height          =   270
            Left            =   975
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000013&
            Caption         =   "Delay in Jiffies (1/60 th sec) between frames.  Default Jiffy = 10.           Total time = Sum(Rates) / 60.  "
            Height          =   615
            Left            =   90
            TabIndex        =   5
            Top             =   585
            Width           =   2655
         End
         Begin VB.Label LabTotTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1065
            TabIndex        =   4
            Top             =   1230
            Width           =   660
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000013&
            Caption         =   "sec"
            Height          =   180
            Left            =   1830
            TabIndex        =   3
            Top             =   1260
            Width           =   330
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         Caption         =   "Title"
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         Caption         =   "Author"
         Height          =   210
         Left            =   195
         TabIndex        =   28
         Top             =   825
         Width           =   690
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         Caption         =   "Number of frames         (Max = 64)"
         Height          =   375
         Left            =   165
         TabIndex        =   27
         Top             =   1530
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         Caption         =   "Size: 32 x 32 fixed"
         Height          =   210
         Left            =   165
         TabIndex        =   26
         Top             =   2085
         Width           =   1425
      End
      Begin VB.Label LabNumFrames 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 2"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1575
         TabIndex        =   25
         Top             =   1650
         Width           =   480
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   408
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Height          =   5775
      Left            =   15
      Top             =   15
      Width           =   6120
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDetails.frm

Option Explicit
Option Base 1

Public xfrm As Single
Public yfrm As Single

Private Sub Form_Load()
   aPlay = False
   If AniTest Then
      RestoreOldCursor
      AniTest = False
   End If
   
   xfrm = Left
   yfrm = Top
End Sub

Private Sub cmdRateTable_Click()
   fraRates.Visible = Not fraRates.Visible
   If fraRates.Visible Then
      Height = 387 * STY
   Else
      Height = 268 * STY
   End If
End Sub


Public Sub LRates_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim N As Long
   Form1.LabFN(picNumHighlighted).BackColor = vbWhite
   LRates(picNumHighlighted).BackColor = vbWhite
   Form1.LabFN(Index).BackColor = vbYellow
   LRates(Index).BackColor = vbYellow
   picNumHighlighted = Index
   LRates(Index).Refresh
   LabRNum = Str$(Index)
   LabTheRate(0).Caption = LRates(Index).Caption
   If LabTheRate(0).Caption = "" Then
      N = 1
   Else
      N = Val(LabTheRate(0).Caption)
      If N = 0 Then N = 1
   End If
   picNum = Index
   avscr = False
   VRates(0).Value = N
   VRates(1).Value = N
   VRates(2).Value = N
   VRates(3).Value = N
   avscr = True
   Form1.CalcDuration
End Sub


Private Sub picMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   xfrm = X
   yfrm = Y
End Sub

Private Sub picMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
   Me.Left = Me.Left + X - xfrm
   Me.Top = Me.Top + Y - yfrm
   End If
End Sub

Private Sub VRates_Change(Index As Integer)
Dim R As Long
Dim k As Long
   If avscr = False Then Exit Sub
   R = VRates(Index).Value
   Select Case Index
   Case 0, 1
      LabTheRate(0).Caption = Str$(R)
      LRates(picNum) = Trim$(Str$(R))
      RateValues(picNum) = R
   Case 2, 3  ' All RateValues
      LabTheRate(1).Caption = Str$(R)
      For k = 1 To NumFrames
         LRates(k) = Trim$(Str$(R))
         RateValues(k) = R
      Next k
   End Select

   Select Case Index
   Case 0: VRates(1).Value = VRates(0).Value
   Case 1: VRates(0).Value = VRates(1).Value
   Case 2: VRates(3).Value = VRates(2).Value
   Case 3: VRates(2).Value = VRates(3).Value
   End Select
   Form1.CalcDuration
End Sub

Public Sub optBPP_Click(Index As Integer)
   If aNew Then
      Select Case Index
      Case 0: BPP = 1
      Case 1: BPP = 4
      Case 2: BPP = 8
      Case 3: BPP = 24
      End Select
      Form1.DefaultPalette
   Else  ' Reset BPP if not NEW
      Select Case BPP
      Case 1: optBPP(0).Value = True
      Case 4: optBPP(1).Value = True
      Case 8: optBPP(2).Value = True
      Case 24: optBPP(3).Value = True
      End Select
   End If
   Form1.LabBPP = Str$(BPP) & " bpp"
End Sub

Private Sub hscrHotX_Change()
   If aMouseDown Then Exit Sub
   HotX = hscrHotX.Value
   LabHotX = Str$(HotX)
   ShowHotXY vbRed
End Sub

Private Sub vscrHotY_Change()
   If aMouseDown Then Exit Sub
   HotY = vscrHotY.Value
   LabHotY = Str$(HotY)
   ShowHotXY vbRed
End Sub

Public Sub ShowHotXY(Cul As Long)
   picHotXY.Cls
   picHotXY.PSet (HotX, HotY - 1), Cul
   picHotXY.PSet (HotX - 1, HotY), Cul
   picHotXY.PSet (HotX, HotY), Cul
   picHotXY.PSet (HotX + 1, HotY), Cul
   picHotXY.PSet (HotX, HotY + 1), vbRed
   picHotXY.Refresh
End Sub

Private Sub vscrNumFrames_Change()
Dim N As Long
   If aMouseDown Then Exit Sub
   If avscr = False Then Exit Sub
   
   N = vscrNumFrames.Value
   
   If N > 0 Then
      If N < NumFrames Then
         picNum = NumFrames
         ClearXORA_ANDA
      End If
      ReDim Preserve XORA(32, 32, N)
      ReDim Preserve ANDA(32, 32, N)
      If N > NumFrames Then
         picNum = N
         ClearXORA_ANDA
      End If
      picNum = 1
      NumFrames = N
      LabNumFrames = Str$(NumFrames)
      Form1.LabFN(picNumHighlighted).BackColor = vbWhite
      frmDetails.LRates(picNumHighlighted).BackColor = vbWhite
      picNumHighlighted = 1         ' For LRates() & LabFN()
      Form1.LabFN(1).BackColor = vbYellow
      frmDetails.LRates(1).BackColor = vbYellow
      Form1.ShowAllIcons   ' Does aPic_MouseUp CInt(picNum), 0, 0, 0, 0
      Form1.Fill_fraINFO
   End If
   VRates_Change 2
End Sub

Private Sub LabQ2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtHelp.Visible = True
End Sub

Private Sub LabQ2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtHelp.Visible = False
End Sub

Private Sub optMinimize_Click()
   xfrm = Left
   yfrm = Top
   optMinimize.Value = Not optMinimize.Value
   WindowState = vbMinimized
End Sub

