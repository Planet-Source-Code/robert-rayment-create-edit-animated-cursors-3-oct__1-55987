VERSION 5.00
Begin VB.Form frmCapture 
   Caption         =   "Capture"
   ClientHeight    =   7212
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   6672
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   556
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReCapture 
      Caption         =   "Re-Capture"
      Height          =   345
      Left            =   165
      TabIndex        =   10
      Top             =   120
      Width           =   1110
   End
   Begin VB.Frame fraCap 
      BackColor       =   &H80000016&
      Caption         =   "Capturing"
      Height          =   885
      Left            =   1965
      TabIndex        =   4
      Top             =   75
      Width           =   4365
      Begin VB.CommandButton cmdACImage 
         Caption         =   "Cancel"
         Height          =   360
         Index           =   1
         Left            =   1770
         TabIndex        =   8
         Top             =   315
         Width           =   735
      End
      Begin VB.CommandButton cmdACImage 
         Caption         =   "Accept"
         Height          =   360
         Index           =   0
         Left            =   930
         TabIndex        =   7
         Top             =   315
         Width           =   735
      End
      Begin VB.PictureBox pic1616 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4665
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   6
         Top             =   150
         Width           =   240
      End
      Begin VB.PictureBox picCAP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   210
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   5
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Move rectangle over image && Click to capture."
         Height          =   555
         Left            =   2640
         TabIndex        =   9
         Top             =   165
         Width           =   1605
      End
   End
   Begin VB.HScrollBar scrH 
      Height          =   240
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6795
      Width           =   5955
   End
   Begin VB.VScrollBar scrV 
      Height          =   5730
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1050
      Width           =   255
   End
   Begin VB.PictureBox picCON 
      Height          =   5730
      Left            =   405
      ScaleHeight     =   474
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   1005
      Width           =   5985
      Begin VB.PictureBox picIMAGE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   30
         ScaleHeight     =   241
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   234
         TabIndex        =   1
         Top             =   30
         Width           =   2805
         Begin VB.Shape shp3232 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            Height          =   480
            Left            =   795
            Top             =   885
            Width           =   480
         End
      End
   End
   Begin VB.Label LabiBPP 
      Caption         =   "LabiBPP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   11
      Top             =   585
      Width           =   975
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmCapture.frm

Option Explicit

'  Windows API to make form stay on top
' -----------------------------------------------------------

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Private Const hWndInsertAfter = -1
Private Const wFlags = &H40 Or &H2

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4


Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
ByVal X As Long, ByVal Y As Long, ByVal NWidth As Long, ByVal NHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'StretchBlt Dhdc,xd,yd,dw,dh,Shdc,xs,ys,sw,sh,vbSrcCopy

Private Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal NWidth As Long, ByVal NHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'BitBlt Dhdc,xd,yd,W,H,Shdc,xs,ys,vbsrccopy
Private iBPP As Integer
Private Ext$
Private BYTEGIF As Byte
Private FF As Long

Private CommonDialog1 As OSDialog

Private Sub Form_Load()
Dim res As Long
On Error GoTo STARTERR
   
   ' Size & Make frmZoom stay on top
   With frmCapture
      res = SetWindowPos(.hwnd, hWndInsertAfter, 15, 15, _
         6765 / STX, 7755 / STY, wFlags)
   End With
   picCAP.Width = 32 * STX
   picCAP.Height = 32 * STY
   shp3232.Width = 32
   shp3232.Height = 32
   LabiBPP = ""
   
   If LenB(ImageSpec$) <> 0 Then
      picIMAGE.Picture = LoadPicture(ImageSpec$)
      Ext$ = LCase$(Right$(ImageSpec$, 3))
      Select Case Ext$
      Case "bmp"
         FF = FreeFile
         Open ImageSpec$ For Binary Access Read As #FF
         Seek #FF, 29
         Get #FF, , iBPP
         Close
         If iBPP <> 1 And iBPP <> 4 And iBPP <> 8 And iBPP <> 24 Then
            LabiBPP = Str$(iBPP) & " bpp"
            MsgBox " BMP, not acceptable palette, make rough", vbCritical, "Capturing"
            Ext$ = "jpg"
            iBPP = 24
            Exit Sub
         End If
      Case "gif"
         iBPP = 8
      Case "jpg"
         iBPP = 24
      End Select
      LabiBPP = Str$(iBPP) & " bpp"
      FixScrollbars picCON, picIMAGE, scrH, scrV
   End If
Exit Sub
'=========
STARTERR:
Close
On Error GoTo 0
MsgBox "Capturing image error, reading file", vbCritical, "Capturing"
picIMAGE.Picture = LoadPicture
End Sub

Private Sub cmdReCapture_Click()
' Re-Capturing
Dim p As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long

   Set CommonDialog1 = New OSDialog
   Title$ = "Extract from image"
   Filt$ = "Open Pic |*.bmp;*.gif;*.jpg"
   If ImageSpec$ = "" Then
      InDir$ = PathSpec$
   Else
      p = InStrRev(ImageSpec$, "\")
      InDir$ = Left$(ImageSpec$, p)
   End If
   CommonDialog1.ShowOpen ImageSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   Set CommonDialog1 = Nothing
   
   If LenB(ImageSpec$) <> 0 Then
      picIMAGE.Picture = LoadPicture(ImageSpec$)
      Ext$ = LCase$(Right$(ImageSpec$, 3))
      Select Case Ext$
      Case "bmp"
         FF = FreeFile
         Open ImageSpec$ For Binary Access Read As #FF
         Seek #FF, 29
         Get #FF, , iBPP
         Close
         If iBPP <> 1 And iBPP <> 4 And iBPP <> 8 And iBPP <> 24 Then
            LabiBPP = Str$(iBPP) & " bpp"
            MsgBox " BMP, not acceptable palette", vbCritical, "Capturing"
            Exit Sub
         End If
      Case "gif"
         iBPP = 8
      Case "jpg"
         iBPP = 24
      End Select
      LabiBPP = Str$(iBPP) & " bpp"
   End If
   FixScrollbars picCON, picIMAGE, scrH, scrV
End Sub

Private Sub cmdACImage_Click(Index As Integer)
Dim ix As Long
Dim iy As Long
Dim Cul As Long
Dim k As Long
   On Error GoTo CAP_ERR
   Select Case Index
   Case 0   ' Accept
      
      Select Case Ext$
      Case "bmp"   ' BMP 24,8,4,1
         FF = FreeFile
         Open ImageSpec$ For Binary Access Read As #FF
         Seek #FF, 29
         Get #FF, , iBPP
         BPP = iBPP
         Select Case BPP
         Case 1: NColors = 2
         Case 4: NColors = 16
         Case 8: NColors = 256
         Case 24
            ' Get rough palette
            ReDim BGRA(0 To 255)
            k = 0
            For iy = 0 To 15
            For ix = 0 To 15
               Cul = pic1616.Point(ix, iy)
               LngToRGB Cul
               BGRA(k).B = bblue
               BGRA(k).G = bgreen
               BGRA(k).R = bred
               k = k + 1
            Next ix
            Next iy
         Case Else   ' Error
            GoTo CAP_ERR
         End Select
         
         Select Case BPP
         Case 1, 4, 8
            ReDim BGRA(0 To NColors - 1)
            Seek #FF, 55
            For k = 0 To NColors - 1
               Get #FF, , BGRA(k).B
               Get #FF, , BGRA(k).G
               Get #FF, , BGRA(k).R
               Get #FF, , BGRA(k).AL
            Next k
         End Select
         Close #FF
      
      Case "gif"   ' GIF 8
         FF = FreeFile
         Open ImageSpec$ For Binary Access Read As #FF
         Seek #FF, 11
         Get #FF, , BYTEGIF
         NColors = 2 ^ ((BYTEGIF And &H7) + 1)
         iBPP = ((BYTEGIF And &H70) / 2 ^ 4) + 1
         ' BUT
         ReDim BGRA(0 To 255)    ' Always make it 8 bpp
         BPP = 8
         Seek #FF, 14
         ' NB Wrong if no Global palette
         For k = 0 To NColors - 1
            Get #FF, , BGRA(k).R
            Get #FF, , BGRA(k).G
            Get #FF, , BGRA(k).B
         Next k
         ' SO
         NColors = 256
         Close #FF
      
      Case "jpg"   ' JPG 24
         BPP = 24
         ' Get rough palette
         ReDim BGRA(0 To 255)
         k = 0
         For iy = 0 To 15
         For ix = 0 To 15
            Cul = pic1616.Point(ix, iy)
            LngToRGB Cul
            BGRA(k).B = bblue
            BGRA(k).G = bgreen
            BGRA(k).R = bred
            k = k + 1
         Next ix
         Next iy
      End Select
      
      ' Transfer image to XORA(1,1,picNum)
      For iy = 0 To 31
      For ix = 0 To 31
         Cul = picCAP.Point(ix, iy)
         XORA(ix + 1, 33 - (iy + 1), picNum) = Cul
         ANDA(ix + 1, 33 - (iy + 1), picNum) = 0
      Next ix
      Next iy
      
      With frmDetails
         .fraINFO.Caption = "INFO"
         Select Case BPP
         Case 1: .optBPP(0).Value = True
         Case 4: .optBPP(1).Value = True
         Case 8: .optBPP(2).Value = True
         Case 24: .optBPP(3).Value = True
         End Select
      End With
      Form1.LabFileName = " Captured "
      Form1.ShowAllIcons
      Form1.ShowPalette
      DoEvents

   Case 1   ' Cancel
   End Select
   
   picIMAGE.Picture = LoadPicture
   picIMAGE.Width = 32
   picIMAGE.Height = 32
   
   Unload Me
   Exit Sub
'==========
CAP_ERR:
   Close
   On Error GoTo 0
   MsgBox "Capture Accept Error  ", vbCritical, "Capturing"
   Unload Me
End Sub

Private Sub Form_Resize()
   If Me.Width >= 6765 And Me.Height >= 7755 Then
      picCON.Width = Me.Width / STX - 20 - picCON.Left
      picCON.Height = Me.Height / STY - 70 - picCON.Top
      FixScrollbars picCON, picIMAGE, scrH, scrV
   Else
      Me.Width = 6765
      Me.Height = 7755
      FixScrollbars picCON, picIMAGE, scrH, scrV
   End If
End Sub

Private Sub picIMAGE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With shp3232
      .Left = X
      .Top = Y
   End With
End Sub

Private Sub picIMAGE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Extract 32x32 image
   picCAP.Picture = LoadPicture
   BitBlt picCAP.hDC, 0, 0, 32, 32, picIMAGE.hDC, X, Y, vbSrcCopy
   picCAP.Refresh
   
   ' Shrink to 16x16 for rough palette
   pic1616.Picture = LoadPicture
   SetStretchBltMode pic1616.hDC, HALFTONE
   StretchBlt pic1616.hDC, 0, 0, 16, 16, picCAP.hDC, 0, 0, 32, 32, vbSrcCopy
   pic1616.Refresh
End Sub

'Private Sub fraCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''Public Xfra As Single
''Public Yfra As Single
'   Xfra = X
'   Yfra = Y
'End Sub
'
'Private Sub fraCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   fraMOVER frmCapture, fraCap, Button, X, Y
'End Sub

Private Sub scrH_Change()
   picIMAGE.Left = -scrH.Value
End Sub

Private Sub scrH_Scroll()
   picIMAGE.Left = -scrH.Value
End Sub

Private Sub scrV_Change()
   picIMAGE.Top = -scrV.Value
End Sub

Private Sub scrV_Scroll()
   picIMAGE.Top = -scrV.Value
End Sub

