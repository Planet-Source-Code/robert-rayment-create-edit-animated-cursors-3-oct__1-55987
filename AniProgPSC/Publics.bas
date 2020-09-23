Attribute VB_Name = "Publics"
' Publics.bas

Option Explicit
Option Base 1

Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function BitBlt Lib "gdi32" _
   (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
   ByVal NWidth As Long, ByVal NHeight As Long, _
   ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)


Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, Source As Any, ByVal Length As Long)

'Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" _
'   (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
'Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
'   (Destination As Any, ByVal Length As Long)

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" _
   (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long

'Private Const EWX_LOGOFF = 0
'Private Const EWX_SHUTDOWN = 1
'Private Const EWX_REBOOT = 2
'Private Const EWX_FORCE = 4
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
'Private Const DM_BITSPERPEL = &H40000
'Private Const DM_PELSWIDTH = &H80000
'Private Const DM_PELSHEIGHT = &H100000
'Private Const CDS_UPDATEREGISTRY = &H1
'Private Const CDS_TEST = &H4
'Private Const DISP_CHANGE_SUCCESSFUL = 0
'Private Const DISP_CHANGE_RESTART = 1
Public Const ENUM_CURRENT_SETTINGS = -1
'Private Const ENUM_REGISTRY_SETTINGS = -2

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public ScreenBits As Long

''''' Used :- ''''''''''''''''''''''''''''''''''''''''''''''
'    "String" to bytes:
'    CopyMemory ByteArr(SIndex), ByVal AString$, Len
'    Bytes to "string":
'    AString$ = Space$(Len)
'    CopyMemory ByVal AString$, ByteArr(SIndex), Len
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Files
Public PathSpec$  ' App path
Public FileSpec$  ' Ani spec
Public PalSpec$   ' JASC_PAL spec
Public ImageSpec$  ' Capture
Public aniFSpec$         ' Temp filespec for Testing anicursor
Public ImageType As Long
Public STX As Long, STY As Long     ' Screen Twips/pixel
Public LCul As Long, RCul As Long   ' Left & Right colors
Public picNum As Long               ' Picture number (1 to NumFrames)
Public picNumHighlighted As Long
Public GridCul As Long              ' Grid & display colors
Public aMouseDown As Boolean        ' Flag MouseDown
Public aNew As Boolean              ' Flag if New ani
Public Tool As Long                 ' Drawing Tools (0-
Public zSF As Single   ' Scale factor for Text help locs

'Public Const GenBackcul = &HE0E0E0
Public avscr As Boolean  ' Block/Unblock vscrNumFrames_Change()

Public aEffects As Boolean
Public Rotation As Long
Public Reduction As Long

Public aPlay As Boolean

' Frame mover
'Public Xfra As Single
'Public Yfra As Single

' frm locations
Public frmHelpLeft As Long   ' frmHelp.Left
Public frmHelpTop As Long   ' frmHelp.Top
Public frmEffectsLeft As Long   ' frmEffects.Left
Public frmEffectsTop As Long   ' frmEffects.Top

Public Const pi# = 3.142159265
Public Const d2r# = pi# / 180

Private MPal() As Byte '0 To 2, 0 To 255) As Byte


Public Sub LngToRGB(LCul As Long)
'Public bred As Byte, bgreen As Byte, bblue As Byte
'Convert Long Colors() to RGB components
   bred = (LCul And &HFF&)
   bgreen = (LCul And &HFF00&) / &H100&
   bblue = (LCul And &HFF0000) / &H10000
End Sub


Public Sub READ_JASC_PAL(FSpec$)
' Public BPP
' Public BGRA(0-255) .B .G .R
' Public NColors
Dim fnum As Long
Dim A$
Dim k As Long
Dim svNColors As Long
   svNColors = NColors
   fnum = FreeFile
   On Error GoTo PalFileError
   ' Read first line
   Open FSpec$ For Input As #fnum
   Line Input #fnum, A$
   Close #fnum
   
   If InStr(1, A$, "JASC") <> 0 Then  'JASC-PAL MAP file
      Open FSpec$ For Input As #fnum
                   'JASC-PAL
                   '0100
                   'Skip 3 lines  '256
      Line Input #fnum, A$
      Line Input #fnum, A$
      Line Input #fnum, A$
      Select Case BPP
      Case 1
         NColors = 2
         ReDim MPal(0 To 2, 0 To NColors - 1)
         For k = 0 To NColors - 1
            If EOF(1) Then Exit For
            Input #fnum, MPal(0, k), MPal(1, k), MPal(2, k)
         Next k
      Case 4
         NColors = 16
         ReDim MPal(0 To 2, 0 To NColors - 1)
         For k = 0 To NColors - 1
            If EOF(1) Then Exit For
            Input #fnum, MPal(0, k), MPal(1, k), MPal(2, k)
         Next k
         Close #fnum
      Case 8, 24
         NColors = 256
         ReDim MPal(0 To 2, 0 To NColors - 1)
         For k = 0 To NColors - 1
            If EOF(1) Then Exit For
            Input #fnum, MPal(0, k), MPal(1, k), MPal(2, k)
         Next k
         Close #fnum
      End Select
      
      
      ReDim BGRA(0 To NColors - 1)
      Select Case BPP
      Case 1, 4, 8, 24
         For k = 0 To NColors - 1
            BGRA(k).B = MPal(2, k)
            BGRA(k).G = MPal(1, k)
            BGRA(k).R = MPal(0, k)
         Next k
      End Select
   Else
      FSpec$ = ""
      Close
   End If
   
   Erase MPal()
   Exit Sub
'===========
PalFileError:
Close
Erase MPal()
NColors = svNColors
FSpec$ = ""
End Sub

Public Sub SAVE_JASC_PAL(FSpec$)
Dim fnum As Long
Dim k As Long
   fnum = FreeFile
   Open FSpec$ For Output As #fnum
      Print #fnum, "JASC-PAL"
      Print #fnum, "0100"
      k = UBound(BGRA())
      Print #fnum, Trim$(Str$(k + 1)) ' eg 256
      For k = 0 To k - 1
         Print #fnum, LTrim$(Str$(BGRA(k).R));
         Print #fnum, Str$(BGRA(k).G);
         Print #fnum, Str$(BGRA(k).B)
      Next k
   Close
End Sub

'Public Function GetLongFromBNum(bNUM() As Byte) As Long
'   ' Make Long from 4 bytes of Array BNum(1,2,3,4)
'   Dim LoWord As Long, HiWord As Long
'
'   If bNUM(2) And &H80 Then
'      LoWord = ((bNUM(2) * &H100&) Or bNUM(1)) Or &HFFFF0000
'   Else
'      LoWord = (bNUM(2) * &H100) Or bNUM(1)
'   End If
'
'   If bNUM(4) And &H80 Then
'      HiWord = ((bNUM(4) * &H100&) Or bNUM(3)) Or &HFFFF0000
'   Else
'      HiWord = (bNUM(4) * &H100) Or bNUM(3)
'   End If
'   ' by Karl E. Peterson see VBSpeed (www.xbeat.net\vbspeed\index.htm)
'   GetLongFromBNum = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
'End Function
'
'Public Sub SplitLngNum2BNum(LNUM As Long, bNUM() As Byte)
'Dim A$
'' Public BNum() As Byte
'' For Undo editting
'' Place bytes in BNum(1-4) in DISPLAY ORDER!
'A$ = Hex$(LNUM)
'If Len(A$) < 8 Then A$ = String(8 - Len(A$), "0") & A$
'bNUM(1) = Val("&H" + Right$(A$, 2))
'bNUM(2) = Val("&H" + Mid$(A$, 5, 2))
'bNUM(3) = Val("&H" + Mid$(A$, 3, 2))
'bNUM(4) = Val("&H" + Left$(A$, 2))
'End Sub

Function zATan2(ByVal zy As Single, ByVal zx As Single)
' Public pi# = Const
' 0 degrees to right
' Find angle Atan from -pi# to +pi#

If zx <> 0 Then
   zATan2 = Atn(zy / zx)
   If (zx < 0) Then
      If (zy < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
   End If
Else  ' zx=0
   If Abs(zy) > Abs(zx) Then   'Must be an overflow
      If zy > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
   Else
      zATan2 = 0   'Must be an underflow
   End If
End If
End Function

'Public Sub fraMOVER(frm As Form, fra As Frame, Button As Integer, X As Single, Y As Single)
''Public Xfra As Single
''Public Yfra As Single
'Dim fraLeft As Long
'Dim fraTop As Long
'
'   If Button = vbLeftButton Then
'
'      fraLeft = fra.Left + (X - Xfra) \ STX
'      If fraLeft < 0 Then fraLeft = 0
'      If fraLeft + fra.Width > frm.Width \ STX + fra.Width \ 2 Then
'         fraLeft = frm.Width \ STX - fra.Width \ 2
'      End If
'      fra.Left = fraLeft
'
'      fraTop = fra.Top + (Y - Yfra) \ STY
'      If fraTop < 8 Then fraTop = 8
'      If fraTop + fra.Height > frm.Height \ STY + fra.Height \ 2 Then
'         fraTop = frm.Height \ STY - fra.Height \ 2
'      End If
'      fra.Top = fraTop
'
'   End If
'End Sub
'#### END GENERAL FRAME MOVER ####################################

'#### POSITION SCROLL BARS AS picbox picP & piccontainer picC ##########

Public Sub FixScrollbars(picCr As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picCr = Picture Container
   ' picP  = Picture
   HS.Max = picP.Width - picCr.Width + 12   ' +4 to allow for border
   VS.Max = picP.Height - picCr.Height + 12 ' +4 to allow for border
   HS.LargeChange = picCr.Width \ 10
   HS.SmallChange = 1
   VS.LargeChange = picCr.Height \ 10
   VS.SmallChange = 1
   HS.Top = picCr.Top + picCr.Height + 1
   HS.Left = picCr.Left
   HS.Width = picCr.Width
   If picP.Width < picCr.Width Then
      HS.Visible = False
      HS.Enabled = False
   Else
      HS.Visible = True
      HS.Enabled = True
   End If
   VS.Top = picCr.Top
   VS.Left = picCr.Left - VS.Width - 1
   VS.Height = picCr.Height
   If picP.Height < picCr.Height Then
      VS.Visible = False
      VS.Enabled = False
   Else
      VS.Visible = True
      VS.Enabled = True
   End If
End Sub
'#### END POSITION SCROLL BARS AS picbox picP & piccontainer picC ##########

Public Sub CenteredPal()
' Method from Stefan Casier Paint256
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim zR As Single, zG As Single, zB As Single
Dim k As Long, k2 As Long, k3 As Long
Dim j As Long
ReDim BGRA(0 To 255)
Dim PCul(0 To 2, 0 To 15) As Byte
   PCul(0, 0) = 0:    PCul(1, 0) = 0:    PCul(2, 0) = 254
   PCul(0, 1) = 86:   PCul(1, 1) = 254:  PCul(2, 1) = 86
   PCul(0, 2) = 254:  PCul(1, 2) = 168:  PCul(2, 2) = 86
   PCul(0, 3) = 92:   PCul(1, 3) = 0:    PCul(2, 3) = 0
   PCul(0, 4) = 254:  PCul(1, 4) = 254:  PCul(2, 4) = 0
   PCul(0, 5) = 0:    PCul(1, 5) = 112:  PCul(2, 5) = 142
   PCul(0, 6) = 254:  PCul(1, 6) = 254:  PCul(2, 6) = 254
   PCul(0, 7) = 174:  PCul(1, 7) = 174:  PCul(2, 7) = 174
   PCul(0, 8) = 138:  PCul(1, 8) = 110:  PCul(2, 8) = 233
   PCul(0, 9) = 0:    PCul(1, 9) = 101:  PCul(2, 9) = 0
   PCul(0, 10) = 0:   PCul(1, 10) = 254: PCul(2, 10) = 254
   PCul(0, 11) = 254: PCul(1, 11) = 0:   PCul(2, 11) = 0
   PCul(0, 12) = 254: PCul(1, 12) = 254: PCul(2, 12) = 0
   PCul(0, 13) = 178: PCul(1, 13) = 0:   PCul(2, 13) = 178
   PCul(0, 14) = 254: PCul(1, 14) = 254: PCul(2, 14) = 254
   PCul(0, 15) = 90:  PCul(1, 15) = 90:  PCul(2, 15) = 90
   k3 = 0
   For k = 0 To 15
      k2 = k + 1
      If k = 15 Then k2 = 0
      zR = (1& * PCul(0, k) - PCul(0, k2)) / 16
      zG = (1& * PCul(1, k) - PCul(1, k2)) / 16
      zB = (1& * PCul(2, k) - PCul(2, k2)) / 16
      R2 = PCul(0, k)
      G2 = PCul(1, k)
      B2 = PCul(2, k)
      For j = 0 To 14
         R1 = R2 - zR
         G1 = G2 - zG
         B1 = B2 - zB
         If R1 < 0 Then R1 = 255
         If R1 > 255 Then R1 = 0
         R2 = R1
         If G1 < 0 Then G1 = 255
         If G1 > 255 Then G1 = 0
         G2 = G1
         If B1 < 0 Then B1 = 255
         If B1 > 255 Then B1 = 0
         B2 = B1
         BGRA(15 + k3).B = B1
         BGRA(15 + k3).G = G1
         BGRA(15 + k3).R = R1
         k3 = k3 + 1
         If k3 > 255 Then Exit For
      Next j
      If k3 > 255 Then Exit For
   Next k
   ' Adjustments
   For k = 0 To 15
      LngToRGB QBColor(k)
      BGRA(k).B = bred
      BGRA(k).G = bgreen
      BGRA(k).R = bblue
   Next k

   BGRA(1).B = 255
   BGRA(1).G = 255
   BGRA(1).R = 255
   BGRA(255).B = 255
   BGRA(255).G = 0
   BGRA(255).R = 0
   
   BGRA(15).B = 200
   BGRA(15).G = 200
   BGRA(15).R = 255
   
   NColors = 256
   
   
   ' SAVE CenBand.Pal
'   Open "CenBand.pal" For Output As #1
'   Print #1, "JASC-PAL"
'   Print #1, "0100"
'   Print #1, "256"
'   For k = 0 To 255
'      Print #1, LTrim$(Str$(BGRA(k).R));
'      Print #1, Str$(BGRA(k).G);
'      Print #1, Str$(BGRA(k).B)
'   Next k
'   Close
End Sub


Public Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub






