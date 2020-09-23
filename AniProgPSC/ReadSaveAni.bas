Attribute VB_Name = "ReadSaveAni"
' ReadSaveAni.bas
' by Robert Rayment

Option Explicit
Option Base 1

Public Declare Function GetNearestPaletteIndex Lib "gdi32" _
(ByVal hPalette As Long, ByVal crColor As Long) As Long


'Ani Info
Public ANI$    ' Whole file ani string
Public AniTitle$
Public AniAuthor$
Public NumFrames As Long
Public BPP As Long
Public HotX As Long, HotY As Long
Public zDuration As Single
Public RateValues() As Long
Public icontype As Long

Public tWidth As Long, tHeight As Long

'Edit box spacing
Public kspace As Long

Public ImageTransparentColor As Long
Public TransparentColor As Long
Public bred As Byte, bgreen As Byte, bblue As Byte

Public Type BBANI
   cbSizeOf As Long
   cFrames As Long
   cSteps As Long
   cx As Long
   cy As Long
   cBitCount As Long
   cPlanes As Long
   JifRate As Long
   cFlags As Long
End Type
Public anih As BBANI

Public Type ICONH          '62 bytes
   iZero As Integer        '2   0
   iType As Integer        '2   2 cursor
   iNumicons As Integer    '2   1
   bW As Byte              '1  32
   bH As Byte              '1  32
   bNc As Byte             '1   0
   bRes As Byte            '1   0
   iHotX As Integer        '2  HotX
   iHotY As Integer        '2  HotY
   Nbytes As Long          '4  32*32\8, 32*32\2, 32*32, 32*32*3 + 128
   NOffset As Long         '4  22 ?
   
   BMIH As Long            '4  40 BITMAPINFOHEADER
   NWidth As Long          '4  32
   NHeight As Long         '4  64
   NPlanes As Integer      '2   1
   NBPP As Integer         '2   1,4,8,24
   NCom As Long            '4   0
   NImage As Long          '4  32*32\8, 32*32\2, 32*32, 32*32*3
   NHRES As Long           '4   0
   NVRES As Long           '4   0
   NIndexes As Long        '4   2,16,256,0
   NImport As Long         '4   0
End Type
Public iconheader As ICONH
Public pframicon As Long   ' Ptr to framicon  1st icon

Public Type PAL4
   B As Byte
   G As Byte
   R As Byte
   AL As Byte
End Type
Public BGRA() As PAL4

Public NColors As Long
Public ImageSize As Long
' XORA(32, 32, NumFrames) Long  ' Filled from Bits(0-1), Nybbles(0-63) or Bytes(0-255) Index
                                ' to palette or RGB (24bit)
' ANDA(32, 32, NumFrames) Byte  ' Filled from Bits(0-1)
Public XORA() As Long
Public ANDA() As Byte

Private B$
Private p2() As Long
Private ptr As Long
Private ByteArray() As Byte

Private ChunkNames() As Long
Private ANIOUT() As Long
Private ABITS() As Byte
Private XBITS() As Byte

Public Sub READ_ANI_FILE(FSpec$)
Dim p As Long
Dim k As Long
Dim FSize As Long
Dim LNUM As Long
Dim N As Long

   ReDim p2(22)
   Close
   On Error GoTo F_AINERR
   Open FSpec$ For Binary Access Read As #1
   FSize = LOF(1)
   ANI$ = Space$(FSize)
   Get #1, , ANI$
   Close
   
   If Left$(ANI$, 4) <> "RIFF" Then
      MsgBox " Not a RIFF file ", vbCritical, "Reading ani file"
      ANI$ = ""
      FSpec$ = ""
      Exit Sub
   End If
   p = InStr(1, ANI$, "ACONLIST")
'   If p = 0 Then
'      MsgBox " No ACONLIST chunk ", vbInformation, "Reading ani file"
'   End If
   
   p = InStr(1, ANI$, "INFOINAM")
   AniTitle$ = "Untitled"
   If p > 0 Then
      p = p + 8
      B$ = Mid$(ANI$, p, 4)
      CopyMemory LNUM, ByVal B$, 4
      p = p + 4
      AniTitle$ = Mid$(ANI$, p, LNUM)
      k = InStr(1, AniTitle$, Chr$(0))
      If k > 0 Then AniTitle$ = Left$(AniTitle$, k - 1)
   End If
   AniAuthor$ = "Noname"
   p = InStr(1, ANI$, "IART")
   If p > 0 Then
      p = p + 4
      B$ = Mid$(ANI$, p, 4)
      CopyMemory LNUM, ByVal B$, 4
      p = p + 4
      AniAuthor$ = Mid$(ANI$, p, LNUM)
      k = InStr(1, AniAuthor$, Chr$(0))
      If k > 0 Then AniAuthor$ = Left$(AniAuthor$, k - 1)
   End If
   If p = 0 Then p = 1
   ptr = p   ' To avoid chunk names in Title or Author text
            ' NB Assumes ACONLIST INFONAM & IART always
            ' before other chunks !!
   
   p = InStr(ptr, ANI$, "anih")
   If p = 0 Then
      MsgBox " No anih chunk ", vbCritical, "Reading ani file"
      ANI$ = ""
      FSpec$ = ""
      Exit Sub
   End If
   p = p + 4
   
   For k = 1 To 9
      p = p + 4
      B$ = Mid$(ANI$, p, 4)
      CopyMemory p2(k), ByVal B$, 4
   Next k
   With anih
      .cbSizeOf = p2(1)
      .cFrames = p2(2)
      .cSteps = p2(3)
      .cx = p2(4)
      .cy = p2(5)
      .cBitCount = p2(6)
      .cPlanes = p2(7)
      .JifRate = p2(8)
      .cFlags = p2(9)
   End With
   NumFrames = anih.cFrames
   'NumFrames = anih.cSteps
   If NumFrames > 64 Then
      MsgBox Str$(NumFrames) & " frames. Reduced to 64", vbInformation, "Read ani file"
      NumFrames = 64
      anih.cFrames = 64
   End If
   ReDim RateValues(64)
   
'   ' CHECK
'   With anih
'      p2(1) = .cbSizeOf
'      NumFrames = .cFrames
'      p2(2) = .cSteps
'      p2(3) = .cx
'      p2(4) = .cy
'      p2(5) = .cBitCount
'      p2(6) = .cPlanes
'      p2(7) = .JifRate
'      p2(8) = .cFlags
'   End With
'''''''''''''''''''''''''''''''''''''''''''''''
   p = InStr(ptr, ANI$, "rate")
   If p > 0 Then
      p = p + 8 ' rate####
      B$ = Mid$(ANI$, p, 4 * NumFrames)
      CopyMemory RateValues(1), ByVal B$, 4 * NumFrames
   Else   ' No rate table
      For N = 1 To NumFrames
         RateValues(N) = 1
      Next N
   End If
      
   ' First icon
   p = InStr(ptr, ANI$, "framicon")
   If p = 0 Then
      MsgBox " No framicon chunk ", vbCritical, "Reading ani file"
      ANI$ = ""
      FSpec$ = ""
      Exit Sub
   End If
   pframicon = p
   p = p + 12     ' framicon####
   
   B$ = Mid$(ANI$, p, 14)
   CopyMemory iconheader.iZero, ByVal B$, 14
   
   p = p + 14
   B$ = Mid$(ANI$, p, 48)
   CopyMemory iconheader.Nbytes, ByVal B$, 48

'   ' CHECK
'   With iconheader
'      p2(1) = .iZero
'      p2(2) = .iType
'      p2(3) = .iNumicons
'      p2(4) = .bW
'      p2(5) = .bH
'      p2(6) = .bNc
'      p2(7) = .bRes
'      p2(8) = .iHotX
'      p2(9) = .iHotY
'
'      p2(10) = .Nbytes
'      p2(11) = .NOffset
'      p2(12) = .BMIH
'      p2(13) = .NWidth
'      p2(14) = .NHeight
'      p2(15) = .NPlanes
'      p2(16) = .NBPP
'      p2(17) = .NCom
'      p2(18) = .NImage
'      p2(19) = .NHRES
'      p2(20) = .NVRES
'      p2(21) = .NIndexes
'      p2(22) = .NImport
'   End With
   
   ' RETURNED VALUES
   ' AniTitle$
   ' AniAuthor$
   ' NumFrames
   ' RateValues(64)
   
   If iconheader.NHeight <> 64 Or iconheader.NWidth <> 32 Then
      MsgBox " Not a 32 x 32 ani cursor or a Problem", vbCritical, "Reading ani file"
      ANI$ = ""
      FSpec$ = ""
      Exit Sub
   End If
   
   With iconheader
      HotX = .iHotX
      HotY = .iHotY
      BPP = .NBPP
      tWidth = .NWidth
      tHeight = .NHeight \ 2
   End With
   
   If BPP > 24 Then
      MsgBox "BPP > 24", vbCritical, "Reading ani file"
      ANI$ = ""
      FSpec$ = ""
      Exit Sub
   End If
   p = p + 48
   ptr = p
   ReDim XORA(32, 32, NumFrames)
   ReDim ANDA(32, 32, NumFrames)
   ' ptr -> Palette(8B(2 colors) or 64B(16 colors) or .NIndexes*4 B(.NIndexes colors) RGBA,RGBA),
   ' or RGB (24bpp) ' or BGRA(32bpp) of first icon
   ' Assumes all palettes the same  ELSE will need ReDim BGRA(0 To (NColors - 1) * 4, NumFrames)
   NColors = iconheader.NIndexes
   Select Case BPP
   Case 1
      ' Get palette
      If NColors = 0 Then NColors = 2
      ReDim BGRA(0 To (NColors - 1))
      B$ = Mid$(ANI$, ptr, NColors * 4)
      CopyMemory BGRA(0).B, ByVal B$, NColors * 4
      
      For N = 1 To NumFrames
         ptr = ptr + NColors * 4      ' Skip palette ptr ->   next XORA image
         GetXORA_Bits_Indexes N     ' Does    ptr = ptr + ImageSize
         GetANDA_Bits N             ' Does    ptr = ptr + ImageSize
         
         If N = 1 And ANDA(1, 1, N) = 1 Then
            ImageTransparentColor = XORA(1, 1, N)  ' usually black @ pal index 0
         End If
         
         p = InStr(ptr, ANI$, "icon")
         'p = p + 8     ' icon####
         'p = p + 14    ' past header 1 iconheader.iZero
         'ptr = p + 48   ' past header 2 iconheader.Nbytes
         ptr = p + 70
      Next N
   Case 4
      ' Get palette
      If NColors = 0 Then NColors = 16
      ReDim BGRA(0 To (NColors - 1))
      B$ = Mid$(ANI$, ptr, NColors * 4)
      CopyMemory BGRA(0).B, ByVal B$, NColors * 4
      
      For N = 1 To NumFrames
         ptr = ptr + NColors * 4      ' Skip palette ptr ->   next XORA image
         GetXORA_Nybbles_Indexes N  ' Does    ptr = ptr + ImageSize
         GetANDA_Bits N             ' Does    ptr = ptr + ImageSize
         If N = 1 And ANDA(1, 1, N) = 1 Then
            ImageTransparentColor = XORA(1, 1, N)  ' usually black @ pal index 0
         End If
         p = InStr(ptr, ANI$, "icon")
         ptr = p + 70
      Next N
   
   Case 8
      ' Get palette
      If NColors = 0 Then NColors = 256
      ReDim BGRA(0 To (NColors - 1))
      B$ = Mid$(ANI$, ptr, NColors * 4)
      CopyMemory BGRA(0).B, ByVal B$, NColors * 4
      
      For N = 1 To NumFrames
         ptr = ptr + NColors * 4      ' Skip palette ptr ->   next XORA image
         GetXORA_Bytes_Indexes N    ' Does    ptr = ptr + ImageSize
         GetANDA_Bits N             ' Does    ptr = ptr + ImageSize
         If N = 1 And ANDA(1, 1, N) = 1 Then
            ImageTransparentColor = XORA(1, 1, N)  ' usually black @ pal index 0
         End If
         p = InStr(ptr, ANI$, "icon")
         ptr = p + 70
      Next N
   
   Case 24
      ' No palette
      ' ptr -> first image
      For N = 1 To NumFrames
         GetXORA_Bytes_RGB N        ' Does    ptr = ptr + ImageSize
         GetANDA_Bits N             ' Does    ptr = ptr + ImageSize
         If N = 1 And ANDA(1, 1, N) = 1 Then
            ImageTransparentColor = 0  ' Default to black ie no palette here
         End If
         p = InStr(ptr, ANI$, "icon")
         ptr = p + 70
      Next N
   End Select
   Exit Sub
'===============
F_AINERR:
Close
MsgBox "Problem Reading ani file", vbCritical, "Reading ani file"
ANI$ = ""
FSpec$ = ""
On Error GoTo 0
End Sub

Public Sub GetXORA_Bits_Indexes(N As Long)
' ptr -> image BPP=1
Dim ix As Long
Dim iy As Long
Dim k As Long
Dim ibx As Long
Dim bit As Long
Dim cul0 As Long
Dim cul1 As Long
   ImageSize = 32 * 32 \ 8
   B$ = Mid$(ANI$, ptr, ImageSize)
   ReDim ByteArray(4, 32)
   CopyMemory ByteArray(1, 1), ByVal B$, ImageSize
   cul0 = RGB(BGRA(0).B, BGRA(0).G, BGRA(0).R)
   cul1 = RGB(BGRA(1).B, BGRA(1).G, BGRA(1).R)
   For iy = 1 To 32
   ibx = 1
   For ix = 1 To 32 Step 8
      bit = 1
      For k = 7 To 0 Step -1
         If ByteArray(ibx, iy) And bit Then
            XORA(ix + k, iy, N) = cul1
         Else
            XORA(ix + k, iy, N) = cul0
         End If
         bit = bit * 2
      Next k
      ibx = ibx + 1
   Next ix
   Next iy
   ptr = ptr + ImageSize
End Sub

Public Sub GetXORA_Nybbles_Indexes(N As Long)
' ptr -> image BPP=4
Dim ix As Long
Dim iy As Long
Dim LoNyb As Long
Dim HiNyb As Long
Dim CulHi As Long
Dim CulLo As Long
   ImageSize = 32 * 32 \ 2
   B$ = Mid$(ANI$, ptr, ImageSize)
   ReDim ByteArray(16, 32)
   CopyMemory ByteArray(1, 1), ByVal B$, ImageSize
   For iy = 1 To 32
   For ix = 1 To 16
     LoNyb = ByteArray(ix, iy) And &HF&
     HiNyb = (ByteArray(ix, iy) And &HF0&) \ 16
     CulLo = RGB(BGRA(LoNyb).R, BGRA(LoNyb).G, BGRA(LoNyb).B)
     CulHi = RGB(BGRA(HiNyb).R, BGRA(HiNyb).G, BGRA(HiNyb).B)
     'If CulLo <> 0 Or CulHi <> 0 Then Stop
     XORA(2 * ix - 1, iy, N) = CulHi  'RGB(BGRA(HiNyb).R, BGRA(HiNyb).G, BGRA(HiNyb).B)
     XORA(2 * ix, iy, N) = CulLo    'RGB(BGRA(LoNyb).R, BGRA(LoNyb).G, BGRA(LoNyb).B)
   Next ix
   Next iy
   ptr = ptr + ImageSize
   ' ptr -> ANDA
End Sub

Public Sub GetXORA_Bytes_Indexes(N As Long)
' ptr -> image BPP=8
Dim ix As Long
Dim iy As Long
Dim k As Long
   ImageSize = 32 * 32
   B$ = Mid$(ANI$, ptr, ImageSize)
   ReDim ByteArray(32, 32)
   CopyMemory ByteArray(1, 1), ByVal B$, ImageSize
   For iy = 1 To 32
   For ix = 1 To 32
      k = ByteArray(ix, iy)
      XORA(ix, iy, N) = RGB(BGRA(k).R, BGRA(k).G, BGRA(k).B)
   Next ix
   Next iy
   ptr = ptr + ImageSize
End Sub

Public Sub GetXORA_Bytes_RGB(N As Long)
' ptr -> image BPP=24
Dim ix As Long
Dim iy As Long
Dim k1 As Long
Dim k2 As Long
Dim k3 As Long
Dim px As Long
   ImageSize = 32 * 32 * 3
   B$ = Mid$(ANI$, ptr, ImageSize)
   ReDim ByteArray(32 * 3, 32)
   CopyMemory ByteArray(1, 1), ByVal B$, ImageSize
   For iy = 1 To 32
   px = 1
   For ix = 1 To 32
      k1 = ByteArray(px, iy)
      k2 = ByteArray(px + 1, iy)
      k3 = ByteArray(px + 2, iy)
      XORA(ix, iy, N) = RGB(k3, k2, k1)
      px = px + 3
   Next ix
   Next iy
   ptr = ptr + ImageSize
End Sub

Public Sub GetANDA_Bits(N As Long)
' ptr -> mask
Dim ix As Long
Dim iy As Long
Dim k As Long
Dim ibx As Long
Dim bit As Long
   ImageSize = 32 * 32 \ 8
   B$ = Mid$(ANI$, ptr, ImageSize)
   ReDim ByteArray(4, 32)
   CopyMemory ByteArray(1, 1), ByVal B$, ImageSize
   For iy = 1 To 32
   ibx = 1
   For ix = 1 To 32 Step 8
      bit = 1
      For k = 7 To 0 Step -1
         If ByteArray(ibx, iy) And bit Then
            ANDA(ix + k, iy, N) = 1 ' Transparent area
         Else
            ANDA(ix + k, iy, N) = 0 ' Image area
         End If
         bit = bit * 2
      Next k
      ibx = ibx + 1
   Next ix
   Next iy
   ptr = ptr + ImageSize
End Sub

'   ' CHECK
'   Open "B66.txt" For Output As #1
'   For p = 1 To 62
'      Print #1, Hex$(Asc(Mid$(B$, p, 1)))
'   Next p
'   Close

'   B$ = Mid$(ANI$, ptr, ImageSize)
'   ReDim ByteArray(4, 32)
'   CopyMemory ByteArray(1, 1), ByVal B$, ImageSize
'   Open "B66.txt" For Output As #1
'   For py = 1 To 32
'   For px = 1 To 4
'      Print #1, ByteArray(px, py); 'Hex$(Asc(Mid$(B$, p, 1)))
'   Next px
'   Print #1,
'   Next py
'   Close

Public Sub SAVE_ANI_FILE(FSpec$)
Dim FileSize As Long
Dim iconsize As Long
Dim k As Long
Dim j As Long
Dim i As Long
Dim LNUM As Long
Dim L As Long
Dim palSize As Long
Dim N As Long
Dim bIndex As Byte
   FileSize = 68 + 8 * NumFrames                ' All Chunk Names + 8 * NumFrames (icon####)
   
   AniTitle$ = AniTitle$ + Chr$(0)
   L = Len(AniTitle$)
   LNUM = (L + 3) And &HFFFFFFFC
   If LNUM > L Then AniTitle$ = AniTitle$ + String$(LNUM - L, Chr$(0))
   FileSize = FileSize + Len(AniTitle$)
   AniAuthor$ = AniAuthor$ + Chr$(0)
   L = Len(AniAuthor$)
   LNUM = (L + 3) And &HFFFFFFFC
   If LNUM > L Then AniAuthor$ = AniAuthor$ + String$(LNUM - L, Chr$(0))
   FileSize = FileSize + Len(AniAuthor$)
   
   FileSize = FileSize + 36                     ' anih data
   FileSize = FileSize + 4 * NumFrames          ' rate data
   iconsize = 62                                ' iconheader
   Select Case BPP
   Case 1
      iconsize = iconsize + 8    ' palette
      iconsize = iconsize + 128  ' xor image
      iconsize = iconsize + 128  ' and
   Case 4
      iconsize = iconsize + 64   ' palette
      iconsize = iconsize + 512  ' xor image
      iconsize = iconsize + 128  ' and
   Case 8
      iconsize = iconsize + 1024 ' palette
      iconsize = iconsize + 1024 ' xor image
      iconsize = iconsize + 128  ' and
   Case 24
      iconsize = iconsize + 3072 ' xor image
      iconsize = iconsize + 128  ' and
   End Select
   FileSize = FileSize + NumFrames * iconsize
   'FileSize = (FileSize + 3) And &HFFFFFFFC
   
   ReDim ANIOUT(FileSize \ 4)
   
   ReDim ChunkNames(10)
   ChunkNames(1) = 1179011410    ' RIFF
   ChunkNames(2) = 1313817409    ' ACON
   ChunkNames(3) = 1414744396    ' LIST
   ChunkNames(4) = 1330007625    ' INFO
   ChunkNames(5) = 1296125513    ' INAM
   ChunkNames(6) = 1414676809    ' IART
   ChunkNames(7) = 1751740001    ' anih
   ChunkNames(8) = 1702125938    ' rate
   ChunkNames(9) = 1835102822    ' fram
   ChunkNames(10) = 1852793705    ' icon
   
   With iconheader
      .iZero = 0              'As Integer        '2   0
      .iType = 2              'As Integer        '2   2 cursor
      .iNumicons = 1          'As Integer    '2   1
      .bW = 32                'As Byte              '1  32
      .bH = 32                'As Byte              '1  32
      .bNc = 0                'As Byte              '1  0
      .bRes = 0               'As Byte              '1  0
      .iHotX = CInt(HotX)     'As Integer        '2  HotX
      .iHotY = CInt(HotY)     'As Integer        '2  HotY
      Select Case BPP
      Case 1: palSize = 8
              .Nbytes = 40 + palSize + 32 * 32 \ 8 + 128 'As Long
      Case 4: palSize = 64
              .Nbytes = 40 + palSize + 32 * 32 \ 2 + 128 'As Long
      Case 8: palSize = 1024
              .Nbytes = 40 + palSize + 32 * 32 + 128  'As Long
      Case 24: palSize = 0
               .Nbytes = 40 + 32 * 32 * 3 + 128 'As Long
      End Select
      .NOffset = 22           'As Long         '4  22 ?
      .BMIH = 40              'As Long            '4  40 BITMAPINFOHEADER
      .NWidth = 32            'As Long          '4  32
      .NHeight = 64           'As Long         '4  64
      .NPlanes = 1            'As Integer      '2   1
      .NBPP = BPP             'As Integer         '2   1,4,8,24
      .NCom = 0               'As Long            '4   0
      Select Case BPP
      Case 1: .NImage = 32 * 32 \ 8   'As Long
      Case 4: .NImage = 32 * 32 \ 2   'As Long
      Case 8: .NImage = 32 * 32       'As Long
      Case 24: .NImage = 32 * 32 * 3  'As Long
      End Select
      .NHRES = 0              'As Long           '4   0
      .NVRES = 0              'As Long           '4   0
      Select Case BPP
      Case 1: .NIndexes = 2   'As Long        '4   2
      Case 4: .NIndexes = 16  'As Long        '4   16
      Case 8: .NIndexes = 256 'As Long        '4   256
      Case 24: .NIndexes = 0  'As Long        '4   0
      End Select
      .NImport = 0                            'As Long         '4   0
   End With
   ANIOUT(1) = ChunkNames(1)    'RIFF
   ANIOUT(2) = FileSize          ' FileSize
   ANIOUT(3) = ChunkNames(2)    ' ACON
   ANIOUT(4) = ChunkNames(3)   ' LIST
   LNUM = 20 + Len(AniTitle$) + Len(AniAuthor$)
   ANIOUT(5) = LNUM
   ANIOUT(6) = ChunkNames(4)   ' INFO
   ANIOUT(7) = ChunkNames(5)   ' INAM
   LNUM = Len(AniTitle$)
   ANIOUT(8) = LNUM
'  AniTitle$
'    "String" to bytes:
   CopyMemory ANIOUT(9), ByVal AniTitle$, LNUM
   k = 9 + LNUM \ 4
   ANIOUT(k) = ChunkNames(6)   ' IART
   LNUM = Len(AniAuthor$)
   ANIOUT(k + 1) = LNUM

'  AniAuthor$
'    "String" to bytes:
   CopyMemory ANIOUT(k + 2), ByVal AniAuthor$, LNUM
   k = k + 2 + LNUM \ 4
   ANIOUT(k) = ChunkNames(7)   ' anih
   LNUM = 36
   ANIOUT(k + 1) = LNUM
   With anih
      .cbSizeOf = 36
      .cFrames = NumFrames
      .cSteps = NumFrames
      .cx = 0
      .cy = 0
      .cBitCount = 4
      .cPlanes = 1
      .JifRate = 10
      .cFlags = 1
   End With

   With anih
      ANIOUT(k + 2) = .cbSizeOf
      ANIOUT(k + 3) = NumFrames
      ANIOUT(k + 4) = NumFrames
      ANIOUT(k + 5) = .cx
      ANIOUT(k + 6) = .cy
      ANIOUT(k + 7) = .cBitCount
      ANIOUT(k + 8) = .cPlanes
      ANIOUT(k + 9) = .JifRate
      ANIOUT(k + 10) = .cFlags
   End With
   k = k + 11
   ANIOUT(k) = ChunkNames(8)   ' rate
   k = k + 1
   LNUM = 4 * NumFrames
   ANIOUT(k) = LNUM
   k = k + 1
   For j = 1 To NumFrames
      ANIOUT(k) = RateValues(j)
      k = k + 1
   Next j
   ANIOUT(k) = ChunkNames(3)  ' LIST
   k = k + 1
   LNUM = FileSize - 4 * k
   ANIOUT(k) = LNUM
   k = k + 1
   ANIOUT(k) = ChunkNames(9)      ' fram
   ReDim Preserve ANIOUT(k)
   
   ' Set image transparency color
   Dim ix As Long
   Dim iy As Long
   For N = 1 To NumFrames
      For iy = 1 To 32
      For ix = 1 To 32
         If XORA(ix, iy, N) = ImageTransparentColor Or _
            XORA(ix, iy, N) = TransparentColor Then
            If BPP <> 24 Then
               XORA(ix, iy, N) = 0 'RGB(BGRA(0).R, BGRA(0).G, BGRA(0).B)
            Else
               XORA(ix, iy, N) = 0
            End If
         End If
      Next ix
      Next iy
   Next N
   
   On Error Resume Next
   Kill FSpec$
   Open FSpec$ For Binary Access Write As #1
   Put #1, , ANIOUT()
   For N = 1 To NumFrames
      Put #1, , ChunkNames(10) ' icon
      LNUM = iconheader.Nbytes + 22
      Put #1, , LNUM
      Put #1, , iconheader
      If BPP <> 24 Then    ' put out palette
         Put #1, , BGRA()
      End If
      
      GetABITS N
         
      Select Case BPP
      Case 1   ' XORA bits ANDA bits
         GetXBITS N
         Put #1, , XBITS()
         Put #1, , ABITS()
      Case 4   ' XORA nybbles ANDA bits
         For j = 1 To 32
         For i = 1 To 31 Step 2
            bIndex = GetIndex2Nybbles(XORA(i, j, N), XORA(i + 1, j, N))
            Put #1, , bIndex
         Next i
         Next j
         Put #1, , ABITS()
      Case 8   ' XORA bytes ANDA bits
         For j = 1 To 32
         For i = 1 To 32
            bIndex = GetPalIndexByte(XORA(i, j, N))
            Put #1, , bIndex
         Next i
         Next j
         Put #1, , ABITS()
      Case 24  ' XORA RGB bytes ANDA bits
         For j = 1 To 32
         For i = 1 To 32
            LngToRGB XORA(i, j, N)
            Put #1, , bblue
            Put #1, , bgreen
            Put #1, , bred
         Next i
         Next j
         Put #1, , ABITS()
      End Select
   Next N
   Close
End Sub

Private Sub GetABITS(N As Long)
' Public ANDA(32,32,N) bytes
' Private ABITS() as Byte
' N = Frame number
Dim j As Long
Dim k As Long
Dim i As Long
   ReDim ABITS(4, 32)
   For j = 1 To 32
   For k = 1 To 32
      i = (k + 7) \ 8
      ABITS(i, j) = 2 * ABITS(i, j)
      If ANDA(k, j, N) = 1 Then ABITS(i, j) = ABITS(i, j) Or 1
   Next k
   Next j
End Sub

Public Sub GetXBITS(N As Long)
' Public XORA(32,32,N) Long
' Private XBITS() as Byte
' N = Frame number
Dim k As Long
Dim i As Long
Dim j As Long
Dim cul1 As Long
   cul1 = RGB(BGRA(1).R, BGRA(1).G, BGRA(1).B)
   ReDim XBITS(4, 32)
   For j = 1 To 32
   For k = 1 To 32
      i = (k + 7) \ 8
      XBITS(i, j) = 2 * XBITS(i, j)
      If XORA(k, j, N) = cul1 Then XBITS(i, j) = XBITS(i, j) Or 1
   Next k
   Next j
End Sub

Public Function GetIndex2Nybbles(cul1 As Long, Cul2 As Long) As Byte
Dim index1 As Long
Dim index2 As Long
   GetIndex2Nybbles = 0
   index1 = CLng(GetPalIndexByte(cul1))
   index2 = CLng(GetPalIndexByte(Cul2))
   GetIndex2Nybbles = 16 * index1 + index2
End Function

Public Function GetPalIndexByte(Cul As Long) As Byte
Dim k As Long
Dim MinD As Long
Dim LongVal As Long
   GetPalIndexByte = 0
   If ScreenBits < 24 Then
      MinD = 1000&
      LngToRGB Cul
      For k = 0 To NColors - 1
         LongVal = Abs(1& * bred - BGRA(k).R) + _
                 Abs(1& * bgreen - BGRA(k).G) + _
                 Abs(1& * bblue - BGRA(k).B)
         If LongVal < MinD Then
            MinD = LongVal
            GetPalIndexByte = k
         End If
      Next k
   Else
      For k = 0 To NColors - 1
         If Cul = RGB(BGRA(k).R, BGRA(k).G, BGRA(k).B) Then
            GetPalIndexByte = k
            Exit For
         End If
      Next k
   End If
End Function

Public Sub READ_CUR_FILE(FSpec$, icontype As Long)
' icontype 1 ico, 2 cur
Dim p As Long
Dim FSize As Long
Dim ix As Long
Dim iy As Long
   
   ReDim p2(22)
   On Error GoTo F_CURERR
   Open FSpec$ For Binary Access Read As #1
   FSize = LOF(1)
   ANI$ = Space$(FSize)
   Get #1, , ANI$
   Close
      
   p = 1
   B$ = Mid$(ANI$, p, 14)
   CopyMemory iconheader.iZero, ByVal B$, 14
   
   ' Restart at a Long !!
   p = p + 14
   B$ = Mid$(ANI$, p, 8)
   CopyMemory iconheader.Nbytes, ByVal B$, 8

   ' CHECK
'   With iconheader
'      p2(1) = .iZero ' 14B
'      p2(2) = .iType
'      p2(3) = .iNumicons
'      p2(4) = .bW
'      p2(5) = .bH
'      p2(6) = .bNc
'      p2(7) = .bRes
'      p2(8) = .iHotX
'      p2(9) = .iHotY
'
'      p2(10) = .Nbytes  ' 8
'      p2(11) = .NOffset
'   End With
   
   If iconheader.bW <> 32 Or iconheader.bH <> 32 Then
      If iconheader.iNumicons = 1 Then
         MsgBox " Not a 32 x 32 cur/ico ", vbCritical, "Reading cur/ico file"
         ANI$ = ""
         FSpec$ = ""
         Exit Sub
      Else  ' Try 2nd icon
         p = p + 8
         B$ = Mid$(ANI$, p, 8)
         CopyMemory iconheader.bW, ByVal B$, 8
         
         ' Restart at a Long !!
         p = p + 8
         B$ = Mid$(ANI$, p, 8)
         CopyMemory iconheader.Nbytes, ByVal B$, 8
         If iconheader.bW <> 32 Or iconheader.bH <> 32 Then
            MsgBox " 2nd ico/cur not a 32 x 32 ", vbCritical, "Reading cur/ico file"
            ANI$ = ""
            FSpec$ = ""
            Exit Sub
         Else ' 2nd ico/cur 32x32
            ' iconheader.NOffset to 2nd cursor BMIH
            p = iconheader.NOffset + 1
            B$ = Mid$(ANI$, p, 40)
            CopyMemory iconheader.BMIH, ByVal B$, 40
         End If
      End If
   Else  ' 32x32 icon
      If iconheader.iNumicons = 1 Then
         p = p + 8
         B$ = Mid$(ANI$, p, 40)
         CopyMemory iconheader.BMIH, ByVal B$, 40
      Else  ' 1st icon 32x32 but more than 1
         p = iconheader.NOffset + 1
         B$ = Mid$(ANI$, p, 40)
         CopyMemory iconheader.BMIH, ByVal B$, 40
      End If
   End If
   
   ' CHECK
'   With iconheader
'      p2(1) = .iZero ' 6B
'      p2(2) = .iType
'      p2(3) = .iNumicons
'
'      p2(4) = .bW    ' 16B
'      p2(5) = .bH
'      p2(6) = .bNc
'      p2(7) = .bRes
'      p2(8) = .iHotX
'      p2(9) = .iHotY
'      p2(10) = .Nbytes
'      p2(11) = .NOffset
'
'      p2(12) = .BMIH    ' 40B
'      p2(13) = .NWidth
'      p2(14) = .NHeight
'      p2(15) = .NPlanes
'      p2(16) = .NBPP
'      p2(17) = .NCom
'      p2(18) = .NImage
'      p2(19) = .NHRES
'      p2(20) = .NVRES
'      p2(21) = .NIndexes
'      p2(22) = .NImport
'   End With
   
   
   If iconheader.NBPP > 24 Then
      MsgBox "BPP > 24", vbCritical, "Reading cur/ico file"
      ANI$ = ""
      FSpec$ = ""
      Exit Sub
   End If
   With iconheader
      HotX = .iHotX
      HotY = .iHotY
      BPP = .NBPP
      tWidth = .NWidth
      tHeight = .NHeight \ 2
   End With
   
   If icontype = 1 Then
      iconheader.iType = 2 ' convert to cur
      iconheader.iHotX = 0
      iconheader.iHotY = 0
      HotX = 0
      HotY = 0
   End If
   
   p = p + 40  ' Skip past BMIH now
   ptr = p
   'ReDim XORA(32, 32, NumFrames)
   'ReDim ANDA(32, 32, NumFrames)
   For iy = 1 To 32
   For ix = 1 To 32
      XORA(ix, iy, picNum) = 0
      ANDA(ix, iy, picNum) = 0
   Next ix
   Next iy
   ' ptr -> Palette(8B(2 colors) or 64B(16 colors) or .NIndexes*4 B(.NIndexes colors) RGBA,RGBA),
   ' or RGB (24bpp)  or BGRA(32bpp) of first icon
   NColors = iconheader.NIndexes
   Select Case BPP
   Case 1
      ' Get palette
      If NColors = 0 Then NColors = 2
      ReDim BGRA(0 To (NColors - 1))
      B$ = Mid$(ANI$, ptr, NColors * 4)
      CopyMemory BGRA(0).B, ByVal B$, NColors * 4
      
      ptr = ptr + NColors * 4      ' Skip palette ptr ->   next XORA image
      GetXORA_Bits_Indexes picNum     ' Does    ptr = ptr + ImageSize
      GetANDA_Bits picNum
      
      If ANDA(1, 1, picNum) = 1 Then
         ImageTransparentColor = XORA(1, 1, picNum)  ' usually black @ pal index 0
      End If
   Case 4
      ' Get palette
      If NColors = 0 Then NColors = 16
      ReDim BGRA(0 To (NColors - 1))
      B$ = Mid$(ANI$, ptr, NColors * 4)
      CopyMemory BGRA(0).B, ByVal B$, NColors * 4
      
      ptr = ptr + NColors * 4      ' Skip palette ptr ->   next XORA image
      GetXORA_Nybbles_Indexes picNum  ' Does    ptr = ptr + ImageSize
      GetANDA_Bits picNum
      If ANDA(1, 1, picNum) = 1 Then
         ImageTransparentColor = XORA(1, 1, picNum)  ' usually black @ pal index 0
      End If
   Case 8
      ' Get palette
      If NColors = 0 Then NColors = 256
      ReDim BGRA(0 To (NColors - 1))
      B$ = Mid$(ANI$, ptr, NColors * 4)
      CopyMemory BGRA(0).B, ByVal B$, NColors * 4
      
      ptr = ptr + NColors * 4      ' Skip palette ptr ->   next XORA image
      GetXORA_Bytes_Indexes picNum    ' Does    ptr = ptr + ImageSize
      GetANDA_Bits picNum
      If ANDA(1, 1, picNum) = 1 Then
         ImageTransparentColor = XORA(1, 1, picNum)  ' usually black @ pal index 0
      End If
   Case 24
      ' No palette
      GetXORA_Bytes_RGB picNum        ' Does    ptr = ptr + ImageSize
      GetANDA_Bits picNum
      If ANDA(1, 1, picNum) = 1 Then
         ImageTransparentColor = 0  ' Default to black ie no palette here
      End If
   End Select
   
'   '??
'   Select Case BPP
'   Case 1, 4
'   For iy = 1 To 32
'   For ix = 1 To 32
'      ANDA(ix, iy, picNum) = 0
'      If XORA(ix, iy, picNum) = ImageTransparentColor Then
'         'XORA(ix, iy, picNum) = TransparentColor
'         ANDA(ix, iy, picNum) = 1
'      End If
'   Next ix
'   Next iy
'   End Select
   
   Exit Sub
'==========
F_CURERR:
Close
MsgBox "Problem Reading cur/ico file", vbCritical, "Reading cur/ico file"
ANI$ = ""
FSpec$ = ""
On Error GoTo 0
End Sub

Public Sub SAVE_CUR_FILE(FSpec$)
Dim FileSize As Long
Dim iconsize As Long
Dim j As Long
Dim i As Long
Dim palSize As Long
Dim bIndex As Byte
Dim ix As Long
Dim iy As Long
Dim Cul As Long
   iconsize = 62                                ' iconheader
   Select Case BPP
   Case 1
      iconsize = iconsize + 8    ' palette
      iconsize = iconsize + 128  ' xor image
      iconsize = iconsize + 128  ' and
   Case 4
      iconsize = iconsize + 64   ' palette
      iconsize = iconsize + 512  ' xor image
      iconsize = iconsize + 128  ' and
   Case 8
      iconsize = iconsize + 1024 ' palette
      iconsize = iconsize + 1024 ' xor image
      iconsize = iconsize + 128  ' and
   Case 24
      iconsize = iconsize + 3072 ' xor image
      iconsize = iconsize + 128  ' and
   End Select
   FileSize = iconsize
   
   ReDim ANIOUT(FileSize \ 4)
   
   With iconheader
      .iZero = 0              'As Integer    '2   0
      .iType = 2              'As Integer    '2   2 cursor
      .iNumicons = 1          'As Integer    '2   1
      .bW = 32                'As Byte       '1  32
      .bH = 32                'As Byte       '1  32
      .bNc = 0                'As Byte       '1  0
      .bRes = 0               'As Byte       '1  0
      .iHotX = CInt(HotX)     'As Integer    '2  HotX
      .iHotY = CInt(HotY)     'As Integer    '2  HotY
      Select Case BPP
      Case 1: palSize = 8
              .Nbytes = 40 + palSize + 32 * 32 \ 8 + 128 'As Long
      Case 4: palSize = 64
              .Nbytes = 40 + palSize + 32 * 32 \ 2 + 128 'As Long
      Case 8: palSize = 1024
              .Nbytes = 40 + palSize + 32 * 32 + 128  'As Long
      Case 24: palSize = 0
               .Nbytes = 40 + 32 * 32 * 3 + 128 'As Long
      End Select
      .NOffset = 22           'As Long         '4  22 ?
      .BMIH = 40              'As Long         '4  40 BITMAPINFOHEADER
      .NWidth = 32            'As Long         '4  32
      .NHeight = 64           'As Long         '4  64
      .NPlanes = 1            'As Integer      '2   1
      .NBPP = BPP             'As Integer      '2   1,4,8,24
      .NCom = 0               'As Long         '4   0
      Select Case BPP
      Case 1: .NImage = 32 * 32 \ 8   'As Long
      Case 4: .NImage = 32 * 32 \ 2   'As Long
      Case 8: .NImage = 32 * 32       'As Long
      Case 24: .NImage = 32 * 32 * 3  'As Long
      End Select
      .NHRES = 0              'As Long        '4   0
      .NVRES = 0              'As Long        '4   0
      Select Case BPP
      Case 1: .NIndexes = 2   'As Long        '4   2
      Case 4: .NIndexes = 16  'As Long        '4   16
      Case 8: .NIndexes = 256 'As Long        '4   256
      Case 24: .NIndexes = 0  'As Long        '4   0
      End Select
      .NImport = 0                            'As Long         '4   0
   End With
   
   ' Set image transparency color
   For iy = 1 To 32
   For ix = 1 To 32
      If XORA(ix, iy, picNum) = ImageTransparentColor Or _
         XORA(ix, iy, picNum) = TransparentColor Then
         If BPP <> 24 Then
            Cul = 0 'RGB(BGRA(0).R, BGRA(0).G, BGRA(0).B)
            XORA(ix, iy, picNum) = Cul
         Else  ' BPP=24
            XORA(ix, iy, picNum) = 0
         End If
      End If
   Next ix
   Next iy
   
   On Error Resume Next
   Kill FSpec$
   Open FSpec$ For Binary Access Write As #1
      Put #1, , iconheader
      If BPP <> 24 Then    ' put out palette
         Put #1, , BGRA()
      End If
      
      GetABITS picNum
         
      Select Case BPP
      Case 1   ' XORA bits ANDA bits
         GetXBITS picNum
         Put #1, , XBITS()
         Put #1, , ABITS()
      Case 4   ' XORA nybbles ANDA bits
         For j = 1 To 32
         For i = 1 To 31 Step 2
            bIndex = GetIndex2Nybbles(XORA(i, j, picNum), XORA(i + 1, j, picNum))
            Put #1, , bIndex
         Next i
         Next j
         Put #1, , ABITS()
      Case 8   ' XORA bytes ANDA bits
         For j = 1 To 32
         For i = 1 To 32
            bIndex = GetPalIndexByte(XORA(i, j, picNum))
            Put #1, , bIndex
         Next i
         Next j
         Put #1, , ABITS()
      Case 24  ' XORA RGB bytes ANDA bits
         For j = 1 To 32
         For i = 1 To 32
            LngToRGB XORA(i, j, picNum)
            Put #1, , bblue
            Put #1, , bgreen
            Put #1, , bred
         Next i
         Next j
         Put #1, , ABITS()
      End Select
   Close
End Sub


