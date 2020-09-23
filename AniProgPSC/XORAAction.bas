Attribute VB_Name = "XORAAction"
' XORAAction.bas

Option Explicit
Option Base 1

Public XORATEMP() As Long
Public ANDATEMP() As Byte

Public Indexes() As Byte
Public DBGRA() As PAL4


' Undo/Redo
Public XORABU() As Long
Public ANDABU() As Byte
Public NumBackUps As Long
Public BackUpNumber As Long

Public aCopy As Boolean

' Selection coords
Public ixs1 As Long, iys1 As Long
Public ixs2 As Long, iys2 As Long
Public aSelect As Boolean

' Frame limits
Public NStart As Long
Public NTot As Long
Public NEnd As Long
Public NStep As Long

' Copy/Paste
Private XORACPY2() As Long
Private ANDACPY2() As Byte

Private ix As Long
Private iy As Long
Private TX() As Long
Private BA() As Byte



' Have picNum

Public Sub ChangeLCulforRCul()
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   For iy = iys1 To iys2
   For ix = ixs1 To ixs2
      If XORA(ix, iy, picNum) = LCul Then
         XORA(ix, iy, picNum) = RCul
         If LCul = TransparentColor And RCul <> TransparentColor Then
            ANDA(ix, iy, picNum) = 0
         ElseIf RCul = TransparentColor Then
            ANDA(ix, iy, picNum) = 1
         End If
      End If
   Next ix
   Next iy
End Sub

Public Sub Roll_LR(Button As Integer)
   ReDim TX(32)
   ReDim BA(32)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   If Button = vbLeftButton Then       ' Columns Left
      For iy = iys1 To iys2
         ' Save first column
         TX(iy) = XORA(ixs1, iy, picNum)
         ' Shift rest to left, longs
         CopyMemory XORA(ixs1, iy, picNum), XORA(ixs1 + 1, iy, picNum), (ixs2 - ixs1) * 4
         ' Also ANDA() bytes
         BA(iy) = ANDA(ixs1, iy, picNum)
         CopyMemory ANDA(ixs1, iy, picNum), ANDA(ixs1 + 1, iy, picNum), (ixs2 - ixs1)
      Next iy
      ' Put TX() & BA() to last column
      For iy = iys1 To iys2
         XORA(ixs2, iy, picNum) = TX(iy)
         ANDA(ixs2, iy, picNum) = BA(iy)
      Next iy
   
   ElseIf Button = vbRightButton Then  ' Columns Right
      ' Save last column
      For iy = iys1 To iys2
         TX(iy) = XORA(ixs2, iy, picNum)
         BA(iy) = ANDA(ixs2, iy, picNum)
      Next iy
      ' Shift rest to right
      For iy = iys1 To iys2
      For ix = ixs2 To (ixs1 + 1) Step -1
         XORA(ix, iy, picNum) = XORA(ix - 1, iy, picNum)
         ANDA(ix, iy, picNum) = ANDA(ix - 1, iy, picNum)
      Next ix
      Next iy
      ' Put TX() & BA() to first column
      For iy = iys1 To iys2
         XORA(ixs1, iy, picNum) = TX(iy)
         ANDA(ixs1, iy, picNum) = BA(iy)
      Next iy
   End If
End Sub

Public Sub Roll_UD(Button As Integer)
   ReDim TX(32)
   ReDim BA(32)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   If Button = vbLeftButton Then       ' Rows Up
      
      CopyMemory TX(1), XORA(ixs1, iys2, picNum), (ixs2 - ixs1 + 1) * 4  ' Save top row
      ' Move rows up
      For iy = iys2 To iys1 + 1 Step -1
         CopyMemory XORA(ixs1, iy, picNum), XORA(ixs1, (iy - 1), picNum), (ixs2 - ixs1 + 1) * 4
      Next iy
      CopyMemory XORA(ixs1, iys1, picNum), TX(1), (ixs2 - ixs1 + 1) * 4  ' TX() to bottom row
   
      CopyMemory BA(1), ANDA(ixs1, iys2, picNum), (ixs2 - ixs1 + 1)       ' Save top row
      ' Move rows up
      For iy = iys2 To iys1 + 1 Step -1
         CopyMemory ANDA(ixs1, iy, picNum), ANDA(ixs1, (iy - 1), picNum), (ixs2 - ixs1 + 1)
      Next iy
      CopyMemory ANDA(ixs1, iys1, picNum), BA(1), (ixs2 - ixs1 + 1)    ' BA() to bottom row
   
   ElseIf Button = vbRightButton Then  ' Rows Down
      CopyMemory TX(1), XORA(ixs1, iys1, picNum), (ixs2 - ixs1 + 1) * 4      ' Save bottom row
      ' Move Rows down
      For iy = iys1 To iys2 - 1
         CopyMemory XORA(ixs1, iy, picNum), XORA(ixs1, (iy + 1), picNum), (ixs2 - ixs1 + 1) * 4
      Next iy
      CopyMemory XORA(ixs1, iys2, picNum), TX(1), (ixs2 - ixs1 + 1) * 4      ' TX() to top row
      
      CopyMemory BA(1), ANDA(ixs1, iys1, picNum), 32      ' Save bottom row
      ' Move Rows down
      For iy = iys1 To iys2 - 1
         CopyMemory ANDA(ixs1, iy, picNum), ANDA(ixs1, (iy + 1), picNum), (ixs2 - ixs1 + 1)
      Next iy
      CopyMemory ANDA(ixs1, iys2, picNum), BA(1), (ixs2 - ixs1 + 1)      ' BA() to top row
   
   End If
End Sub

Public Sub Shift_LR(Button As Integer)
   ReDim TX(32)
   ReDim BA(32)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   If Button = vbLeftButton Then       ' Columns Left
      ' Make TX() BA() = 1 TransparentColor, 1
      For iy = iys1 To iys2
         TX(iy) = TransparentColor
         BA(iy) = 1
         CopyMemory XORA(ixs1, iy, picNum), XORA(ixs1 + 1, iy, picNum), (ixs2 - ixs1) * 4
         CopyMemory ANDA(ixs1, iy, picNum), ANDA(ixs1 + 1, iy, picNum), (ixs2 - ixs1)
      Next iy
      ' Put TX() & BA() to column 32
      For iy = iys1 To iys2
         XORA(ixs2, iy, picNum) = TX(iy)
         ANDA(ixs2, iy, picNum) = BA(iy)
      Next iy
   
   ElseIf Button = vbRightButton Then  ' Columns Right
      ' Make TX() BA() = 1 TransparentColor, 1
      For iy = iys1 To iys2
         TX(iy) = TransparentColor
         BA(iy) = 1
      Next iy
      For iy = iys1 To iys2
      For ix = ixs2 To ixs1 + 1 Step -1
         XORA(ix, iy, picNum) = XORA(ix - 1, iy, picNum)
         ANDA(ix, iy, picNum) = ANDA(ix - 1, iy, picNum)
      Next ix
      Next iy
      ' Put TX() & BA() to column 1
      For iy = iys1 To iys2
         XORA(ixs1, iy, picNum) = TX(iy)
         ANDA(ixs1, iy, picNum) = BA(iy)
      Next iy
   End If
End Sub

Public Sub Shift_UD(Button As Integer)
   ReDim TX(32)
   ReDim BA(32)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   If Button = vbLeftButton Then       ' Rows Up
      ' Make TX() BA() = 1 TransparentColor, 1
      For iy = 1 To 32
         TX(iy) = TransparentColor
         BA(iy) = 1
      Next iy
      ' Move rows up
      For iy = iys2 To iys1 + 1 Step -1
         CopyMemory XORA(ixs1, iy, picNum), XORA(ixs1, iy - 1, picNum), (ixs2 - ixs1 + 1) * 4
      Next iy
      CopyMemory XORA(ixs1, iys1, picNum), TX(1), (ixs2 - ixs1 + 1) * 4   ' TX() to Row 1
   
      ' Move rows up
      For iy = iys2 To iys1 + 1 Step -1
         CopyMemory ANDA(ixs1, iy, picNum), ANDA(ixs1, iy - 1, picNum), (ixs2 - ixs1 + 1)
      Next iy
      CopyMemory ANDA(ixs1, iys1, picNum), BA(1), (ixs2 - ixs1 + 1)      ' BA() to Row 1
   
   ElseIf Button = vbRightButton Then  ' Rows Down
      ' Make TX() BA() = 1 TransparentColor, 1
      For iy = 1 To 32
         TX(iy) = TransparentColor
         BA(iy) = 1
      Next iy
      
      ' Move Rows down
      For iy = iys1 To iys2 - 1
         CopyMemory XORA(ixs1, iy, picNum), XORA(ixs1, (iy + 1), picNum), (ixs2 - ixs1 + 1) * 4
      Next iy
      CopyMemory XORA(ixs1, iys2, picNum), TX(1), (ixs2 - ixs1 + 1) * 4     ' TX() to Row 32
      
      ' Move Rows down
      For iy = iys1 To iys2 - 1
         CopyMemory ANDA(ixs1, iy, picNum), ANDA(ixs1, (iy + 1), picNum), (ixs2 - ixs1 + 1)
      Next iy
      CopyMemory ANDA(ixs1, iys2, picNum), BA(1), (ixs2 - ixs1 + 1)     ' BA() to Row 32
   End If
End Sub

Public Sub Rotate90(Button As Integer)
Dim xc As Single, yc As Single
Dim ixs As Long, iys As Long
Dim iSin As Long
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   ReDim XORATEMP(32, 32)
   ReDim ANDATEMP(32, 32)
   CopyMemory XORATEMP(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
   CopyMemory ANDATEMP(1, 1), ANDA(1, 1, picNum), 32 * 32
   
   If Button = vbLeftButton Then
      iSin = 1    ' +90
   Else
      iSin = -1   ' -90
   End If

   xc = (ixs2 + ixs1) \ 2
   yc = (iys2 + iys1) \ 2

   For iy = 1 To 32
   For ix = 1 To 32
      'For each point find rotated source point
      ixs = xc + (iy - yc) * iSin
      iys = yc - (ix - xc) * iSin
      If ixs >= ixs1 Then
      If ixs <= ixs2 Then
      If iys >= iys1 Then
      If iys <= iys2 Then
            XORA(ix, iy, picNum) = XORATEMP(ixs, iys)
            ANDA(ix, iy, picNum) = ANDATEMP(ixs, iys)
      End If
      End If
      End If
      End If
   Next ix
   Next iy
End Sub


Public Sub Flip_LR()
Dim ixL As Long
Dim k As Long, k1 As Long
Dim W As Long
   ReDim TX(1)
   ReDim BA(1)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   W = ixs2 - ixs1 + 1
   ixL = ixs1 + W \ 2 - 1
   k1 = 0
   If (W Mod 2) <> 0 Then
      ixL = ixs1 + W \ 2
   End If
   
   For iy = iys1 To iys2
   k = k1
   For ix = ixs1 To ixL  'ixs1 + (ixs2 - ixs1 + 1) \ 2
      TX(1) = XORA(ixs2 - k, iy, picNum)
      BA(1) = ANDA(ixs2 - k, iy, picNum)
      XORA(ixs2 - k, iy, picNum) = XORA(ix, iy, picNum)
      ANDA(ixs2 - k, iy, picNum) = ANDA(ix, iy, picNum)
      XORA(ix, iy, picNum) = TX(1)
      ANDA(ix, iy, picNum) = BA(1)
      k = k + 1
   Next ix
   Next iy
End Sub

Public Sub Flip_UD()
Dim iyL As Long
Dim k As Long, k1 As Long
Dim H As Long
   ReDim TX(1)
   ReDim BA(1)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   H = iys2 - iys1 + 1
   iyL = iys1 + H \ 2 - 1
   k1 = 0
   If (H Mod 2) <> 0 Then
      iyL = iys1 + H \ 2
   End If
   
   For ix = ixs1 To ixs2
   k = k1
   For iy = iys1 To iyL
      TX(1) = XORA(ix, iys2 - k, picNum)
      BA(1) = ANDA(ix, iys2 - k, picNum)
      XORA(ix, iys2 - k, picNum) = XORA(ix, iy, picNum)
      ANDA(ix, iys2 - k, picNum) = ANDA(ix, iy, picNum)
      XORA(ix, iy, picNum) = TX(1)
      ANDA(ix, iy, picNum) = BA(1)
      k = k + 1
   Next iy
   Next ix
End Sub

Public Sub MirrorLR(Button As Integer)
Dim ixL As Long
Dim k As Long, k1 As Long
Dim W As Long
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   W = ixs2 - ixs1 + 1
   If Button = vbLeftButton Then
      ixL = ixs1 + W \ 2 - 1
      k1 = 0
      If (W Mod 2) <> 0 Then
         ixL = ixs1 + W \ 2
      End If
      
      For iy = iys1 To iys2
      k = k1
      For ix = ixs1 To ixL
         XORA(ixs2 - k, iy, picNum) = XORA(ix, iy, picNum)
         ANDA(ixs2 - k, iy, picNum) = ANDA(ix, iy, picNum)
         k = k + 1
      Next ix
      Next iy
   ElseIf Button = vbRightButton Then
      ixL = ixs1 + W \ 2
      k1 = 1
      If (W Mod 2) <> 0 Then k1 = 0
      
      For iy = iys1 To iys2
      k = k1
      For ix = ixL To ixs2
         XORA(ixL - k, iy, picNum) = XORA(ix, iy, picNum)
         ANDA(ixL - k, iy, picNum) = ANDA(ix, iy, picNum)
         k = k + 1
      Next ix
      Next iy
   End If
End Sub

'################################################################

Public Sub CopyXORA_ANDA()
' Have picNum
   ReDim XORACPY2(32, 32)
   ReDim ANDACPY2(32, 32)
   CopyMemory XORACPY2(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
   CopyMemory ANDACPY2(1, 1), ANDA(1, 1, picNum), 32 * 32
   aCopy = True
End Sub

Public Sub PasteXORA_ANDA()
' Have picNum
   If aCopy Then
      CopyMemory XORA(1, 1, picNum), XORACPY2(1, 1), 32 * 32 * 4
      CopyMemory ANDA(1, 1, picNum), ANDACPY2(1, 1), 32 * 32
   End If
End Sub

Public Sub ClearXORA_ANDA()
   For iy = 1 To 32
   For ix = 1 To 32
      XORA(ix, iy, picNum) = 0 'TransparentColor
      ANDA(ix, iy, picNum) = 1 'Blocks XORA() image
   Next ix
   Next iy
End Sub

Public Sub CopyBU()
   CopyMemory XORABU(1, 1, NumBackUps), XORA(1, 1, picNum), 32 * 32 * 4
   CopyMemory ANDABU(1, 1, NumBackUps), ANDA(1, 1, picNum), 32 * 32
End Sub

Public Sub Delete_picNum()
Dim N As Long
Dim ix As Long
Dim iy As Long
   If picNum < NumFrames Then
      For N = picNum + 1 To NumFrames
         For iy = 1 To 32
         For ix = 1 To 32
            XORA(ix, iy, N - 1) = XORA(ix, iy, N)
            ANDA(ix, iy, N - 1) = ANDA(ix, iy, N)
         Next ix
         Next iy
      Next N
   End If
   picNum = NumFrames
   ClearXORA_ANDA    ' ie picNum
   picNum = 1
   NumFrames = NumFrames - 1
End Sub

Public Sub Swap(Index As Integer)
' NumFrames > 1
   ReDim TX(32, 32)
   ReDim BA(32, 32)
   Select Case Index
   Case 0   ' Swap with Next
      If picNum < NumFrames Then
         CopyMemory TX(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum), XORA(1, 1, picNum + 1), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum + 1), TX(1, 1), 32 * 32 * 4
         CopyMemory BA(1, 1), ANDA(1, 1, picNum), 32 * 32
         CopyMemory ANDA(1, 1, picNum), ANDA(1, 1, picNum + 1), 32 * 32
         CopyMemory ANDA(1, 1, picNum + 1), BA(1, 1), 32 * 32
      End If
   Case 1   ' Swap with Previous
      If picNum > 1 Then
         CopyMemory TX(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum), XORA(1, 1, picNum - 1), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum - 1), TX(1, 1), 32 * 32 * 4
         CopyMemory BA(1, 1), ANDA(1, 1, picNum), 32 * 32
         CopyMemory ANDA(1, 1, picNum), ANDA(1, 1, picNum - 1), 32 * 32
         CopyMemory ANDA(1, 1, picNum - 1), BA(1, 1), 32 * 32
      End If
   Case 2   ' Swap with First
      If picNum > 1 Then
         CopyMemory TX(1, 1), XORA(1, 1, 1), 32 * 32 * 4
         CopyMemory XORA(1, 1, 1), XORA(1, 1, picNum), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum), TX(1, 1), 32 * 32 * 4
         CopyMemory BA(1, 1), ANDA(1, 1, 1), 32 * 32
         CopyMemory ANDA(1, 1, 1), ANDA(1, 1, picNum), 32 * 32
         CopyMemory ANDA(1, 1, picNum), BA(1, 1), 32 * 32
      End If
   Case 3   ' Swap with Last
      If picNum < NumFrames Then
         CopyMemory TX(1, 1), XORA(1, 1, NumFrames), 32 * 32 * 4
         CopyMemory XORA(1, 1, NumFrames), XORA(1, 1, picNum), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum), TX(1, 1), 32 * 32 * 4
         CopyMemory BA(1, 1), ANDA(1, 1, NumFrames), 32 * 32
         CopyMemory ANDA(1, 1, NumFrames), ANDA(1, 1, picNum), 32 * 32
         CopyMemory ANDA(1, 1, picNum), BA(1, 1), 32 * 32
      End If
   End Select
End Sub


Public Sub Gradates(Button As Integer)
' LC to left, RC to right
' Rotation deg,  Reduction %
' NumFrames, picNum (XORA etc picNum+1
Dim N As Long

Dim zangle As Single
Dim zreduc As Single
Dim zIncrAngle As Single
Dim zIncrReduc As Single
Dim zcen As Single
Dim k As Long
Dim ixd As Long
Dim iyd As Long
Dim zrad As Single
Dim zang As Single
Dim zred As Single

Dim xs As Single
Dim ys As Single
Dim ixs As Long
Dim iys As Long
   
   NStart = picNum
   ReDim XORATEMP(32, 32)
   ReDim ANDATEMP(32, 32)
   
   'Copy picNum frame
   CopyMemory XORATEMP(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
   CopyMemory ANDATEMP(1, 1), ANDA(1, 1, picNum), 32 * 32
   
   FrameLimits Button
   
   ' Now NStart =(picnum + 1) or (picNum - 1)
   '     NEnd   = NumFrames or 1
   '     NStep  =  1 or -1
   
   If Rotation > 0 Or Reduction > 0 Then
      For N = NStart To NEnd Step NStep
         For iy = 1 To 32
         For ix = 1 To 32
            XORA(ix, iy, N) = 0 'TransparentColor
            ANDA(ix, iy, N) = 1
         Next ix
         Next iy
      Next N
   End If
   
   zangle = Rotation * d2r# ' radians
   zreduc = Reduction / 100 ' Frac %
   zIncrAngle = zangle / (NTot - 1)
   zIncrReduc = zreduc / (NTot - 1)
   zcen = 16.5
   k = 1
   
   ' picNum frame remains unaltered
   For N = NStart To NEnd Step NStep
      
      If Reduction > 0 And Rotation > 0 Then
         'Copy picNum frame
         CopyMemory XORATEMP(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
         CopyMemory ANDATEMP(1, 1), ANDA(1, 1, picNum), 32 * 32
      End If
      
      If Reduction > 0 Then
         'REDUCE
         zred = 1 - (k * zIncrReduc)
         If zred < 0 Then zred = 0
         For iy = 1 To 32
         For ix = 1 To 32
            zrad = Sqr((iy - zcen) ^ 2 + (ix - zcen) ^ 2)
            zang = zATan2(iy - zcen, ix - zcen)
            ixd = zcen + zrad * zred * Cos(zang)
            iyd = zcen + zrad * zred * Sin(zang)
            
            If ixd > 0 Then
            If ixd < 33 Then
            If iyd > 0 Then
            If iyd < 33 Then
               XORA(ixd, iyd, N) = XORATEMP(ix, iy)
               ANDA(ixd, iyd, N) = ANDATEMP(ix, iy)
            End If
            End If
            End If
            End If
         Next ix
         Next iy
      End If
      
      If Reduction > 0 And Rotation > 0 Then
         'Copy last reduced frame to TEMP
         CopyMemory XORATEMP(1, 1), XORA(1, 1, N), 32 * 32 * 4
         CopyMemory ANDATEMP(1, 1), ANDA(1, 1, N), 32 * 32
      End If
      
      If Rotation > 0 Then
         'ROTATE
         zang = k * zIncrAngle
         For iy = 1 To 32
         For ix = 1 To 32
            xs = zcen + (ix - zcen) * Cos(zang) - (iy - zcen) * Sin(zang)
            ys = zcen + (iy - zcen) * Cos(zang) + (ix - zcen) * Sin(zang)
            ixs = CLng(xs)
            iys = CLng(ys)
            If ixs > 0 Then
            If ixs < 33 Then
            If iys > 0 Then
            If iys < 33 Then
               XORA(ix, iy, N) = XORATEMP(ixs, iys)
               ANDA(ix, iy, N) = ANDATEMP(ixs, iys)
            End If
            End If
            End If
            End If
         Next ix
         Next iy
      End If
      
      k = k + 1
   Next N
End Sub

Public Sub Pepper(Button As Integer)
' Button 1 pepper to left, 2 pepper to right
' Have NumFrames & picNum (XORA etc picNum
' Start = picNum
Dim zRndLim As Single
Dim zPepperFrac As Single
Dim N As Long
Dim k As Long
   
   FrameLimits Button

   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   zPepperFrac = 1 / NumFrames
   For N = picNum To NEnd Step NStep
     ' PEPPER
      zRndLim = 1 - zPepperFrac * k
      For iy = iys1 To iys2
      For ix = ixs1 To ixs2
         ANDA(ix, iy, N) = 1  ' Block image
         If XORA(ix, iy, N) <> 0 Then
            If Rnd < zRndLim Then
               ANDA(ix, iy, N) = 0  ' Pass image thru
            End If
         End If
      Next ix
      Next iy
      k = k + 1
   Next N
End Sub

Public Sub Swirl(Button As Integer)
' Button 1 swirl to left, 2 swirl to right
' Have NumFrames & picNum (XORA etc picNum
' Start = picNum
Dim N As Long
Dim ixs As Long
Dim iys As Long
Dim k As Long

Dim zcenx As Single
Dim zceny As Single
Dim zrad As Single
Dim zang As Single
Dim zk As Single
Dim zm As Single
   NStart = picNum
   
   FrameLimits Button
   
   zm = 0.5 / (NTot - 1)
   k = 1
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   zcenx = (ixs1 + ixs2) / 2
   zceny = (iys1 + iys2) / 2
   
   ' Start frame remains unaltered
   ReDim XORATEMP(32, 32)
   ReDim ANDATEMP(32, 32)
   For N = NStart To NEnd Step NStep
      ' Copy frame N
      CopyMemory XORATEMP(1, 1), XORA(1, 1, N), 32 * 32 * 4
      CopyMemory ANDATEMP(1, 1), ANDA(1, 1, N), 32 * 32
      zk = zm * k + 0.1
      
      ClearFrameN N
      
      For iy = iys1 To iys2
      For ix = ixs1 To ixs2
         zrad = Sqr((iy - zceny) ^ 2 + (ix - zcenx) ^ 2)
         zang = zATan2(iy - zceny, ix - zcenx)
         zang = zang + zrad * zk
         ixs = zcenx + zrad * Cos(zang)
         iys = zceny + zrad * Sin(zang)
         If ixs > 0 Then
         If ixs < 33 Then
         If iys > 0 Then
         If iys < 33 Then
            XORA(ix, iy, N) = XORATEMP(ixs, iys)
            ANDA(ix, iy, N) = ANDATEMP(ixs, iys)
         End If
         End If
         End If
         End If
      Next ix
      Next iy
      k = k + 1
   Next N
End Sub

Public Sub Wave(Button As Integer)
' Button 1 wav to left, 2 wav to right
' Have NumFrames & picNum (XORA etc picNum
' Start = picNum
Dim N As Long
Dim iys As Long
   NStart = picNum
   ReDim XORATEMP(32, 32)
   ReDim ANDATEMP(32, 32)
   
   FrameLimits Button
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   
   ReDim XORATEMP(32, 32)
   ReDim ANDATEMP(32, 32)
   
   ' Start frame remains unaltered
   For N = NStart To NEnd Step NStep
      ' Copy frame N
      CopyMemory XORATEMP(1, 1), XORA(1, 1, N), 32 * 32 * 4
      CopyMemory ANDATEMP(1, 1), ANDA(1, 1, N), 32 * 32
      
      ClearFrameN N
      
      For iy = iys1 To iys2
      For ix = ixs1 To ixs2
         iys = CLng(iy + 2 * Sin(2 * pi# * ix / 24 + N))
         If iys > 0 Then
         If iys < 33 Then
            XORA(ix, iy, N) = XORATEMP(ix, iys)
            ANDA(ix, iy, N) = ANDATEMP(ix, iys)
         End If
         End If
      Next ix
      Next iy
   Next N
End Sub

Public Sub Swivel(Button As Integer)
' Button 1 swivel to left, 2 swivel to right
' Have NumFrames & picNum (XORA etc picNum
' Start = picNum
Dim N As Long
Dim ixd As Long
Dim zIncrReduc As Single
Dim k As Long
Dim zred As Single
Dim zcen As Single
   
   ReDim XORATEMP(32, 32)
   ReDim ANDATEMP(32, 32)
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   End If
   NStart = picNum
   FrameLimits Button
   
   zcen = ixs1 + (ixs2 - ixs1) / 2
   ' Start frame remains unaltered
   k = 1
   zIncrReduc = 1 / (NTot - 1)
   
   For N = NStart To NEnd Step NStep
      zred = 1 - (k * zIncrReduc)
      If zred < 0 Then zred = 0
      ' Copy frame N
      CopyMemory XORATEMP(1, 1), XORA(1, 1, N), 32 * 32 * 4
      CopyMemory ANDATEMP(1, 1), ANDA(1, 1, N), 32 * 32
      
      ClearFrameN N
      
      For iy = iys1 To iys2
      For ix = ixs1 To ixs2
         ixd = zcen + (ix - zcen) * zred
         If ixd > 0 Then
         If ixd < 33 Then
            XORA(ixd, iy, N) = XORATEMP(ix, iy)
            ANDA(ixd, iy, N) = ANDATEMP(ix, iy)
         End If
         End If
      Next ix
      Next iy
      k = k + 1
   Next N
End Sub

Public Sub XFade(Button As Integer)
' Button 1 Fade picNum to First
' Button 2 Fade picNum to Last
' BPP=24 & NumFrames > 2
' Have NumFrames & picNum (XORA etc picNum
' Start = picNum
Dim N As Long
Dim zIncrReduc As Single
Dim k As Long
Dim zred As Single
Dim zcen As Single
   
Dim cul1 As Long
Dim R1 As Long, G1 As Long, B1 As Long
Dim Cul2 As Long
   ReDim ANDATEMP(32, 32)   ' 0s
   
   If Not aSelect Then
      iys1 = 1: iys2 = 32
      ixs1 = 1: ixs2 = 32
   Else  ' Modify picNum mask to make
         ' selection only visible
      CopyMemory ANDA(1, 1, picNum), ANDATEMP(1, 1), 32 * 32  ' 0s
      For iy = 1 To 32
      For ix = 1 To 32
         'If ix,iy outside ixs1,iys1 - ixs2,iys2 then make 1
         If ix < ixs1 Or ix > ixs2 Then
            ANDA(ix, iy, picNum) = 1
         Else  ' x >= ixs1 Or x <= ixs2
            If iy < iys1 Or iy > iys2 Then
               ANDA(ix, iy, picNum) = 1
            End If
         End If
      Next ix
      Next iy
   End If
   NStart = picNum
   FrameLimits Button
   
   zcen = ixs1 + (ixs2 - ixs1) / 2
   ' Start frame remains unaltered unless SR
   k = 1
   zIncrReduc = 1 / (NTot - 1)
   
   For N = NStart To NEnd Step NStep
      zred = 1 - (k * zIncrReduc)
      If zred < 0 Then zred = 0
      
      ' Copy picNum mask to N
      CopyMemory ANDA(1, 1, N), ANDA(1, 1, picNum), 32 * 32
      
      For iy = iys1 To iys2
      For ix = ixs1 To ixs2
         cul1 = XORA(ix, iy, picNum)
         LngToRGB cul1
         R1 = bred
         G1 = bgreen
         B1 = bblue
         Cul2 = XORA(ix, iy, NEnd)
         LngToRGB Cul2
         R1 = R1 * zred + bred * (1 - zred)
         G1 = G1 * zred + bgreen * (1 - zred)
         B1 = B1 * zred + bblue * (1 - zred)
         If R1 < 0 Then R1 = 0
         If R1 > 255 Then R1 = 255
         If G1 < 0 Then G1 = 0
         If G1 > 255 Then G1 = 255
         If B1 < 0 Then B1 = 0
         If B1 > 255 Then B1 = 255
         XORA(ix, iy, N) = RGB(R1, G1, B1)
      Next ix
      Next iy
      k = k + 1
   Next N
End Sub

Public Sub ReverseXORA_ANDA(Button As Integer)
' Button 1 rev to left, 2 rev to right
' Have NumFrames & picNum (XORA etc picNum
' Start = picNum
Dim N As Long
Dim NHalf As Long
   
   If NumFrames > 1 Then
      ReDim XORATEMP(32, 32)
      ReDim ANDATEMP(32, 32)
      
      If Button = vbLeftButton Then
         ' Rev picNum to 1
         NTot = picNum
         NHalf = NTot \ 2 + 1
         NEnd = 1
         NStep = -1
      ElseIf Button = vbRightButton Then
         ' Rev picNum to NumFrames
         NTot = NumFrames - picNum + 1
         NHalf = picNum + NTot \ 2 - 1
         NEnd = picNum + NTot - 1
         NStep = 1
      End If
      For N = picNum To NHalf Step NStep
      
         CopyMemory XORATEMP(1, 1), XORA(1, 1, NEnd - (N - picNum)), 32 * 32 * 4
         CopyMemory XORA(1, 1, NEnd - (N - picNum)), XORA(1, 1, picNum + (N - picNum)), 32 * 32 * 4
         CopyMemory XORA(1, 1, picNum + (N - picNum)), XORATEMP(1, 1), 32 * 32 * 4
         
         CopyMemory ANDATEMP(1, 1), ANDA(1, 1, NEnd - (N - picNum)), 32 * 32
         CopyMemory ANDA(1, 1, NEnd - (N - picNum)), ANDA(1, 1, picNum + (N - picNum)), 32 * 32
         CopyMemory ANDA(1, 1, picNum + (N - picNum)), ANDATEMP(1, 1), 32 * 32
      Next N
   End If
End Sub

Public Sub RollerXORA_ANDA(Button As Integer)
' Rolls forward
' Button 1 roll left , 2 roll right set of frames
' Have NumFrames & picNum (XORA etc picNum
Dim N As Long
   If NumFrames > 1 Then
      ReDim XORATEMP(32, 32)
      ReDim ANDATEMP(32, 32)
      If Button = vbLeftButton Then
         ' Roll picNum to 1
         'Copy picNum
         CopyMemory XORATEMP(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
         CopyMemory ANDATEMP(1, 1), ANDA(1, 1, picNum), 32 * 32
         For N = picNum To 2 Step -1
            CopyMemory XORA(1, 1, N), XORA(1, 1, N - 1), 32 * 32 * 4
            CopyMemory ANDA(1, 1, N), ANDA(1, 1, N - 1), 32 * 32
         Next N
         ' Copy to 1
         CopyMemory XORA(1, 1, 1), XORATEMP(1, 1), 32 * 32 * 4
         CopyMemory ANDA(1, 1, 1), ANDATEMP(1, 1), 32 * 32
         
      ElseIf Button = vbRightButton Then
         ' Roll picNum to NumFrames
         'Copy NumFrames
         CopyMemory XORATEMP(1, 1), XORA(1, 1, NumFrames), 32 * 32 * 4
         CopyMemory ANDATEMP(1, 1), ANDA(1, 1, NumFrames), 32 * 32
         For N = NumFrames To picNum + 1 Step -1
            CopyMemory XORA(1, 1, N), XORA(1, 1, N - 1), 32 * 32 * 4
            CopyMemory ANDA(1, 1, N), ANDA(1, 1, N - 1), 32 * 32
         Next N
         ' Copy to picNum
         CopyMemory XORA(1, 1, picNum), XORATEMP(1, 1), 32 * 32 * 4
         CopyMemory ANDA(1, 1, picNum), ANDATEMP(1, 1), 32 * 32
      End If
   End If
End Sub

Public Sub CopyLR(Button As Integer)
Dim N As Long
   If NumFrames > 1 Then
      ReDim XORATEMP(32, 32)
      ReDim ANDATEMP(32, 32)
      NStart = picNum
      CopyMemory XORATEMP(1, 1), XORA(1, 1, picNum), 32 * 32 * 4
      CopyMemory ANDATEMP(1, 1), ANDA(1, 1, picNum), 32 * 32
      
      FrameLimits Button
      
      For N = NStart To NEnd Step NStep
         CopyMemory XORA(1, 1, N), XORATEMP(1, 1), 32 * 32 * 4
         CopyMemory ANDA(1, 1, N), ANDATEMP(1, 1), 32 * 32
      Next N
   End If
End Sub

Public Sub FrameLimits(Button As Integer)
' NStart=picNum in some cases
'Private NStart As Long
'Private NTot As Long
'Private NEnd As Long
'Private NStep As Long
   If Button = vbLeftButton Then
      ' Act picNum to 1
      NStart = NStart - 1
      NTot = picNum
      NEnd = 1
      NStep = -1
   ElseIf Button = vbRightButton Then
      ' Act picNum to NumFrames
      NTot = NumFrames - NStart + 1
      NStart = NStart + 1
      NEnd = NumFrames
      NStep = 1
   End If
End Sub


Private Sub ClearFrameN(N As Long)
' Clear frame N selected rectangle
'Public ixs1 As Long, iys1 As Long
'Public ixs2 As Long, iys2 As Long
   For iy = iys1 To iys2
   For ix = ixs1 To ixs2
      XORA(ix, iy, N) = 0
      ANDA(ix, iy, N) = 1  ' Block
   Next ix
   Next iy
End Sub

'      'AA ROTATE
'Dim xsf As Single
'Dim ysf As Single
'Dim sumr As Long
'Dim sumg As Long
'Dim sumb As Long
'Dim sumr0 As Long
'Dim sumg0 As Long
'Dim sumb0 As Long
'Dim sumr1 As Long
'Dim sumg1 As Long
'Dim sumb1 As Long
'      xsf = xs - ixs
'      ysf = ys - iys
'      LngToRGB XORATEMP(ix, iy)  'bred,bgreen,bblue
'      sumr = (1 - xsf) * bred: sumg = (1 - xsf) * bgreen: sumb = (1 - xsf) * bblue
'      LngToRGB XORATEMP(ix + 1, iy)
'      sumr0 = sumr + xsf * bred: sumg0 = sumg + xsf * bgreen: sumb0 = sumb + xsf * bblue
'      LngToRGB XORATEMP(ix, iy + 1)
'      sumr = (1 - xsf) * bred: sumg = (1 - xsf) * bgreen: sumb = (1 - xsf) * bblue
'      LngToRGB XORATEMP(ix + 1, iy + 1)
'      sumr1 = sumr + xsf * bred: sumg1 = sumg + xsf * bgreen: sumb1 = sumb + xsf * bblue
'      sumr = (1 - ysf) * sumr0 + ysf * sumr1
'      sumg = (1 - ysf) * sumg0 + ysf * sumg1
'      sumb = (1 - ysf) * sumb0 + ysf * sumb1
'      If sumr < 0 Then sumr = 0
'      If sumr > 255 Then sumr = 255
'      If sumg < 0 Then sumg = 0
'      If sumg > 255 Then sumg = 255
'      If sumb < 0 Then sumb = 0
'      If sumb > 255 Then sumb = 255
'      XORA(ixd, iyd, N) = RGB(sumr, sumg, sumb)
'      ANDA(ixd, iyd, N) = ANDATEMP(ix, iy)
