Attribute VB_Name = "AniCursor"
' AniCursor.bas

Option Explicit

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Const GCL_HCURSOR As Long = -12


'Public Declare Function CopyImage Lib "user32" _
'(ByVal handle As Long, _
'ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Public Const LR_COPYRETURNORG = &H4
'Public Const IMAGE_BITMAP = 0
'Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
'Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Public Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
'Private Declare Function GetCursor Lib "user32" () As Long
'Public Const OCR_NORMAL As Long = 32512

Public currenthcurs As Long
Public tempcurs As Long
Public tempcurs2 As Long
Public newhcurs As Long
Public AniTest As Boolean

' ORIGINAL CODE WILL NOT RESTORE AN ANIMATED CURSOR
'Public Sub ShowNewCursor(FilePath As String)
'    currenthcurs = GetCursor()
''    tempcurs = CopyIcon(currenthcurs)
'    tempcurs = CopyImage(Form1.PIC.Picture, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
'    newhcurs = LoadCursorFromFile(FilePath)
'    Call SetSystemCursor(newhcurs, OCR_NORMAL)
'End Sub
''
'Public Sub RestoreOldCursor()
'    Call SetSystemCursor(tempcurs, OCR_NORMAL)
'End Sub

'' WILL RESTORE AN ANIMATED CURSOR BUT TEST ANI WILL ONLY SHOW
'' OVER FORM1
Public Sub ShowNewCursor(FilePath As String)
    newhcurs = LoadCursorFromFile(FilePath)
    tempcurs = SetClassLong(Form1.hwnd, GCL_HCURSOR, newhcurs)
    tempcurs2 = SetClassLong(Form1.cmdTestAni.hwnd, GCL_HCURSOR, newhcurs)
End Sub
'
Public Sub RestoreOldCursor()
    tempcurs = SetClassLong(Form1.hwnd, GCL_HCURSOR, tempcurs)
    tempcurs = SetClassLong(Form1.cmdTestAni.hwnd, GCL_HCURSOR, tempcurs2)
End Sub
'
