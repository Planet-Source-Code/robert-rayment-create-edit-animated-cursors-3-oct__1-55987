VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  AniProg Help"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4635
      IntegralHeight  =   0   'False
      Left            =   2415
      TabIndex        =   2
      Top             =   135
      Width           =   6285
   End
   Begin VB.ListBox List1 
      Height          =   2160
      IntegralHeight  =   0   'False
      Left            =   135
      TabIndex        =   1
      Top             =   855
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&H CLOSE"
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   195
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Contents"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   585
      Width           =   735
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmHelp (APhelp.frm)

Option Explicit

' -----------------------------------------------------------
' Windows APIs -  Function & constants to locate & make Window stay on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const wFlags = &H40 Or &H2

'--------------------------------------------------------------
' Windows APIs - For searching list box
'Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

' NB  lParam needs to be As Long for some functions
' but As Any for Search List Box using LB_FINDSTRINGEXACT
'--------------------------------------------------

Private A$
Private resp As Long

Private Sub Form_Load()
'Public frmHelpLeft As Long   ' frmHelp.Left
'Public frmHelpTop As Long   ' frmHelp.Top

Dim Contents As Long
Dim i As Long
   
   aPlay = False
   If AniTest Then
      RestoreOldCursor
      AniTest = False
      On Error Resume Next
      Kill aniFSpec$
      DoEvents
   End If
   
   ' Size & make form stay on top
   With frmHelp
      resp = SetWindowPos(.hwnd, HWND_NOTOPMOST, _
             frmHelpLeft / STX, frmHelpTop / STY, .Width / STX, .Height / STY, wFlags) ' X,Y,W,H
   End With
   
   frmHelp.Left = frmHelpLeft
   frmHelp.Top = frmHelpTop
      
   frmHelp.Show
   DoEvents
   
   Screen.MousePointer = vbHourglass
   
   Open PathSpec$ & "APHelp.txt" For Input As #1
   Input #1, Contents
   For i = 1 To Contents    ' Number of FVHelp Contents' items
      Line Input #1, A$
      List1.AddItem A$
   Next i
   
   Do Until EOF(1)
      Line Input #1, A$
      List2.AddItem A$
   Loop
   
   Close
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub List1_Click()
'Select item
Dim i As Long
   i = List1.ListIndex
   A$ = List1.List(i) & Chr$(0)
   If Len(A$) <> 0 Then
      'Search List2 for Text$ & place at top
      resp = SendMessageLong(List2.hwnd, LB_FINDSTRINGEXACT, -1&, _
      ByVal A$)
      
      List2.ListIndex = resp
      If List2.ListIndex > 0 Then
         List2.TopIndex = List2.ListIndex - 1
      End If
   End If
End Sub

Private Sub Command1_Click()
   frmHelpLeft = frmHelp.Left
   frmHelpTop = frmHelp.Top
   Unload frmHelp
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmHelpLeft = frmHelp.Left
   frmHelpTop = frmHelp.Top
   Unload Me
End Sub

