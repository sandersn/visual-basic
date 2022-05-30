Option Explicit
'--------------------------------------------------
' Global variables, constants and declaration.
'--------------------------------------------------

' Functions and constants used to play sounds.
Declare Function sndPlaySound Lib "MMSystem" (ByVal lpsound As String, ByVal flag As Integer) As Integer
Declare Function sndStopSound Lib "MMSystem" Alias "sndPlaySound" (ByVal lpszNull As Long, ByVal flags As Integer) As Integer

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_MEMORY = &H4
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

' Color Constants
Global Const DARK_GRAY = &H808080
Global Const WHITE = &HFFFFFF
Global Const BLACK = &H0

' KeyCode constants
Global Const KEY_LEFT = &H25
Global Const KEY_RIGHT = &H27

' 3D effect constants
Global Const BORDER_INSET = 0
Global Const BORDER_RAISED = 1

' A general purpose data structure used for tracking bitmaps.
' This structure can also be passed to Windows API calls requiring
' a RECT (rectangle structure).
Type tBitMap
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
    Width As Integer
    Height As Integer
End Type

' Windows GDI Bitmap API constants and functions
Global Const SRCCOPY = &HCC0020
Global Const SRCINVERT = &H660046
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

' Windows API rectangle functions
Declare Function IntersectRect Lib "User" (ResultRect As tBitMap, Rect1 As tBitMap, Rect2 As tBitMap) As Integer

Sub Make3D (pic As Form, ctl As Control, ByVal BorderStyle As Integer)
'--------------------------------------------------
' Wrap a 3D effect around a control on a form.
'--------------------------------------------------
Dim AdjustX As Integer, AdjustY As Integer
Dim RightSide As Single
Dim BW As Integer, BorderWidth As Integer
Dim LeftTopColor As Long, RightBottomColor As Long
Dim i As Integer

    If Not ctl.Visible Then Exit Sub

    AdjustX = Screen.TwipsPerPixelX
    AdjustY = Screen.TwipsPerPixelY

    BorderWidth = 3

    Select Case BorderStyle
    Case 0: ' Inset
        LeftTopColor = DARK_GRAY
        RightBottomColor = WHITE
    Case 1: ' Raised
        LeftTopColor = WHITE
        RightBottomColor = DARK_GRAY
    End Select
    

    ' Set the top shading line.
    For BW = 1 To BorderWidth
        ' Top
        pic.CurrentX = ctl.Left - (AdjustX * BW)
        pic.CurrentY = ctl.Top - (AdjustY * BW)
        pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top - (AdjustY * BW)), LeftTopColor
        ' Right
        pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
        ' Bottom
        pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
        ' Left
        pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top - (AdjustY * BW)), LeftTopColor
    Next
End Sub

