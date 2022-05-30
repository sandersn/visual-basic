Option Explicit
'------------------------------------------------------------
' Define the character data structure and the global used
' to communicate between the main Map form and the Display
' form.
'------------------------------------------------------------

Global mFlag As Integer
Global frmFlag As Integer

'Variable used in RND Battle
Global stepnum As Integer

Global Const DESC_DEAD = "You're dead.  I guess you won't be needed that chalice after all ... thanks for playing."
       
Type Character
    X As Integer
    Y As Integer
    Symbol As Integer
    Health As Integer
    Weapon As Integer
    MinX As Integer 'used for monsters
    MaxX As Integer 'used for monsters
    MinY As Integer 'used for monsters
    MaxY As Integer 'used for monsters
End Type

'Set Vairables for weapon types
Global WValue As Integer
Global WAttack  As Integer
Global Wevade As Integer
Global WParry As Integer



' result of user choice from the Display form
Global DisplayChoice As Integer

' Constants for 3D Borders
Global Const BORDER_INSET = 0
Global Const BORDER_RAISED = 1

' ScaleMode constants
Global Const TWIPS = 1
Global Const PIXELS = 3

Global Expepool As Integer
Global Atta As Integer
Global Stre As Integer
Global Luck As Integer
Global WeaponAtt As Integer
Global Health As Integer
Global evade As Integer
Global Parry As Integer
Global TechP As Integer
Global Def As Integer
Global Inte As Integer
Global Level As Integer
Global MaxHealth As Integer
Global MaxTechP As Integer
Global TechArray(0 To 50) As String

Sub Make3D (AForm As Form, ctl As Control, ByVal BorderStyle As Integer, ByVal BorderWidth As Integer)
'----------------------------------------------------------------------
' Wrap a 3D effect around a control on a form.
'----------------------------------------------------------------------
' Color Constants
Const DARK_GRAY = &H808080
Const WHITE = &HFFFFFF
Const BLACK = &H0

Dim RightSide As Single
Dim BW As Integer
Dim LeftTopColor As Long, RightBottomColor As Long
Dim i As Integer
Dim SaveMode As Integer

    If Not ctl.Visible Then Exit Sub

    SaveMode = AForm.ScaleMode
    AForm.ScaleMode = PIXELS

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
	AForm.CurrentX = ctl.Left - BW
	AForm.CurrentY = ctl.Top - BW
	AForm.Line -(ctl.Left + ctl.Width + (BW - 1), ctl.Top - BW), LeftTopColor
	' Right
	AForm.Line -(ctl.Left + ctl.Width + (BW - 1), ctl.Top + ctl.Height + (BW - 1)), RightBottomColor
	' Bottom
	AForm.Line -(ctl.Left - BW, ctl.Top + ctl.Height + (BW - 1)), RightBottomColor
	' Left
	AForm.Line -(ctl.Left - BW, ctl.Top - BW), LeftTopColor
    Next

    AForm.ScaleMode = SaveMode
End Sub

