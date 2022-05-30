VERSION 2.00
Begin Form frmKisa 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Knight in Shining Armor"
   ClientHeight    =   5685
   ClientLeft      =   660
   ClientTop       =   630
   ClientWidth     =   8280
   Height          =   6375
   KeyPreview      =   -1  'True
   Left            =   600
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   540
   Top             =   0
   Width           =   8400
   Begin Timer tmrEnemymov 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   1080
      Top             =   5400
   End
   Begin Timer tmrKisamov 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   480
      Top             =   5400
   End
   Begin PictureBox picSpriteWork 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   8655
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin PictureBox picSpriteMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   8655
      ScaleHeight     =   465
      ScaleWidth      =   450
      TabIndex        =   2
      Top             =   1110
      Width           =   480
   End
   Begin PictureBox picSpriteImage 
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   8640
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin PictureBox picBG 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5250
      Left            =   0
      ScaleHeight     =   5250
      ScaleWidth      =   8220
      TabIndex        =   0
      Top             =   -45
      Width           =   8220
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuNewGame 
         Caption         =   "&New Game"
      End
      Begin Menu mnuSep1 
         Caption         =   "-"
      End
      Begin Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin Menu mnuHelp 
      Caption         =   "&Help"
      Begin Menu mnuKISAHelp 
         Caption         =   "&Help for KISA"
         Shortcut        =   {F1}
      End
      Begin Menu mnuAboutKISA 
         Caption         =   "&About KISA"
      End
   End
End
Option Explicit
Dim SavedeyBoardSpeed As Long
Dim SavedKeyBoardDelay As Long
'Optimized keyboard delay and speed rate constants.
Const KEYBOARDDELAY = 0
Const KEYBOARDSPEED = 31
'Start DIMing the characters in the game
Dim Kisa As tChar
Dim Enemy(8) As tChar
'The variables used for not letting ZackMan go off the
'screen.
Dim FormHeight As Long
Dim FormWidth As Long
'The variables to keep track of the current screen.
Dim X, Y As Integer
'The variable to tell the timer which button was pushed.
Dim Mover As String
'I just found out that you cannot have Move as a variable.
'Hence, it is now Mover(the variable).
Const NUMENEMY = 8

Sub AnimatedSpriteMove (ASprite As tSprite, Row As Integer, FrameNum As Integer)
'-----------------------------------------------------
' Move the animated sprite to its next position.
'-----------------------------------------------------
Dim Dummy As Integer
Static SpriteNum As Integer
    
    ' Copy a section of the large bitmap into the work area.
    Dummy = BitBlt(ASprite.Work.hDC, 0, 0, ASprite.Work.Width, ASprite.Work.Height, picBG.hDC, Kisa.Sprite.Left, Kisa.Sprite.Top, SRCCOPY)
    
    ' Draw the sprite mask into the work area.                                                                   These two tell \/ which position \/ to get the frame from
    Dummy = BitBlt(ASprite.Work.hDC, Increment, Increment, Kisa.Sprite.Mask.Width, Kisa.Sprite.Mask.Height, Kisa.Sprite.Mask.hDC, 32 * FrameNum, 32 * MaskRow, SRCAND)

    ' The picEarthSprites bitmap contains 8 "frames".  Each frame
    ' is displayed in sequence to animate the object.
    Dummy = BitBlt(ASprite.Work.hDC, Increment, Increment, Kisa.Sprite.Width, Kisa.Sprite.Height, Kisa.Sprite.hDC, 32 * FrameNum, 32 * Row, SRCPAINT)
    
    
    ' Copy the work area onto the background.
    Dummy = BitBlt(ASprite.BG.hDC, Kisa.Left - Increment, Kisa.Top - Increment, picWork.ScaleWidth, picWork.ScaleHeight, picWork.hDC, 0, 0, SRCCOPY)

    DoEvents
End Sub

Function BlockCollided (A As tSprite, B As Image) As Integer
'Check if the bitmap (A), and the image control (B), overlap
Dim ACenterY, BCenterY As Integer
Dim ACenterX, BCenterX As Integer

    ACenterY = (A.Height \ 2) + A.Top
    BCenterY = (A.Height \ 2) + B.Top
    ACenterX = (A.Width \ 2) + A.Left
    BCenterX = (B.Width \ 2) + B.Top
    BlockCollided = False
    'See if they intersect in the same Y range.
    If Abs(ACenterY - BCenterY) < ((A.Height + B.Height) \ 2) Then
        'see if they intersect in the same X range.
        If Abs(ACenterX - BCenterX) < ((A.Width + B.Width) \ 2) Then
            BlockCollided = True
        End If
    End If
End Function

Sub ChangeEnemyPic (Sprite As String)
    
End Sub

Function Collided (A As tSprite, B As tSprite) As Integer
'Check if rectangles (bitmaps) intersect,
'using the IntersectRect API call.
Dim ResultRect As tBitmap
    'Calculate the right and bottom of the rectangles needed by the
    'IntercectRect API call.
    A.Right = A.Left + A.Width - 1
    A.Bottom = A.Top - 1
    B.Right = B.Left + B.Width - 1
    B.Bottom = B.Top - 1
    'IntersectRect will only return 0 (false) if the two
    'rectangles do NOT intersect.
    Collided = IntersectRect(ResultRect, A, B)


End Function

Sub Form_KeyDown (KeyCode As Integer, Shift As Integer)
  'The Select Case to detect which button was pushed.
    Select Case KeyCode
        Case KEY_RIGHT
            Mover = "Right"
            Timer1.Enabled = True
        Case KEY_UP
        'Now repeat the process for all four arrow keys.
            Mover = "Up"
            Timer1.Enabled = True
        Case KEY_LEFT
            Mover = "Left"
            Timer1.Enabled = True
        Case KEY_DOWN
            Mover = "Down"
            Timer1.Enabled = True
        Case KEY_SPACE
            'Maybe put a weapon or something in here.
            '(Probably not yet)
    End Select
End Sub

Sub Form_Load ()
Dim Dummy As Integer
Dim lParam As Long

    'First, get the keyboard speed.
    Dummy = SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, lParam, 0)
    'Next, save the old keyboard speed.
    SavedKeyBoardSpeed = lParam
    'Now do the same thing for keyboard delay.
    Dummy = SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, lParam, 0)
    SavedKeyBoardDelay = lParam
    'Now, SET the keyboard speed and delay.
    Dummy = SystemParametersInfo(SPI_SETKEYBOARDSPEED, KEYBOARDSPEED, 0, 0)
    Dummy = SystemParametersInfo(SPI_SETKEYBOARDDELAY, KEYBOARDDELAY, 0, 0)
    'Initialize the Width and Height variables (not letting
    'ZackMan go off the screen.
    FormHeight = frmZackMan.ScaleHeight - picZackMain.Height
    FormWidth = frmZackMan.ScaleWidth - picZackMain.Width

    'Now SpriteInit the various sprites.
    SpriteInit Kisa.Sprite, frmKisa, picSpriteImage, picSpriteMask, picBG, 30
    For Counter = 0 To NUMENEMY - 1
        SpriteInite Enemy(Counter).Sprite, frmKisa, picSpriteImage, picSpriteMask, picBG, 30
    Next
    'Initialize the move variables so that Kisa does not
    'start at (0, 0).
    Kisa.Left = frmKisa.ScaleWidth \ 2
    Kisa.Top = frmKisa.ScaleHeight \ 2
    'Tell the computer what screen Kisa starts at.
    X = 2
    Y = 2
End Sub

Sub Form_Unload (Cancel As Integer)
Dim Dummy As Integer
    Dummy = SystemParametersInfo(SPI_SETKEYBOARDDELAY, SavedKeyBoardDelay, 0, 0)
    Dummy = SystemParametersInfo(SPI_SETKEYBOARDSPEED, SavedKeyBoardSpeed, 0, 0)
End Sub

Sub mnuExit_Click ()
    End
End Sub

Sub MoveScreen (Mover As String)
Dim Counter As Integer
    If Mover = "Right" Then
        X = X + 1
    ElseIf Mover = "Left" Then
        X = X - 1
    ElseIf Mover = "Up" Then
        Y = Y - 1
    ElseIf Mover = "Down" Then
        Y = Y + 1
    End If
    If X = 2 And Y = 2 Then 'Center screen
        frmZackMan.Picture = LoadPicture("C:\vbprog\zackman\X2Y2.bmp")
        For Counter = 0 To NUMENEMY - 1
            Enemy(Counter).Active = False
        Next
        Counter = 0
    ElseIf X = 2 And Y = 1 Then  'Upper screen
        frmZackMan.Picture = LoadPicture("C:\vbprog\zackman\X1Y2.bmp")
        Enemy(0).Active = True
        Enemy(0).Name = "mudBlob"
        SpriteUpdatePic frmKisa, Enemy(0).Sprite, picMudBlob, picMudBlob_m
        Enemy(1).Active = True
        Enemy(1).Name = "mudBlob"
        SpriteUpdatePic frmKisa, Enemy(0).Sprite, picModBlob, picMudBlob_m
        Enemy(2).Active = True
        Enemy(2).Name = "mudBlob"
        SpriteUpdatePic frmKisa, Enemy(0).Sprite, picMudBlob, picMudBlob_m
    ElseIf X = 1 And Y = 2 Then  'Left screen
        frmZackMan.Picture = LoadPicture("C:\vbprog\zackman\X1Y2.bmp")
    ElseIf X = 2 And Y = 3 Then  'Lower screen
        frmZackMan.Picture = LoadPicture("C:\vbprog\zackman\X1Y2.bmp")
    ElseIf X = 3 And Y = 2 Then  'Right screen
        frmZackMan.Picture = LoadPicture("C:\vbprog\zackman\X1Y2.bmp")
    Else 'If it is none of the screens, then make sure you
    'can't go off the edge.
        If Mover = "Right" Then
            X = X - 1
            LeftEdge = FormWidth
            Exit Sub
        ElseIf Mover = "Left" Then
            X = X + 1
            LeftEdge = 0
            Exit Sub
        ElseIf Mover = "Up" Then
            Y = Y + 1
            TopEdge = 0
            Exit Sub
        ElseIf Mover = "Down" Then
            Y = Y - 1
            TopEdge = FormHeight
            Exit Sub
        End If
    End If
    'Move ZackMan to the other side of the screen.
    'First make him invisible.
    picZackMain.Visible = False
    If Mover = "Right" Then
        Kisa.Left = 0
        picZackMain.Move LeftEdge, TopEdge
    ElseIf Mover = "Left" Then
        Kisa.Left = FormWidth
        picZackMain.Move LeftEdge, TopEdge
    ElseIf Mover = "Up" Then
        Kisa.Top = FormHeight
        picZackMain.Move LeftEdge, TopEdge
    ElseIf Mover = "Down" Then
        Kisa.Top = 0
        picZackMain.Move LeftEdge, TopEdge
    End If
    picZackMain.Visible = True
    'Add code for more screens later...
End Sub

Sub tmrEnemymov_Timer ()
Dim Counter As Integer
    For Counter = 0 To NUMENEMY - 1
        If Enemy(Counter).Active = False Then
        Else
            If Enemy(Counter).Name = "stmKnight" Then
                Increment = 10
            ElseIf Enemy(Counter).Name = "mudBlob" Then
                Increment = 3
            End If 'Add more enemy possibilities later...
            Enemy(Counter).Left = Enemy
        End If
    Next
End Sub

Sub tmrKisamov_Timer ()
Static FrameNum
Dim BlockCount As Integer
    Select Case Mover
        Case "Right"
            Kisa.Left = Kisa.Left + 10
            Row = 2
            AnimatedSpriteMove Row, FrameNum
            If Kisa.Left > FormWidth Then MoveScreen Mover
        Case "Up"
            Kisa.Top = Kisa.Top - 10
            Row = 3
            AnimatedSpriteMove Row, FrameNum
            If TopEdge < 0 Then MoveScreen Mover
        Case "Left"
            Kisa.Left = Kisa.Left - 10
            Row = 1
            AnimatedSpriteMove
            If LeftEdge < 0 Then MoveScreen Mover
        Case "Down"
            Kisa.Top = Kisa.Top + 10
            Row = 4
            AnimatedSpriteMove Row, FrameNum
            If TopEdge > FormHeight Then MoveScreen Mover
    End Select
    For BlockCount = 0 To NumBlocks - 1
        If BlockCollided(Kisa, imgBlock(BlockCount)) Then
            If Mover = "Left" Then
                Kisa.Left = Kisa.Left + 10
            ElseIf Mover = "Right" Then
                Kisa.Left = Kisa.Left - 10
            ElseIf Mover = "Top" Then
                Kisa.Top = Kisa.Top + 10
            ElseIf Mover = "Bottom" Then
                Kisa.Top = Kisa.Top - 10
            End If
            AnimatedSpriteMove Row, FrameNum
        End If
    Next
    FrameNum = FrameNum + 1
    If FrameNum = 5 Then
        FrameNum = 0
        Timer1.Enabled = False
        Exit Sub
    End If
End Sub

