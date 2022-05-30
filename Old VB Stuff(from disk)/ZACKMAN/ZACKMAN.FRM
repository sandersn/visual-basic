VERSION 5.00
Begin VB.Form frmZackman 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ZackMan: The Computer Game"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ZACKMAN.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   9480
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   6120
   End
   Begin VB.PictureBox picZackrmov 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   600
      Picture         =   "ZACKMAN.frx":23362
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackrmid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1080
      Picture         =   "ZACKMAN.frx":235D4
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackrstl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1680
      Picture         =   "ZACKMAN.frx":23846
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackustl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1680
      Picture         =   "ZACKMAN.frx":23AB8
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackumid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1680
      Picture         =   "ZACKMAN.frx":23D2A
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackumov 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1680
      Picture         =   "ZACKMAN.frx":23F9C
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackdstl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      Picture         =   "ZACKMAN.frx":2420E
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackdmid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      Picture         =   "ZACKMAN.frx":24480
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackdmov 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      Picture         =   "ZACKMAN.frx":246F2
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZacklmov 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      Picture         =   "ZACKMAN.frx":24964
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZacklmid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   720
      Picture         =   "ZACKMAN.frx":24BD6
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZacklstl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1200
      Picture         =   "ZACKMAN.frx":24E48
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4320
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   2880
      Width           =   465
   End
   Begin VB.Image imgDont 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   960
      Top             =   5400
      Width           =   975
   End
   Begin VB.Image imgStop 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   6480
      Top             =   4560
      Width           =   1695
   End
End
Attribute VB_Name = "frmZackman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The variables for the current location of ZackMan.
Dim LeftEdge As Long
Dim TopEdge As Long
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
'To everyone currently reading this:
'Mistakes and pitfalls of this type of game:(so far)
'The most recent: Using pictures as opposed to images, all
'pictures MUST have their Enabled properties set to False
'otherwise they will have the focus when the program loads.
'It is possible, however, to MOVE all of the code in
'Form_KeyDown to pic..._KeyDown(in this case probably ZackMain)
'However, this should be implemented only in an extreme emergency.
'Figuring out how to change screens: This was one of the
'hardest things to figure out.

Dim right_stop As Integer
Dim bottom_stop As Integer
Dim right_dont As Integer
Dim bottom_dont As Integer
Dim right_main As Integer
Dim bottom_main As Integer

Private Sub Form_Click()
    'Show the About Zackman form.(The About screen.)
    frmZackAbout.Show 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Select Case Shift
        Case 4
            Print "Alt was pressed"
        Case 2
            Print "Ctrl was pressed"
        Case 1
            Print "Shift was pressed"
    End Select
End Sub

Private Sub Form_Load()
    'Initialize the Width and Height variables (not letting
    'ZackMan go off the screen.
    FormHeight = frmZackman.ScaleHeight - picZackMain.Height
    FormWidth = frmZackman.ScaleWidth - picZackMain.Width
    'Initialize the picture of picZackMain.
    picZackMain.Picture = picZackrstl.Picture
    'Initialize the move variables so that ZackMan does not
    'start at (0, 0).
    LeftEdge = picZackMain.Left
    TopEdge = picZackMain.Top
    'Tell the computer where ZackMan starts.
    X = 2
    Y = 2
    
    right_stop = imgStop.Left + imgStop.Top
    bottom_stop = imgStop.Top + imgStop.Left
    right_dont = imgDont.Left + imgDont.Top
    bottom_dont = imgDont.Top + imgDont.Left
    
    right_main = picZackMain.Left + picZackMain.Top
    bottom_main = picZackMain.Top + picZackMain.Left
    
    

End Sub

Private Sub MoveScreen(Mover As String)
Cls
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
        frmZackman.Picture = LoadPicture("C:\My Documents\Visual Basic\Old VB Stuff(from disk)\zackman\X2Y2.bmp")
    ElseIf X = 2 And Y = 1 Then  'Upper screen
        frmZackman.Picture = LoadPicture("C:\My Documents\Visual Basic\Old VB Stuff(from disk)\zackman\X1Y2.bmp")
    ElseIf X = 1 And Y = 2 Then  'Left screen
        frmZackman.Picture = LoadPicture("C:\My Documents\Visual Basic\Old VB Stuff(from disk)\zackman\X1Y2.bmp")
    ElseIf X = 2 And Y = 3 Then  'Lower screen
        frmZackman.Picture = LoadPicture("C:\My Documents\Visual Basic\Old VB Stuff(from disk)\zackman\X1Y2.bmp")
    ElseIf X = 3 And Y = 2 Then  'Right screen
        frmZackman.Picture = LoadPicture("C:\My Documents\Visual Basic\Old VB Stuff(from disk)\zackman\X1Y2.bmp")
    Else 'If it is none of the screens, then make sure you
    'can't go off the edge.
        If Mover = "Right" Then
            Beep
            X = X - 1
            Print X
            LeftEdge = FormWidth
            Exit Sub
        ElseIf Mover = "Left" Then
            Beep
            
            X = X + 1
            Print X
            LeftEdge = 0
            Exit Sub
        ElseIf Mover = "Up" Then
            Beep
            Y = Y + 1
            Print Y
            TopEdge = 0
            Exit Sub
        ElseIf Mover = "Down" Then
            Beep
            Y = Y - 1
            Print Y
            TopEdge = FormHeight
            Exit Sub
        End If
    End If
    'Move ZackMan to the other side of the screen.
    'First make him invisible.
    picZackMain.Visible = False
    If Mover = "Right" Then
        LeftEdge = 0
        picZackMain.Move LeftEdge, TopEdge
    ElseIf Mover = "Left" Then
        LeftEdge = FormWidth
        picZackMain.Move LeftEdge, TopEdge
    ElseIf Mover = "Up" Then
        TopEdge = FormHeight
        picZackMain.Move LeftEdge, TopEdge
    ElseIf Mover = "Down" Then
        TopEdge = 0
        picZackMain.Move LeftEdge, TopEdge
    End If
    picZackMain.Visible = True
    'Add code for more screens later...

End Sub

Private Sub picZackMain_Click()
    MsgBox "Hi!!", 16, "Greetings"
End Sub

Private Sub Timer1_Timer()
Static Counter
    Counter = Counter + 1
    If Counter = 5 Then
        Counter = 0
        Timer1.Enabled = False
        Exit Sub
    End If
    Select Case Mover
        Case "Right"
            LeftEdge = LeftEdge + 30
            If Counter = 1 Then
                picZackMain.Picture = picZackrmid.Picture
            ElseIf Counter = 2 Then
                picZackMain.Picture = picZackrmov.Picture
            ElseIf Counter = 3 Then
                picZackMain.Picture = picZackrmid.Picture
            ElseIf Counter = 4 Then
                picZackMain.Picture = picZackrstl.Picture
            End If
            If LeftEdge > FormWidth Then MoveScreen Mover
        Case "Up"
            TopEdge = TopEdge - 30
            If Counter = 1 Then
                picZackMain.Picture = picZackumid.Picture
            ElseIf Counter = 2 Then
                picZackMain.Picture = picZackumov.Picture
            ElseIf Counter = 3 Then
                picZackMain.Picture = picZackumid.Picture
            ElseIf Counter = 4 Then
                picZackMain.Picture = picZackustl.Picture
            End If
            If TopEdge < 0 Then MoveScreen Mover
        Case "Left"
            LeftEdge = LeftEdge - 30
            If Counter = 1 Then
                picZackMain.Picture = picZacklmid.Picture
            ElseIf Counter = 2 Then
                picZackMain.Picture = picZacklmov.Picture
            ElseIf Counter = 3 Then
                picZackMain.Picture = picZacklmid.Picture
            ElseIf Counter = 4 Then
                picZackMain.Picture = picZacklstl.Picture
            End If
            If LeftEdge < 0 Then MoveScreen Mover
        Case "Down"
            TopEdge = TopEdge + 30
            If Counter = 1 Then
                picZackMain.Picture = picZackdmid.Picture
            ElseIf Counter = 2 Then
                picZackMain.Picture = picZackdmov.Picture
            ElseIf Counter = 3 Then
                picZackMain.Picture = picZackdmid.Picture
            ElseIf Counter = 4 Then
                picZackMain.Picture = picZackdstl.Picture
            End If
            If TopEdge > FormHeight Then MoveScreen Mover
    End Select
    picZackMain.Move LeftEdge, TopEdge
End Sub

