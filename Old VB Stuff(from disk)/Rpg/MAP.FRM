VERSION 2.00
Begin Form Map 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Role Playing Game"
   ClientHeight    =   6570
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   7875
   ForeColor       =   &H00000000&
   Height          =   7380
   Icon            =   MAP.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7875
   Top             =   -180
   Width           =   7995
   Begin TextBox txtGold 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin TextBox txtHealth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin TextBox txtInfo1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   480
      Width           =   1935
   End
   Begin TextBox txtX 
      Height          =   285
      Left            =   6240
      TabIndex        =   30
      Text            =   "X:"
      Top             =   4320
      Width           =   735
   End
   Begin TextBox txtY 
      Height          =   285
      Left            =   6240
      TabIndex        =   28
      Text            =   "Y:"
      Top             =   4680
      Width           =   735
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   MAP.FRX:0302
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox picMapCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4830
      Left            =   4980
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   29
      Top             =   5760
      Width           =   4830
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   35
      Left            =   4080
      Picture         =   MAP.FRX:057C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   34
      Left            =   3600
      Picture         =   MAP.FRX:087E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   33
      Left            =   3120
      Picture         =   MAP.FRX:0B80
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Characters 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   1680
      Picture         =   MAP.FRX:0DFA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Characters 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   MAP.FRX:1074
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   8
      Left            =   4080
      Picture         =   MAP.FRX:12EE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   17
      Left            =   3840
      Picture         =   MAP.FRX:15F0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   16
      Left            =   3360
      Picture         =   MAP.FRX:1A7A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   15
      Left            =   2880
      Picture         =   MAP.FRX:1F04
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   14
      Left            =   2400
      Picture         =   MAP.FRX:238E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   12
      Left            =   1440
      Picture         =   MAP.FRX:2818
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   13
      Left            =   1920
      Picture         =   MAP.FRX:2CA2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   11
      Left            =   960
      Picture         =   MAP.FRX:312C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   10
      Left            =   480
      Picture         =   MAP.FRX:35B6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   7
      Left            =   3600
      Picture         =   MAP.FRX:3A40
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Characters 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   MAP.FRX:3ECA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   5640
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   3120
      Picture         =   MAP.FRX:4354
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   2640
      Picture         =   MAP.FRX:4656
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Picture1 
      Height          =   4830
      Left            =   120
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   5
      Top             =   240
      Width           =   4830
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   2160
      Picture         =   MAP.FRX:48D0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   1680
      Picture         =   MAP.FRX:4D5A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   MAP.FRX:51E4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   6240
      Width           =   480
   End
   Begin PictureBox Terrain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   MAP.FRX:566E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   6240
      Width           =   480
   End
   Begin Label lblFrame 
      BackStyle       =   0  'Transparent
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   4815
      Left            =   5280
      TabIndex        =   19
      Top             =   360
      Width           =   2415
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Health"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5400
      TabIndex        =   31
      Top             =   2640
      Width           =   690
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Gold"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6240
      TabIndex        =   32
      Top             =   2640
      Width           =   510
   End
   Begin Menu mnuFile 
      Caption         =   "File"
      Begin Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin Menu mnuChar 
      Caption         =   "Chacter"
      Begin Menu mnuAttributes 
         Caption         =   "Attributes"
      End
   End
End
Option Explicit
'------------------------------------------------------------
' Declare all the functions, constants, and variables used
' in MAP.FRM.
'------------------------------------------------------------
'variables for the towns
Dim loadmaps As String
Dim Town As Integer
Dim OldTownX As Integer
Dim OldTownY As Integer

'Variable used in RND Battle
Dim Danger As Integer



' Declare Windows API functions.
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer

Declare Function sndPlaySound Lib "MMSYSTEM" (ByVal lpSound As String, ByVal Flags As Integer) As Integer

' constants for snyPlaySound
Const SND_SYNC = (0)
Const SND_ASYNC = (1)
Const SND_NODEFAULT = (2)
Const SND_MEMORY = (4)
Const SND_LOOP = (8)
Const SND_NOSTOP = (16)

' constants for keystrokes
Const VK_PRIOR = &H21
Const VK_NEXT = &H22
Const VK_END = &H23
Const VK_HOME = &H24
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28

' size of displayed map, in tiles
Const MapXSize = 10
Const MapYSize = 10

' size of entire map, in tiles
Const BigMapXSize = 150
Const BigMapYSize = 150

' constant for BitBlt
Const SRCCOPY = &HCC0020

Const WEAPON_CLUB = 0
Const WEAPON_DAGGER = 1
Const WEAPON_SWORD = 2

Dim FlagSylph As Integer ' True if sylph has been visited
Dim FlagHermit As Integer ' True if the hermit has been visited

Dim StepNumber As Integer 'keeps track of wave file to play

Dim StatusCountdown As Integer  ' keeps track of when status display
                                ' should be cleared

' holds map
Dim MapValues(1 To BigMapXSize) As String
Dim MapTopX As Integer
Dim MapTopY As Integer

' characters in game
Dim NumChars As Integer
Dim Chars(0 To 50) As Character

Const DESC_SYLPH_SICK = "Hello there, my old friend.  It's been far too long since you visited.  You look quite ill.  I can help with that ... but, then I must go.  If you want to cure the poison, you should seek out the Golden Chalice.  Rumor has it you can find it to east."
Const DESC_SYLPH_WELL = "Hello there, my old friend.  It's been far too long since you visited.  I wish I could stick around and talk, but, I must go.  If you want to cure the poison, you should seek out the Golden Chalice.  Rumor has it you can find it to east."
Const DESC_BOATDOCK = "You're on a dock.  There is a rickety old row boat tied up here.  You can see an island off to the west."
Const DESC_FOUNTAIN = "You see a magnificant fountain.  All of the plants around the fountain are thriving.  The water coming from it practically glows.  The water looks so clean, so good.  What do you do?"
Const DESC_BOAT_SINKS = "As you pull up to the dock, the boat springs a leak.  You manage to leap out before it sinks, but, the boat is gone."
Const DESC_CHALICE = "You found it!  Mmmm... you drink from it and feel MUCH better.  Time to wander home and return to the 'hard' life of your pipe and your hammock ... The End"

Const DESC_SLEEPING = "It's late at night, and you're sleeping soundly.  Ahhh, a nice long night's rest is nice after a long day lounging in your hammock smoking your pipe and drinking homebrew.  But, what's that?  Something's in bed with you."
Const DESC_SNAKEBITE = "Ouch!  It bit you ... you feel a pain shooting up your leg.  It's a huge snake with big dripping fangs.  You grab the club you keep next to your bed that you use for rats and quickly dispatch the snake.  It'll make a nice belt now ..."
Const DESC_BEGIN_PLAY = "Unfortunately, you're feeling very, very weak.  If it wasn't for your strong constitution, you'd probably be dead now.  Maybe it's time to go looking for a cure and get some revenge.  Your friend Sylvia the Sylph can help.  She lives in the forest in the valley to the south."
Const DESC_HERMIT = "Leave me alone!  Can't you see that I'm taking care of Clyde?  He's my only friend in the world.  Some nasty ruffians just ran through here and uprooted him.  Now before you go asking a bunch of questions, I'll just tell you what I know ... there's some treasure to the northwest ... and the big evil guy ran straight north.  That's all!  Go away!"

Dim BigTownYSize As Integer

Sub AdjustDisplay ()
'------------------------------------------------------------
' Adjust the Display form's controls based on the value of
' its Tag property.  The Tag property indicates if a bitmap
' will be displayed on the form, or just text.
'------------------------------------------------------------
Dim margin As Integer
Dim BW As Integer

    margin = (9 * Screen.TwipsPerPixelX)
    BW = Display.Width - Display.ScaleWidth

    ' Make room for a bitmap
    If UCase$(Display.Tag) = "SHOW" Then
        Display.Pict.Visible = True

        Display.Width = BW + Display.Pict.Width + (margin * 2) + Display.Description.Width
        Display.Description.Left = margin + Display.Pict.Width
        Display.Choice(0).Left = margin + Display.Pict.Width
        Display.Choice(1).Left = margin + Display.Pict.Width

    ' Just show the text
    Else
        Display.Pict.Visible = False

        Display.Width = BW + (margin * 2) + Display.Description.Width
        Display.Description.Left = margin
        Display.Choice(0).Left = margin
        Display.Choice(1).Left = margin
    End If
    
    ' hide choice button(s) if needed
    Display.Choice(0).Visible = (Len(Display.Choice(0).Caption) <> 0)
    Display.Choice(1).Visible = (Len(Display.Choice(1).Caption) <> 0)

    ' center form on screen
    Display.Move (Screen.Width - Display.Width) / 2, (Screen.Height - Display.Height) / 2

End Sub

Sub Battle ()
'Turning on the Button's
frmBattle.cmdAttack.Enabled = True
frmBattle.cmdTech.Enabled = True
frmBattle.cmdExit.Enabled = False

'Make shure no monsters
mFlag = 0

'Reset the Steps
stepnum = 1

'Can we enter this form
frmFlag = 1

'Turn map form off
map.Enabled = False

'turn the battle form on
frmBattle.Enabled = True

'Show the battle form
frmBattle.Show

'Set the focus on the battle form
frmBattle.SetFocus

'Hide the old Form
map.Hide






End Sub

Sub BoatDockRow (XStart As Integer, XEnd As Integer, XInc As Integer, Symbol As Integer)
'------------------------------------------------------------
' Handle journey across the lake.
'------------------------------------------------------------
Dim I As Integer
Dim TimeTemp As String

    ' play boat rowing sound
    PlaySound "BOAT", False

    ' row to island
    Chars(0).Y = 13

    ' use west headed boat for char
    Chars(0).Symbol = Symbol

    ' move across the lake
    For I = XStart To XEnd Step XInc
        Chars(0).x = I
        CenterMap
        Paintmap

        ' wait a second before moving
        TimeTemp = Time$
        Do While TimeTemp = Time$
        Loop
    Next I

    ' return symbol to normal
    Chars(0).Symbol = 0
End Sub

Sub CenterMap ()
'------------------------------------------------------------
' Keep player centered in map display.
'------------------------------------------------------------

    MapTopX = Chars(0).x - (MapXSize / 2)
    MapTopY = Chars(0).Y - (MapYSize / 2)
End Sub

Sub Chalice ()
'------------------------------------------------------------
' The player has found the Chalice!
'------------------------------------------------------------
    
    PlaySound "CHALICE", False

    ' display the chalice and winning message
    ShowChoice DESC_CHALICE, "", "Exit", 2

    ' exit the program
    End
End Sub

Sub Dead ()
'------------------------------------------------------------
' Uh-oh. Your dead!
'------------------------------------------------------------
    
    ' display the chalice and winning message
      ShowChoice DESC_DEAD, "", "Exit", 4

    ' exit the program
    End
End Sub

Sub defualt ()
terrain(0) = LoadPicture(app.Path & "\Desert1.ico")
terrain(2) = LoadPicture(app.Path & "\mt.ico")
terrain(5) = LoadPicture(app.Path & "\vbridge7.ico")
terrain(6) = LoadPicture(app.Path & "\Hbridge7.ico")
terrain(7) = LoadPicture(app.Path & "\Cabin5.ico")
End Sub

Sub ExitButton_Click ()
'-----------------------------------------------------------
' Exit the program
'------------------------------------------------------------
    
    End
End Sub

Sub Form_KeyDown (KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------
' Handle the player's request to move to another tile.
'------------------------------------------------------------


Randomize Timer
'Movment variables
Dim OldX As Integer
Dim OldY As Integer
Dim OldMapX As Integer
Dim OldMapY As Integer
Dim NewTerrain As Integer

    
    ' save the old position of the character / screen
    OldX = Chars(0).x
    OldY = Chars(0).Y
    OldMapX = MapTopX
    OldMapY = MapTopY

    ' get the new position of the character
    Select Case KeyCode
        Case VK_DOWN:
            If Chars(0).Y < BigMapYSize Then
                Chars(0).Y = Chars(0).Y + 1
            End If
        Case VK_UP:
            If Chars(0).Y > 1 Then
                Chars(0).Y = Chars(0).Y - 1
            End If
        Case VK_LEFT:
            If Chars(0).x > 1 Then
                Chars(0).x = Chars(0).x - 1
            End If
        Case VK_RIGHT:
            If Chars(0).x < BigMapXSize Then
                Chars(0).x = Chars(0).x + 1
            End If
        Case VK_HOME:
            If Chars(0).x > 1 Then
                Chars(0).x = Chars(0).x - 1
            End If

            If Chars(0).Y > 1 Then
                Chars(0).Y = Chars(0).Y - 1
            End If
        Case VK_PRIOR:
            If Chars(0).x < BigMapXSize Then
                Chars(0).x = Chars(0).x + 1
            End If

            If Chars(0).Y > 1 Then
                Chars(0).Y = Chars(0).Y - 1
            End If
        Case VK_NEXT:
            If Chars(0).x < BigMapXSize Then
                Chars(0).x = Chars(0).x + 1
            End If

            If Chars(0).Y < BigMapYSize Then
                Chars(0).Y = Chars(0).Y + 1
            End If
        Case VK_END:
            If Chars(0).x > 1 Then
                Chars(0).x = Chars(0).x - 1
            End If

            If Chars(0).Y < BigMapYSize Then
                Chars(0).Y = Chars(0).Y + 1
            End If
    End Select

    'Build the new Terrain
    NewTerrain = TerrainVal(Mid(MapValues(Chars(0).Y), Chars(0).x, 1))

    ' if this is water or mountain, don't move
    If NewTerrain = 2 Or NewTerrain = 4 Then
        Chars(0).x = OldX
        Chars(0).Y = OldY


        ' we don't need to move
        Chars(0).x = OldX
        Chars(0).Y = OldY
    
        'Paint the map
        Paintmap
    
    End If
    
    'Set the Sextant
    txtX = "X:" & Chars(0).x
    txtY = "Y:" & Chars(0).Y

    If (Chars(0).x <> OldX) Or (Chars(0).Y <> OldY) Then
        Dim TempResult As Integer
    End If
    
    If Chars(0).x = 5 And Chars(0).Y = 5 And Expepool >= 300 Then
       Gain
    End If
    

        
        'Another Baby Step please
        stepnum = stepnum + 1
        
        Danger = Rnd * 10
        
        'Ok Its Battle Time
        If Town = 0 Then
            If stepnum + Danger >= 20 Then
                Battle
            End If
        End If
        
        If Chars(0).x = 5 And Chars(0).Y = 5 Then
            txtinfo1 = "You Walk in to your house and find a mace on the floor"
            WValue = 1
            Weaponsload

        ElseIf StatusCountdown > 0 Then
            ' count down
            StatusCountdown = StatusCountdown - 1
        End If

            
        If NewTerrain = 35 Then
            Chalice
        End If

        ' move to the next step sound
        StepNumber = 3 - StepNumber

        ' recompute x position of screen
        If ((Chars(0).x - MapTopX) < 3) And (Chars(0).x > 3) Then
            MapTopX = Chars(0).x - 3
        ElseIf ((Chars(0).x - MapTopX) > 6) And (Chars(0).x < 198) Then
            MapTopX = Chars(0).x - 6
        End If

        ' recompute y position of screen
        If ((Chars(0).Y - MapTopY) < 3) And (Chars(0).Y > 3) Then
            MapTopY = Chars(0).Y - 3
        ElseIf ((Chars(0).Y - MapTopY) > 6) And (Chars(0).Y < 198) Then
            MapTopY = Chars(0).Y - 6
        End If

        Paintmap
    

    ' check for special stuff
    If (Chars(0).x = 41) And (Chars(0).Y = 18) Then
        LakePath
    ElseIf (Chars(0).x = 66) And (Chars(0).Y = 52) Then
        SecretCave
    ' check for sign at entrance to player's valley
    ElseIf (Chars(0).x = 6) And (Chars(0).Y = 9) Then
    
    ElseIf ((Chars(0).x = 62) And (Chars(0).Y = 70)) Or ((Chars(0).x = 78) And (Chars(0).Y = 63)) Then
    
    ElseIf (Chars(0).x = 7) And (Chars(0).Y = 32) Then
        SylviaTheSylph
    ' check for boat/island
    ' check for hermit
    ElseIf (Chars(0).x = 126) And (Chars(0).Y = 57) Then
        Hermit
    ElseIf (Chars(0).x = 127) And (Chars(0).Y = 6) Then
    
    ElseIf (Chars(0).x = 127) And (Chars(0).Y = 5) Then
    
    End If
    
    'Entering Zetser
    If Chars(0).x = 59 And Chars(0).Y = 9 Then
        OldTownX = Chars(0).x
        OldTownY = Chars(0).Y
        Zetser
        Town = 1
        Chars(0).x = 25
        Chars(0).Y = 3
        picMapCanvas.Cls
        loadmaps = "\Zetser.txt"
        loadmap
    End If
    
    'Leaving the town so soon
    If Town = 1 And Chars(0).Y >= 49 Then
        Chars(0).x = OldTownX
        Chars(0).Y = OldTownY
        defualt
        picMapCanvas.Cls
        loadmaps = "\Map.txt"
        loadmap
        Town = 0
    End If
    
    'Leaving the town so soon
    If Town = 1 And Chars(0).x >= 49 Then
        Chars(0).x = OldTownX
        Chars(0).Y = OldTownY
        defualt
        picMapCanvas.Cls
        loadmaps = "\Map.txt"
        loadmap
        Town = 0
    End If
    
    'Leaving the town so soon
    If Town = 1 And Chars(0).Y <= 2 Then
        Chars(0).x = OldTownX
        Chars(0).Y = OldTownY
        defualt
        picMapCanvas.Cls
        loadmaps = "\Map.txt"
        loadmap
        Town = 0
    End If
    
    'Leaving the town so soon
    If Town = 1 And Chars(0).x <= 2 Then
        Chars(0).x = OldTownX
        Chars(0).Y = OldTownY
        defualt
        picMapCanvas.Cls
        loadmaps = "\Map.txt"
        loadmap
        Town = 0
    End If

End Sub

Sub Form_Load ()
'------------------------------------------------------------
' Load the map information and run the opening sequence.
'------------------------------------------------------------
    
    
    
    'Make Shure its null
    frmFlag = 0
    
    'Set Heros Values
    Atta = 10
    Stre = 10
    luck = 4
    WeaponAtt = 10
    Health = 100
        txtHealth = Health
    TechP = 20
    Evade = 10
    Parry = 3
    inte = 5
    Expepool = 1
    Level = 1
    MaxHealth = 100
    MaxTechP = 20
    'end hero Values

    ' set up help file
    app.HelpFile = app.Path & "\RPG.HLP"

    'no Money Honey
    txtGold = 0

    ' make sure that things are random
    Randomize

    ' set the height
    map.Height = 6000
    
    ' load picture form, but don't show
    Load PictureForm
    Load Display
    
    loadmaps = "\map.txt"
    ' load map from disk
    loadmap

    ' set current coordinates
    MapTopX = 1
    MapTopY = 1

    ' set initial character position
    Chars(0).x = 5
    Chars(0).Y = 5
    SetHealth 0, 1

    

    ' start up monsters / NPCs
    FlagSylph = False
    
    'Baby Steps(this is the first one)
    StepNumber = 1

    ' show initial messages
    PlaySound "SNORE", False

    ShowChoice DESC_SLEEPING, "", "Continue ...", 5

    ' show snake screen
    PlaySound "BITE", False

    ShowChoice DESC_SNAKEBITE, "", "Continue ...", 1

    ShowChoice DESC_BEGIN_PLAY, "", "Begin Play", 4

    ' paint initial map
    Paintmap

    ' Center the main map form.
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    
    'Sextant
    txtX = "X:" & Chars(0).x
    txtY = "Y:" & Chars(0).Y
    
    'Let it know that we are not in a town
    Town = 0

    
    frmBattle.Hide
End Sub

Sub Form_Paint ()
    Make3D Me, Picture1, BORDER_RAISED, 2
    Make3D Me, lblFrame, BORDER_RAISED, 2
End Sub

Sub Form_Unload (Cancel As Integer)
'------------------------------------------------------------
' End the program.
'------------------------------------------------------------
    
    End
End Sub

Sub Gain ()
Level = Level + 1
txtinfo1 = "You enter the training room and you go up a level yo are now at level " & Level
If Level = 2 Then
    Atta = Atta + 3
    Stre = Stre + 5
    luck = luck + 9
    inte = inte + 4
    Evade = Evade + 2
    Parry = Parry + 1
    Def = Def + 6
    'Health
    'TechP
End If

End Sub

Sub GainHealth ()
txtHealth = Health
End Sub

Sub Hermit ()
'------------------------------------------------------------
' The player encounters the Hermit.
'------------------------------------------------------------
    
    ' if we've been here before, don't do it again
    If FlagHermit Then
        Exit Sub
    End If

    ' play the Sylvia's Theme
    PlaySound "HERMIT", False
    
    ' make sure we don't come back here again
    FlagHermit = True
    
    ' give hermit's message to player
    ShowChoice DESC_HERMIT, "", "Continue ...", 3

End Sub

Sub imgEast_Click ()
'------------------------------------------------------------
' Move east via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_RIGHT, 0
End Sub

Sub imgNE_Click ()
'------------------------------------------------------------
' Move northeast via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_PRIOR, 0
End Sub

Sub imgNorth_Click ()
'------------------------------------------------------------
' Move north via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_UP, 0
End Sub

Sub imgNW_Click ()
'------------------------------------------------------------
' Move northwest via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_HOME, 0
End Sub

Sub imgSE_Click ()
'------------------------------------------------------------
' Move southeast via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_NEXT, 0
End Sub

Sub imgSouth_Click ()
'------------------------------------------------------------
' Move south via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_DOWN, 0
End Sub

Sub imgSW_Click ()
'------------------------------------------------------------
' Move southwest via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_END, 0
End Sub

Sub imgWest_Click ()
'------------------------------------------------------------
' Move west via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_LEFT, 0
End Sub

Sub LakePath ()
'------------------------------------------------------------
' The player has encountered a secret lake path.
'------------------------------------------------------------
    
    ' check to see if we've done this before
    If Mid(MapValues(17), 40, 1) = "C" Then
        ' already done, get out
        Exit Sub
    End If

    PlaySound "LAKEPATH", False

    ' tell user they found a path
    

    ' show the path to the user - This actually modifies the map.
    StrOverwrite MapValues(17), "C", 40
    Paintmap

    StrOverwrite MapValues(17), "E", 39
    Paintmap

    StrOverwrite MapValues(16), "C", 38
    Paintmap

    StrOverwrite MapValues(16), "F", 37
    Paintmap

    StrOverwrite MapValues(17), "D", 36
    Paintmap

    StrOverwrite MapValues(17), "E", 35
    Paintmap

    StrOverwrite MapValues(16), "C", 34
    Paintmap

    StrOverwrite MapValues(16), "E", 33
    Paintmap
End Sub

Sub loadmap ()

'------------------------------------------------------------
' Read the map file and store in MapValues string array.
'------------------------------------------------------------



picMapCanvas = LoadPicture("")

Dim FileHandle As Integer ' file handle for read
Dim YIndex                ' index for map read

    ' get handle to use
    FileHandle = FreeFile

    
    
    
    On Error Resume Next
    Open app.Path & "" & loadmaps For Input As #FileHandle

    If Err = 0 Then
        For YIndex = 1 To BigMapYSize
            Line Input #FileHandle, MapValues(YIndex)
        Next YIndex
    End If

    Close #FileHandle

    On Error GoTo 0
End Sub

Sub mnuAttributes_Click ()
map.Enabled = False
map.Hide
CharA.Enabled = True
CharA.Show
CharA.SetFocus
End Sub

Sub mnuExit_Click ()
End

End Sub

Sub Paintmap ()
'------------------------------------------------------------
' Build a new map bitmap in memory (using Windows API calls),
' then BitBlt this new bitmap into the on-screen PictureBox.
'------------------------------------------------------------
Dim x
Dim XIndex, YIndex
Dim CharacterIndex As Integer

    For XIndex = 0 To MapXSize - 1
        For YIndex = 0 To MapYSize - 1
            x = BitBlt(picMapCanvas.hDC, XIndex * 32, YIndex * 32, 32, 32, terrain(TerrainVal(Mid(MapValues(YIndex + MapTopY), XIndex + MapTopX, 1))).hDC, 0, 0, SRCCOPY)
        Next YIndex
    Next XIndex

    For CharacterIndex = 0 To NumChars
        If (Chars(CharacterIndex).x >= MapTopX) And (Chars(CharacterIndex).x < MapTopX + MapXSize) And (Chars(CharacterIndex).Y >= MapTopY) And (Chars(CharacterIndex).Y < MapTopY + MapYSize) And (Chars(CharacterIndex).Health > 0) Then
            ' paint character
            x = BitBlt(picMapCanvas.hDC, (Chars(CharacterIndex).x - MapTopX) * 32, (Chars(CharacterIndex).Y - MapTopY) * 32, 32, 32, Characters(Chars(CharacterIndex).Symbol).hDC, 0, 0, SRCCOPY)
        End If
    Next CharacterIndex

    x = BitBlt(Picture1.hDC, 0, 0, MapXSize * 32, MapYSize * 32, picMapCanvas.hDC, 0, 0, SRCCOPY)

End Sub

Sub PaintMapSquare (x As Integer, Y As Integer)
'------------------------------------------------------------
' Redraw just a single tile on the map.
'------------------------------------------------------------
Dim Temp

    If (x <> Chars(0).x) Or (Y <> Chars(0).Y) Then
        Temp = BitBlt(Picture1.hDC, (x - MapTopX) * 32, (Y - MapTopY) * 32, 32, 32, terrain(TerrainVal(Mid(MapValues(Y), x, 1))).hDC, 0, 0, SRCCOPY)
    Else
        Temp = BitBlt(Picture1.hDC, (x - MapTopX) * 32, (Y - MapTopY) * 32, 32, 32, Characters(0).hDC, 0, 0, SRCCOPY)
    End If
End Sub

Sub PaintMapX ()
'------------------------------------------------------------
' Build a new map bitmap in memory (using Windows API calls),
' then BitBlt this new bitmap into the on-screen PictureBox.
'------------------------------------------------------------
Dim x
Dim XIndex, YIndex
Dim CharacterIndex As Integer
Dim hdcMem As Integer
Dim hbitmap As Integer
Dim hbitmapOld As Integer

    hdcMem = CreateCompatibleDC(Picture1.hDC)
    hbitmap = CreateCompatibleBitmap(Picture1.hDC, MapXSize * 32, MapYSize * 32)
    hbitmapOld = SelectObject(hdcMem, hbitmap)

    For XIndex = 0 To MapXSize - 1
        For YIndex = 0 To MapYSize - 1
            x = BitBlt(hdcMem, XIndex * 32, YIndex * 32, 32, 32, terrain(TerrainVal(Mid(MapValues(YIndex + MapTopY), XIndex + MapTopX, 1))).hDC, 0, 0, SRCCOPY)
        Next YIndex
    Next XIndex

    For CharacterIndex = 0 To NumChars
        If (Chars(CharacterIndex).x >= MapTopX) And (Chars(CharacterIndex).x < MapTopX + MapXSize) And (Chars(CharacterIndex).Y >= MapTopY) And (Chars(CharacterIndex).Y < MapTopY + MapYSize) And (Chars(CharacterIndex).Health > 0) Then
            ' paint character
            x = BitBlt(hdcMem, (Chars(CharacterIndex).x - MapTopX) * 32, (Chars(CharacterIndex).Y - MapTopY) * 32, 32, 32, Characters(Chars(CharacterIndex).Symbol).hDC, 0, 0, SRCCOPY)
        End If
    Next CharacterIndex

    x = BitBlt(Picture1.hDC, 0, 0, MapXSize * 32, MapYSize * 32, hdcMem, 0, 0, SRCCOPY)

    hbitmapOld = SelectObject(hdcMem, hbitmapOld)
    x = DeleteObject(hbitmap)
    x = DeleteDC(hdcMem)
End Sub

Sub Picture1_Paint ()
'------------------------------------------------------------
' Repaint the map as necessary.
'------------------------------------------------------------
    
    ' ensure we've got a good map
    Paintmap
End Sub

Sub PlaySound (SoundName As String, Exclusive As Integer)
'------------------------------------------------------------
' Play a wave audio file.
'------------------------------------------------------------
Dim TempResult As Integer

    If Exclusive Then
        TempResult = sndPlaySound(app.Path & "\" & SoundName & ".WAV", SND_NODEFAULT)
    Else
        TempResult = sndPlaySound(app.Path & "\" & SoundName & ".WAV", SND_NODEFAULT + SND_ASYNC)
    End If
End Sub

Sub SecretCave ()
'------------------------------------------------------------
' Reveal the Secret Cave.
'------------------------------------------------------------
    
    ' check to see if we've done this before
    If Mid(MapValues(51), 66, 1) = "A" Then
        ' already done, get out
        Exit Sub
    End If

    ' play sound
    PlaySound "SECCAVE", False
    
    ' tell user they found a cave
    

    ' show the cave to the user - This actually modifies the map.
    StrOverwrite MapValues(51), "A", 66
    Paintmap

    StrOverwrite MapValues(50), "H", 65
    Paintmap

    StrOverwrite MapValues(49), "0", 65
    StrOverwrite MapValues(48), "0", 64
    StrOverwrite MapValues(48), "X", 65
    StrOverwrite MapValues(48), "0", 66
    StrOverwrite MapValues(47), "0", 65
    Paintmap
End Sub

Sub SetHealth (CharNum As Integer, NewHealth As Integer)
txtHealth = Health
    ' set the new health
    Chars(CharNum).Health = NewHealth

    ' this is creature is dead ...
    If Chars(CharNum).Health <= 0 Then
        ' make it scream
        PlaySound "DEAD", True
        Chars(CharNum).Health = 0

        ' if this is the player, end the game
        If CharNum = 0 Then
            Dead
        End If
    End If

    ' if this is the user
    If CharNum = 0 Then
        ' display the new health
        
    End If

End Sub

Sub ShowChoice (Description As String, ChoiceA As String, ChoiceB As String, PicNum As Integer)
'------------------------------------------------------------
' When the player encounters something (a monster or object,
' for example), the player may be asked to make a choice.
' This choice is presented in the Display form.  This routine
' sets property values and shows the Display form.
'------------------------------------------------------------

    If PicNum >= 0 Then
        Display.Pict.Picture = PictureForm.Picture1(PicNum)
        
        Display.Tag = "SHOW"
    Else
        Display.Tag = ""
    End If

    Display.Description.Caption = Description
    Display.Choice(0).Caption = ChoiceA
    Display.Choice(1).Caption = ChoiceB
    AdjustDisplay
    Display.Show 1

End Sub

Sub StrOverwrite (InputStr As String, OverStr As String, Offset As Integer)
'------------------------------------------------------------
' Used to modify a specified location (or locations) in the
' map array.
'------------------------------------------------------------
    
    InputStr = Left(InputStr, Offset - 1) & OverStr & Mid(InputStr, Offset + Len(OverStr))
End Sub

Sub SylviaTheSylph ()
'------------------------------------------------------------
' Handle visiting Sylvia the Sylph
'------------------------------------------------------------
    
    ' if we've been here before, get out

    ' play the Sylvia's Theme
    PlaySound "SYLPH", False
    
    ' make sure we don't come back here again
    FlagSylph = True
    
    ' if the character is wounded
    If Chars(0).Health < 3 Then
        ShowChoice DESC_SYLPH_SICK, "Accept Sylvia's Help", "Refuse Her Help", 0
    
        If DisplayChoice = 0 Then
            ' play the healing sound
            PlaySound "CURE", False
    
            ' increase the character's health
             Health = MaxHealth
             TechP = MaxTechP
             GainHealth
        End If
    Else
        ' the character is fully healed already
        ShowChoice DESC_SYLPH_WELL, "", "Continue Your Adventure", 0
    End If

End Sub

Function TerrainVal (C As String) As Integer
'------------------------------------------------------------
' Returns a terrain value.
'------------------------------------------------------------
    
    TerrainVal = InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", C) - 1
End Function

Sub Weaponsload ()
If WValue = 1 Then
    Wattack = 10
    WEvade = 2
    WParry = 3
End If


End Sub

Sub Zetser ()


terrain(2).Picture = LoadPicture(app.Path & "\Wall.ico")
terrain(0).Picture = LoadPicture(app.Path & "\Board.ico")
terrain(7).Picture = LoadPicture(app.Path & "\Road.ico")
terrain(5).Picture = LoadPicture(app.Path & "\Road.ico")
terrain(6).Picture = LoadPicture(app.Path & "\RCross.ico")
Chars(0).x = 1
Chars(0).Y = 25
BigTownYSize = 60

End Sub

