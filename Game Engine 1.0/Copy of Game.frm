VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmGame 
   Caption         =   "Game Engine 1.0 Demo"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwPossessions 
      Height          =   6615
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   11668
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327680
      Icons           =   "imlThings"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MouseIcon       =   "Game.frx":0000
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Name"
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "Desc"
         Object.Tag             =   ""
         Text            =   "Desc Number"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "Movement"
         Object.Tag             =   ""
         Text            =   "Movement Type"
         Object.Width           =   2514
      EndProperty
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4680
      Top             =   5040
   End
   Begin VB.PictureBox picViewport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   120
      MousePointer    =   4  'Icon
      ScaleHeight     =   318
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin ComctlLib.ImageList imlTerrain 
      Left            =   4560
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   44
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":001C
            Key             =   "Bush"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":0336
            Key             =   "Cave Floor"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":0650
            Key             =   "Pool"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":096A
            Key             =   "Cave Wall"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":0C84
            Key             =   "Fire"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":0F9E
            Key             =   "Forest"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":12B8
            Key             =   "Lawn"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":15D2
            Key             =   "Gravel"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":18EC
            Key             =   "House"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1C06
            Key             =   "Mountain"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1F20
            Key             =   "Stalagmite"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":223A
            Key             =   "Fruit Tree"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2554
            Key             =   "Water"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":286E
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":34C0
            Key             =   "Brick Wall"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3CF2
            Key             =   "Carpet"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4544
            Key             =   "Door"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4D96
            Key             =   "Windowed Door"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":55E8
            Key             =   "Cobbles"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":623A
            Key             =   "Krops"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6E8C
            Key             =   "Dead Krops"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":7ADE
            Key             =   "CFlower"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":8730
            Key             =   "MFlower"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":9382
            Key             =   "KFlower"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":9FD4
            Key             =   "SeaSandUp"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":AC26
            Key             =   "SeaSandLt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":B878
            Key             =   "SeaSandRt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":C4CA
            Key             =   "SeaSandDn"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":D11C
            Key             =   "Sand"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":DD6E
            Key             =   "Sea"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":E9C0
            Key             =   "Tile"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":F612
            Key             =   "Dirt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":10264
            Key             =   "Dandelions"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":10EB6
            Key             =   "Grass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":11B08
            Key             =   "Tracks Left"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1275A
            Key             =   "Tracks right"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":133AC
            Key             =   "Tracks up"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":13FFE
            Key             =   "Tracks down"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":14C50
            Key             =   "Stone Walk"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":158A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":164F4
            Key             =   "Leafy Bush"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":17146
            Key             =   "Blue Bush"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":17D98
            Key             =   "Boring Grass"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":189EA
            Key             =   "Pomarbo"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlThings 
      Left            =   4560
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   29
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1963C
            Key             =   "Purina Table"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1A28E
            Key             =   "Potted Bush"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1AEE0
            Key             =   "blank(do not use)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1BB32
            Key             =   "Potted Palm"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1C784
            Key             =   "Clay Pot"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1D3D6
            Key             =   "Haystack"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1E028
            Key             =   "Iron Pot"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1EC7A
            Key             =   "Inscribed Pot"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1F8CC
            Key             =   "Ballot Box"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2051E
            Key             =   "Brick"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":21170
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":21DC2
            Key             =   "Mikey le Mouse"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":22054
            Key             =   "The Chatty Lady"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":22CA6
            Key             =   "Miney le Mouse"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":22F38
            Key             =   "Professor"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":23B8A
            Key             =   "Mega Mouse"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":247DC
            Key             =   "Fred the Freeloader"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2542E
            Key             =   "Macky Le Mouse"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":256C0
            Key             =   "Kat"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":26312
            Key             =   "Easy Fix"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":26B64
            Key             =   "Shuffler"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":273B6
            Key             =   "Grandpa Clone"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":27C08
            Key             =   "GrandPa #13"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2845A
            Key             =   "Ganwa"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":28CAC
            Key             =   "Flame Warpher"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":294FE
            Key             =   "Lady Bug"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2A150
            Key             =   "Bad Spider"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2ADA2
            Key             =   "Live Flower"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2B9F4
            Key             =   "Giant Roach"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAd 
      Caption         =   "Your Ad Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   4815
   End
   Begin VB.Label lblPosition 
      Caption         =   "Position Label"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Menu mnuContext 
      Caption         =   "&Context"
      Visible         =   0   'False
      Begin VB.Menu mnuWhatsThis 
         Caption         =   "&What's This?"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExamine 
         Caption         =   "E&xamine"
      End
      Begin VB.Menu mnuGet 
         Caption         =   "&Get"
      End
      Begin VB.Menu mnuUse 
         Caption         =   "&Use"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "&Drop"
      End
      Begin VB.Menu mnuTalk 
         Caption         =   "&Talk"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By (who else) Nathan Sanders
'History of Game:
'(09/26/1998)
'Just finished talk code. It is not certified bug free, but looks like what has been tested,
'works.(most of the commands that alter the map or your inventory are as yet untested, but
'the question, chthread, and chat commands work to a 'T'. Just as long as you don't try to
'give the player anything as a reward, I know that everything will be fine for your script.
'Oh, I have found it is *extremely* necessary to make no mistakes in your jumptable. It will
'tolerate exactly NO mistakes.              NEW NOTE:all current script commands worked but have not been documented properly!
'Recent bug fixes(i.e. the ones I can remember) are the switch from using the Key to using the Tag
'to store the index of an object of the listview for the same object in your pack. Also fixed
'afterwards to accomodate TWO digits when you added or removed something from your pack(and hence,
'the listview)
'(30/09/1998)1. Fixed a bug where the position of a number inside the 'Thread' command was incorrectly figured by
'subtracting the position of the '-' from the length of the string instead of the position of the ':' from the position
'of the '-'.(Two places, also present in the 'IsThread' command) Sidenote: Just fixed a very similar bug in the Isthread
'command. This one misfigured the Right$ function, which is rather unique, for me.
'So now the engine is substantially complete and is just waiting to be:
'   1.Played
'   2.Game designed for
'   3.and Upgraded.
'(02/10/1998)
'   1.Fixed a position error where the two row choice buttons were too far off the edge of the screen except for
'1024x768 and above.
'   2.Fixed an error wherein when you nested 'if'(e.g. have, question, etc.) statements in the script file, the
'parser could not distinguish between the yes/no block of a question and the yes/no block of a have. Fixed by adding
'new syntax: each question block is ansyes/ansno; each have is havyes/havno; each isthread is isyes/isno.
'also, you must have a /question line after both yes/no blocks for all 'if' statements.(/have, /isthread)
'   3.Added an additional search for the 'remove'. If an object with the desc number supplied is not on the current
'screen, 'remove' now searches the whole object array(NOT the whole disk file!!!, for those of you who are confused.)
'   4.Added a changemap function. This operates by directly changing the map after parsing the absolute position and the
'new tile value from the line of script.
'   (08/10)5. Changed the code underpicViewport_Mousedown so that mnuGet only displays if you are within +-1 space of the
'object.
'    6.Added a complete keyboard interface(finally). Now you can...
'   (29/12) 7.Added the Use code finally. Now you can finally do Use on things. Unfortunately, I took the easy way out and simply did a select case
' inside of mnuUse to determine the action taken. Oh, well; that means that you have to edit the mnuUse function to create a different game now...
'   8.re-ordered the select case's inside InterpretScriptLine and InterpretStoredScript for (probably) optimal speed.
'   9.Oops, lots of other undocumented changes between 6. and 7. I know that one of them was adding InterpretStoredScript so you can now call
'all the cool functions in InterpretScriptLine separately. Also added *a lot* of constants. You should never need to use a real number again!
'Also optimized a couple of if's somewhere...
'   10.Added five new script commands: x+, x-, y+, y-, and warp. They are fairly self explanatory. However, when called in InterpretStoredScript, pass
' NONE inside StoredScript.Desc to act upon you rather than an object in the array. Otherwise StoredScript.Desc is the index to the array. InterpretStoredScript
'will then fetch the object from the array and behave appropiateley. Note that I *am* designing this to work with StoreScriptCommand, so if x+, etc. makes
'the object run off the screen, it will be saved, and saved, and saved. Since the StoreScriptCommand array might run out of room,
'when you have somebody walk off-screen, just walk them along until they're
'out of sight, and then 'warp' them to where they are supposed to be. That way only one or two commands will be Stored.
'Games have done this since the hills began, so don't feel bad that this is trickery. :)
'Note, however that the constants for x+, etc. are spelled out for InterpretStoredScript.
'*** Constants ***
Private Const CHAR_MAXPOSSESSIONS As Integer = 50

Private Const CHAR_MAXXRANGE As Integer = 7
Private Const CHAR_MAXYRANGE As Integer = 7
Private Const CHAR_MINXRANGE As Integer = 3
Private Const CHAR_MINYRANGE As Integer = 3

Private Const SCRIPT_BYE As Integer = 6

Private Const LVW_DESC As Integer = 1
Private Const LVW_MOVEMENT As Integer = 2

' *** End Constants ***
'The player's type.
Private Type Char
    X As Integer   'this is figured JUST LIKE CellX, CellY
    Y As Integer
    Possessions(0 To CHAR_MAXPOSSESSIONS - 1) As Thing
    'Health as byte
    'exp as integer
    ':
    ':
    'Whatever else Casey wants to put in here.
End Type

'a type that will store all stored scripting commands so that they can be stored at a future time and restored when
'we change screens.
Private Type StoredScript
    ScriptType As Byte    'tells the command that should be executed on this 'Thing'
    
        'after this StoredScript mostly matches struct 'Thing'..........oops! I mean Type Thing.
    Desc As Integer
        'NOTE: if you do not use X,Y with the InterpretScriptCommand() fill them with NOT_GIVEN(just in case)
        'this is mainly because (currently the remove command) the command may be able to use the X,Y optionally
        'and won't be able to tell your empty values from values of (0,0)
    X As Integer
    Y As Integer
    Movement As Byte
    Type As Integer
    Tag As String   'this is an enhancement to the 'Tag' in the 'Thing' structure(not limited to 4 chars). The reason
        'I am providing this is this: The InterpretStoredScript function (which does the same thing as InterpretScriptLine
        ', but acts on a single command only) may soon use this to parse additional data needed. Plus, it provides a way to pass strings.
        'However, I mainly just provided it because I could :)  (i.e. there aren't scads of them stored to disk, so no need to be
        'economical.)
End Type
'Stored Script Declares
Private Const REMOVE = 1
Private Const PUTTED = 2
Private Const CHMAP = 3
Private Const CHAT = 4
Private Const CHTHREAD = 5
Private Const SLEEP = 6
Private Const GIVE = 7
Private Const TAKE = 8
Private Const WARP = 9
Private Const XPLUS = 10
Private Const XMINUS = 11
Private Const YPLUS = 12
Private Const YMINUS = 13   'woo-ooo! unlucky!

Private Const MAXSTOREDCOMM = 99
Dim StoredScriptCommands(0 To MAXSTOREDCOMM) As StoredScript

'Player position Declares
Dim You As Char
Dim ScreenX As Long
Dim ScreenY As Long
Dim TopX As Integer
Dim TopY As Integer

'Selection Declares
'Dim bBlocking As Boolean    'so that you can drag the mouse to paint
Dim wSelect As Byte 'a flag that tells us what we're doing with the 'selection' of a 'thing'
Dim SelectX As Integer, SelectY As Integer 'whats the square that the user has selected to
'work with for an object?
Private Const UNSELECTED = 0 'means that no'thing' is selected right now
Private Const SELECTED = 1 'means that the user has selected an X,Y and filled SelectX,Y with them.
Private Const DROPPING = 2
Private Const EXAMINING = 3
Private Const GETTING = 4
Private Const TALKING = 5
Private Const USING = 6
Private Const WHATSTHIS = 7
'Filename declares(more inside individual functions that open and close particular files within one function call.)
Dim Fileno As Integer
Dim ObjFileno As Integer    'this is the File of the 'Things' file: currently 65% of the
'size of the map file, but that will change if we change the structure of 'Thing'.

'Animation Declares
Dim MoveState As Integer    'this will keep track of how the player is moving
'Const STILL = 0    'Note that this is already defined inside Packrat.bas so it is redundant to define it twice
Private Const UP = 1    'all these stupid states are good preparation for animation in V2.0
Private Const DOWN = 2  'plus they're useful now for tracking the mouse.
Private Const LFT = 3
Private Const RGHT = 4
Private Const UPLEFT = 5
Private Const UPRIGHT = 6
Private Const DOWNLEFT = 7
Private Const DOWNRIGHT = 8
'Dim bToolTips As Boolean    'this is to keep track of whether the user wants tooltips or not(currently unused)

'Object declares
Dim Description() As String    'this holds all the descriptions of the different 'Things'

'Talking and storyboard declares
Private Const NUMTHREADS As Integer = 1 '(0 to 1)
Private Const NUMHEADINGS As Integer = 2 '(0 to 2)
Private Const INITIALSCRIPTSIZE = 2000 '(0 to 2000) this is 2000 right now, but could change in the future if necessary.
Dim Script() As String  'this is a string array that holds the whole of the current script
Dim Threads(0 To NUMTHREADS) As Integer  'I hope that integer is big enough
Dim ThreadLineno(0 To NUMTHREADS) As Integer 'these are bookmarks of the line numbers at which the threads
'start.

Private Sub LoadDescriptions()
    Dim DescFileno As Integer
    Dim DescFilename As String, Temp As String
    Dim i As Integer
    Dim NumLines As Long
    DescFileno = FreeFile
    DescFilename = Left(Opener, Len(Opener) - 4) & ".dsc"
    Open DescFilename For Input As #DescFileno
    Line Input #DescFileno, Temp    'get how many descriptions are in this file.
    NumLines = CLng(Temp)
    ReDim Description(1 To NumLines)
    For i = 2 To NumLines + 1
        Line Input #DescFileno, Description(i - 1)
    Next i
    Close DescFileno
End Sub
Private Sub PaintViewport()
    'here we have the edge map test code that used to be in PaintMap(explore.bas)
'Static bSaved As Boolean    'alert: just fixed the problem wherein the map got VERY slow at
'land's end. I forgot to make bSaved static and it came up as False every time.(Boy do I
'feel stupid.) **This commented because we *shouldn't* be letting the player bash into land's end. So don't let him do it!!

'First move the array and clip it to the edges.
    If TopX = (MAP_ARRAYX - MAP_SCREENX) Then
        'If ScreenX + MAP_ARRAYX < MapXSize Then    'see else for commenting reason
            TopX = MAP_SCREENX   'reset the viewport to center of array
            'SaveMap Fileno, ScreenX, ScreenY    'save changes of current position to disk
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenX = ScreenX + MAP_SCREENX  'move array over 10 cells to next pos.
            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
            LoadThings ObjFileno, ScreenX, ScreenY
'            bSaved = False
            RestoreScriptCommands
'this commented out since a completed map SHOULD NOT let the player walk to the edge of it!
'        Else    'ScreenX 30 = MapXSize and we're at map edge.
'            If bSaved = False Then
'                'For the game engine:I have removed the call to SaveMap since we aren't
'                'changing it constantly. You can, however, still call it if you drastically
'                'want to change the map on a one-time basis.
'                'SaveMap Fileno, ScreenX, ScreenY    'save the array to disk but DO NOT move the array
'                'over to next position because it would otherwise go off the edge.
'                '(or reset the viewport)
'                SaveThings ObjFileno, ScreenX, ScreenY
'                bSaved = True 'turn on a switch to make sure we don't repeatedly save to disk
'                'when moving along the edge of the map(because we don't reset position when
'                'moving along edge of map)
'                RestoreScriptCommands
'            End If
'        End If
    End If
    
    If TopY = (MAP_ARRAYY - MAP_SCREENY) Then
'        If ScreenY + MAP_ARRAYY < MapYSize Then 'we're not at the edge of the map, so business as usual
            TopY = MAP_SCREENY
            'SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenY = ScreenY + MAP_SCREENY
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
            RestoreScriptCommands
'        Else
'            If bSaved = False Then
'                'SaveMap Fileno, ScreenX, ScreenY
'                SaveThings ObjFileno, ScreenX, ScreenY
'                bSaved = True
'                RestoreScriptCommands
'            End If
'        End If
    End If
    'oops, forgot to add top, left checking(I was really tired last night)
    If TopX = 0 Then
'        If ScreenX > 0 Then
            TopX = MAP_SCREENX
'            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenX = ScreenX - MAP_SCREENX
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
'            bSaved = False
            RestoreScriptCommands
'        Else   'screenx = 0
'            If bSaved = False Then
'                'SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
'                SaveThings ObjFileno, ScreenX, ScreenY
'                bSaved = True
'                RestoreScriptCommands
'            End If
'        End If
    End If
    If TopY = 0 Then
'        If ScreenY > 0 Then
            TopY = MAP_SCREENY
'            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenY = ScreenY - MAP_SCREENY
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
'            bSaved = False
            RestoreScriptCommands
'        Else
'            If bSaved = False Then
''               SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
'                SaveThings ObjFileno, ScreenX, ScreenY
'                bSaved = True
'                RestoreScriptCommands
'            End If
'        End If
    End If
    PaintMapFast picViewport, imlTerrain, TopX, TopY
    
    'now paint the objects
    PaintThings picViewport, imlThings, TopX, TopY, ScreenX, ScreenY 'paint the 'Things' onto picViewport as
    'well.
    
    'now paint 'you' on screen.
    imlThings.ListImages("Professor").Draw picViewport.hDC, (You.X - ScreenX - TopX) * MAP_TILEXSIZE, (You.Y - ScreenY - TopY) * MAP_TILEYSIZE, imlTransparent
    
'This section of code currently unused because I have not, nor intend to currently, implemented a command button system
'of object manipulation. There are three reasons for this:
'   1.Command buttons can retain the focus and shunt KeyDown messages away from the form. This is very bad for a game.
'   2.It is way too much bother to add command buttons to an already working right-click interface.
'   3.The right-click menu selection method is by far easier anyway.
'    If wSelect = SELECTING Then 'if they're selecting, display a hilite! that is
'    'always one cell ^ 2.
'        DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
'        32, 32
'    End If
    
    If wSelect >= DROPPING Then 'if we're selecting something with the keyboard show them a cursor.
        DrawHighLight picViewport, (SelectX - TopX - ScreenX) * MAP_TILEXSIZE, (SelectY - TopY - ScreenY) * MAP_TILEYSIZE, _
        MAP_TILEXSIZE, MAP_TILEYSIZE
    End If
    
    'tell the user his position(for debugging purposes currently)
    lblPosition.Caption = "X: " & You.X & " Y: " & You.Y '& " TopX = " & TopX & " TopY = " & TopY don't need
    'TopX,Y info any more.
End Sub

Private Sub MoveViewport()
Dim iOldYouX As Integer
Dim iOldYouY As Integer
Dim iThingNum As Integer
    iOldYouX = You.X 'this so it is very easy to restore your settings if you walked through walls.
    iOldYouY = You.Y
    Select Case MoveState   'move 'you' in the appropriate direction
        Case STILL
            Exit Sub
        Case UP
            You.Y = You.Y - 1
        Case DOWN
            You.Y = You.Y + 1
        Case LFT
            You.X = You.X - 1
        Case RGHT
            You.X = You.X + 1
        Case UPLEFT
            You.X = You.X - 1
            You.Y = You.Y - 1
        Case UPRIGHT
            You.X = You.X + 1
            You.Y = You.Y - 1
        Case DOWNLEFT
            You.X = You.X - 1
            You.Y = You.Y + 1
        Case DOWNRIGHT
            You.X = You.X + 1
            You.Y = You.Y + 1
    End Select
    If imlTerrain.ListImages(Map(((You.Y - ScreenY) * MAP_ARRAYX) + (You.X - ScreenX))).Tag = "" Then   'oops, we hit a solid rock.
        You.X = iOldYouX
        You.Y = iOldYouY
        MoveState = STILL
    ElseIf IsThing(You.X, You.Y, iThingNum) Then
        'we bumped into an irate person. here we should init battles, maybe move people out of your
        'way, and maybe move objects out of your way...But for now, we'll just stop you flat in your
        'tracks
        You.X = iOldYouX
        You.Y = iOldYouY
        MoveState = STILL
    End If
        'now that we've moved and clipped you, we need to move the Things, AND clip them.
    MoveThings
    If (You.X - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
        TopX = TopX + 1                     'screen a little
        If TopX = (MAP_ARRAYX - MAP_SCREENX) + 1 Then 'we're on the edge of the map, so bounce the player back
            TopX = MAP_ARRAYX - MAP_SCREENX
            You.X = You.X - 1
        End If
    ElseIf (You.X - TopX - ScreenX) < CHAR_MINXRANGE Then
        TopX = TopX - 1
        If TopX = -1 Then
            TopX = 0
            You.X = You.X + 1
        End If
    End If
    If (You.Y - TopY - ScreenY) > CHAR_MAXYRANGE Then    'we've gone out of our range of movement, so scroll the
        TopY = TopY + 1                     'screen a little
        If TopY = (MAP_ARRAYY - MAP_SCREENY) + 1 Then 'we're on the edge of the map, so bounce the player back
            TopY = MAP_ARRAYY - MAP_SCREENY
            You.Y = You.Y - 1
        End If
    ElseIf (You.Y - TopY - ScreenY) < CHAR_MINYRANGE Then
        TopY = TopY - 1
        If TopY = -1 Then
            TopY = 0
            You.Y = You.Y + 1
        End If
    End If
    PaintViewport    'new sub that calls all Paintxxxx functions that are in the *.bas files

End Sub
Private Sub AdLabel()
        'generate a hilarious comment for the Ad Label
        'think up new comments and send them to Nathan!!
'NOTE: This sub not subject to constants as it is rather optional...
Dim Comment As Integer
    Comment = CInt(Rnd * 12) + 1 'thirteen 'ads'
    With lblAd
    .FontSize = 36
    Select Case Comment
        Case 1
            .Caption = "Your Ad Here"
        Case 2
            .Caption = "Buy Shoes -- " & vbCrLf & "Nike® Shoes"
        Case 3
            .Caption = "Feed the Birds"
        Case 4
            .Caption = "Don't Worry; Be Happy"
        Case 5
            .Caption = "Watch Star Truck. Ha. Ha."
        Case 6
            .FontSize = 24
            .Caption = "Save the Environment," & vbCrLf & " Kill All Cows," & vbCrLf & " Eat At McDonalds"
        Case 7
            .Caption = "Starring Mikey the Mouse!!"
        Case 8
            .Caption = "Eat mor chikin!™"
        Case 9
            .FontSize = 24
            .Caption = "Emus taste good, " & vbCrLf & "like poultry should."
        Case 10
            .FontSize = 24
            .Caption = "Player Beware!" & vbCrLf & "You choose the mayor!"
        Case 11
            .FontSize = 24
            .Caption = "Ding Dong." & vbCrLf & "Ho Ho." & vbCrLf & "Buy our game!"
        Case 12
            .Caption = "Eat Twinkies®. They're GOOD!"
        Case 13
            .Caption = "With Pride..." & vbCrLf & "Since 1998"
        Case Else
            MsgBox "Select Case Error in AdLabel!"
            .Caption = "Your Ad Here"
    End Select
    End With

End Sub
Sub MoveThings()
Dim i As Integer
Dim newX As Integer, newY As Integer
Dim Direction As Integer
    For i = 0 To OBJ_MAXTHINGSARRAY - 1 Step 1
            If Things(i).X > NONE Then    'make sure we've got a valid 'Thing'
            newX = Things(i).X
            newY = Things(i).Y
            Select Case Things(i).Movement
                Case RANDOM 'here we generate a random movement
                    Direction = CInt(Rnd * 8)   '0 to 8(i think)
                    'note:0 to 8 are based on the 8 movement constants used below
                    Select Case Direction
                        Case STILL
                            GoTo Continue
                        Case UP
                            newY = Things(i).Y - 1
                        Case DOWN
                            newY = Things(i).Y + 1
                        Case Left   'i hope this is the right value
                            newX = Things(i).X - 1
                        Case RGHT
                            newX = Things(i).X + 1
                        Case UPLEFT
                            newY = Things(i).Y - 1
                            newX = Things(i).X - 1
                        Case UPRIGHT
                            newY = Things(i).Y - 1
                            newX = Things(i).X + 1
                        Case DOWNLEFT
                            newY = Things(i).Y + 1
                            newX = Things(i).X - 1
                        Case DOWNRIGHT
                            newY = Things(i).Y + 1
                            newX = Things(i).X + 1
                    End Select
                Case FOLLOW 'here they try to follow you until they run into the edge of the screen.
                    If You.X < Things(i).X Then
                        newX = Things(i).X - 1
                    ElseIf You.X > Things(i).X Then
                        newX = Things(i).X + 1
                    End If
                    If You.Y < Things(i).Y Then
                        newY = Things(i).Y - 1
                    ElseIf You.Y > Things(i).Y Then
                        newY = Things(i).Y + 1
                    End If
                Case Else
                    GoTo Continue   'this hack works like the C++ statement 'continue' which
                    'VB carelessly never implemented.
            End Select
            'now clip the object
            If (newX - ScreenX) = -1 Or (newY - ScreenY) = -1 Or (newX - ScreenX) = 30 Or (newY - ScreenY) = 30 Then
                GoTo Continue 'the continue; hack
            ElseIf newX = You.X And newY = You.Y Then  'the player is already here...Don't get in his way!!
                GoTo Continue   'the continue; hack
            ElseIf imlTerrain.ListImages(Map(((newY - ScreenY) * 30) + (newX - ScreenX))).Tag = "" Then  'the 'Thing' hit a solid tile.
                GoTo Continue   'the continue; hack
            ElseIf IsThingExclude(newX, newY, i) = True Then    'you bumped into another 'Thing'
                GoTo Continue 'the continue; hack
            'now if the 'Thing' survived the clipping, then we'll move it. But , movething has inherent screen clipping, so
            'we still won't let whatever it is go off the edge of the screen it started on.
            Else
                MoveThing i, newX, newY 'clipping to the original screen is done within this
                'function
            End If
            
        End If
Continue:
    Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ArrayNum As Integer
    Select Case KeyCode
        Case vbKeyLeft
            MoveState = LFT
        Case vbKeyDown
            MoveState = DOWN
        Case vbKeyUp
            MoveState = UP
        Case vbKeyRight
            MoveState = RGHT
        Case vbKeyNumpad2
            MoveState = DOWN
        Case vbKeyNumpad4
            MoveState = LFT
        Case vbKeyNumpad6
            MoveState = RGHT
        Case vbKeyNumpad8
            MoveState = UP
        Case vbKeyEnd
            MoveState = DOWNLEFT
        Case vbKeyPageDown
            MoveState = DOWNRIGHT
        Case vbKeyHome
            MoveState = UPLEFT
        Case vbKeyPageUp
            MoveState = UPRIGHT
        Case vbKeyNumpad1
            MoveState = DOWNLEFT
        Case vbKeyNumpad3
            MoveState = DOWNRIGHT
        Case vbKeyNumpad7
            MoveState = UPLEFT
        Case vbKeyNumpad9
            MoveState = UPRIGHT
        Case vbKeyReturn, vbKeySpace    'both do the same, thing(i.e. accept changes.)
        Dim bResult As Boolean
            'check what wSelect is and act accordingly
            '(we already have set SelectX,Y with the select case wSelect below)
            bResult = IsThing(SelectX, SelectY, ArrayNum)
            Select Case wSelect
                Case DROPPING
                    If bResult = False Then mnuDrop_Click
                Case GETTING
                'why the two line construction? Well, it turns out VB again doesn't function
                'like C does... It tests BOTH sides of an If...And...Then *before* seeing
                'if an If is false. Therefore, if bResult = false, you have problems,
                'because Typeofthing crashes without a valid object to work on. Hence the
                'double If construction...
                    If bResult = True Then
                        If TypeOfThing(ArrayNum) = OBJ Then mnuGet_Click
                    End If
                Case EXAMINING
                    If bResult = True Then
                        If TypeOfThing(ArrayNum) = OBJ Then mnuExamine_Click
                    End If
                Case TALKING
                    If bResult = True Then
                        If TypeOfThing(ArrayNum) = PERSON Then mnuTalk_Click
                    End If
                Case USING
                    If bResult = True Then
                       If TypeOfThing(ArrayNum) = OBJ Then mnuUse_Click
                    End If
                Case WHATSTHIS
                    mnuWhatsThis_Click  'just call WhatsThis--no need for a person/object
            End Select
        Case vbKeyEscape    'cancel changes. or if no key has been pressed, ask if the user wants to quit.
            If wSelect > UNSELECTED Then
                wSelect = UNSELECTED
                SelectX = NONE
                SelectY = NONE
                PaintViewport
            Else    'maybe they're trying to quit
                'so unload the form in the standard way if they do, in fact, want to quit.
                If vbYes = MsgBox("Are you sure you want to quit?", vbYesNo, "Pressed ESC") Then Form_Unload 0
            End If
        Case vbKeyD 'drop
            wSelect = DROPPING
            SelectX = You.X
            SelectY = You.Y
        Case vbKeyG 'get
            wSelect = GETTING
            SelectX = You.X
            SelectY = You.Y
        Case vbKeyE, vbKeyX 'both E and X should work for EXamine
            wSelect = EXAMINING
            SelectX = You.X
            SelectY = You.Y
        Case vbKeyT 'talk
            wSelect = TALKING
            SelectX = You.X
            SelectY = You.Y
        Case vbKeyU 'use(unimplemented)
            wSelect = USING
            SelectX = You.X
            SelectY = You.Y
        Case vbKeyW, vbKeyL 'what's this should also work for What's this AND Look
            wSelect = WHATSTHIS
            SelectX = You.X
            SelectY = You.Y
    End Select
    'now find out what we should do now that we've found out which way we're moving
    Select Case wSelect
    Dim OldX As Integer, OldY As Integer
        Case UNSELECTED 'just normal movement...
            MoveViewport
        Case Is >= DROPPING 'this time we're moving the selected cursor around, not you...
            OldX = SelectX
            OldY = SelectY
            Select Case MoveState   'move the cursor in the appropriate direction
                Case UP
                    SelectY = SelectY - 1
                Case DOWN
                    SelectY = SelectY + 1
                Case LFT
                    SelectX = SelectX - 1
                Case RGHT
                    SelectX = SelectX + 1
                Case UPLEFT
                    SelectX = SelectX - 1
                    SelectY = SelectY - 1
                Case UPRIGHT
                    SelectX = SelectX + 1
                    SelectY = SelectY - 1
                Case DOWNLEFT
                    SelectX = SelectX - 1
                    SelectY = SelectY + 1
                Case DOWNRIGHT
                    SelectX = SelectX + 1
                    SelectY = SelectY + 1
            End Select

            If wSelect = GETTING Or wSelect = USING Then    'don't let the cursor get more than one space away.
                If Abs(You.X - SelectX) > 1 Or Abs(You.Y - SelectY) > 1 Then SelectX = OldX: SelectY = OldY
                'note that the colon trick is something I picked up from an old QBasic book which still used
                'some one-liner type tricks that make for unreadable code if you use them too much.
            Else    'it's another command. Don't allow the cursor off the screen.
                If (SelectX - TopX - ScreenX) > (MAP_SCREENX - 1) Or (SelectX - TopX - ScreenX) < 0 Then SelectX = OldX
                If (SelectY - TopY - ScreenY) > (MAP_SCREENY - 1) Or (SelectY - TopY - ScreenY) < 0 Then SelectY = OldY
            End If
            MoveState = STILL   'make sure we don't start suddenly moving somewhere(not too much of a problem, since
            'I'm not going to call MoveViewPort, but a stitch in time saves nine.
            PaintViewport   'thereby painting the cursor, but not moving anything.
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveState = STILL
End Sub

Private Sub Form_Load()
Dim Dummy As Integer
Dim DummyThing As Thing
    Randomize Timer
    frmGame.ScaleMode = vbPixels
  'initialize the variables(more later)
  'change these to something meaningful later.
    ScreenX = 0
    ScreenY = 0
    TopX = ScreenX + 10
    TopY = ScreenY + 10
    SelectX = NONE
    SelectY = NONE
    You.X = TopX + 4
    You.Y = TopY + 4
    Threads(0) = 0  'these are the start values of the threads for the story
    Threads(1) = 8
    'now open the map file and thing file
    Opener = App.Path & "\Hello.map"    'allow specification of these at start-up sometime...
    ObjOpener = App.Path & "\Hello.thi"
    Fileno = FreeFile
    Open Opener For Random As #Fileno Len = Len(Dummy)
    
    DummyThing.Desc = NONE
    DummyThing.Movement = 0
    DummyThing.Type = NONE
    DummyThing.X = NONE
    DummyThing.Y = NONE
    'init your possessions
Dim i As Integer
    For i = 0 To CHAR_MAXPOSSESSIONS - 1 Step 1
        You.Possessions(i) = DummyThing
    Next i
    ObjFileno = FreeFile
    Open ObjOpener For Random As #ObjFileno Len = Len(DummyThing)
    
    Get #Fileno, 1, MapXSize    'get the size of the map
    Get #Fileno, 2, MapYSize
    
    LoadMap Fileno, ScreenX, ScreenY 'call all of the load functions
    LoadThings ObjFileno, ScreenX, ScreenY
    LoadDescriptions
    'now init stored script commands array.
    For i = 0 To MAXSTOREDCOMM Step 1
        StoredScriptCommands(i).X = NONE
    Next i
    PaintViewport
    AdLabel 'generate free advertising for various copyrighted products.
    LoadTalkBox
    'these lines commented out so that the game 1. doesn't require a bmp file in the zip and 2. so that it doesn't take so
    'long to start the game...
'    SetTalkBoxFont "Times New Roman", 24
'    OpenTalkBox "5th Floor Software" & vbCrLf & "Presents:"
'    OpenTalkBox "The Adventures of You, Mikey, and Miney"
'    SetTalkBoxFont "Times New Roman", 14
'    SetTalkBoxBackGround "C:\My Documents\Visual Basic\Game Engine 1.0\green hills.bmp"
    OpenTalkBox "Your mission is to find out what's happening with the election and to stop Mega Mouse from stealing your sugar."
    OpenTalkBox "Programming: Nathan Sanders"
'    OpenTalkBox "StoryLine: Nathan Sanders"
    OpenTalkBox "Art: Rachel Sanders"
'    OpenTalkBox "Mikey as Himself"
    UnloadTalkBox
End Sub

Private Sub Form_Resize()
    'here I'm going to resize the listview to cater to people with high-res screens...
    With lvwPossessions
    Dim RGT As Integer, Bottom As Integer
        RGT = .Left + .Width
        Bottom = .Top + .Height
        RGT = (frmGame.ScaleWidth - RGT) + .Width
        If RGT > 17 Then .Width = RGT
        Bottom = (frmGame.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sometime add saving your 'stuff' here to yet ANOTHER file...(also add some
    'mechanism to save without closing the program)
    SaveMap Fileno, ScreenX, ScreenY
    SaveThings ObjFileno, ScreenX, ScreenY
    Close Fileno
    Close ObjFileno
    End 'just in case...What was that other way the Bob talked about??
    'maybe frmGame.Unload
    'maybe Set frmGame = Nothing
    'maybe...
End Sub

Private Sub lblAd_Click()
    AdLabel
End Sub

Private Sub lblAd_DblClick()
    RestoreThings ScreenX   'screeny is optional here(provided mainly so that it will look
    'just like the other functions. ScreenY is NOT needed!
End Sub



Private Sub lvwPossessions_DblClick()
    With lvwPossessions
        If .View = lvwReport Then
            .View = lvwIcon
        Else
            .View = lvwReport
        End If
    End With
End Sub

Private Sub lvwPossessions_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyU
            'they are trying to use something...should we just setup SelectX,Y, etc. and call mnuUse or do it ourself?
            'for test purposes, just call mnuUse
            wSelect = USING
            SelectX = NONE    'this is a flag to tell mnuUse that lvwPossessions is calling it rather than Form_Keydown
            SelectY = NONE
            mnuUse_Click
        Case vbKeyD
            wSelect = DROPPING
            SelectX = You.X
            SelectY = You.Y
            picViewport.SetFocus
        Case vbKeySpace
            lvwPossessions_DblClick
    End Select
End Sub

Private Sub lvwPossessions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then lvwPossessions_DblClick
End Sub

Private Sub mnuDrop_Click()
    'wSelect should be set to selected and selectx,y initialized for this function.
Dim Temp As Thing
Dim Count As Integer
    If lvwPossessions.ListItems.Count = 0 Then Exit Sub 'can't drop anything when you don't have anything to drop
    Count = CInt(Right(lvwPossessions.SelectedItem.Key, Len(lvwPossessions.SelectedItem.Key) - 1)) 'SelectedItem is set to the first one if none selected
    Temp = You.Possessions(Count)
    Temp.X = SelectX
    Temp.Y = SelectY
    
    If PutThing(Temp, ScreenX, ScreenY) = True Then
        TakeThing You, lvwPossessions, Count
    Else
        MsgBox "Screen Full! Try dropping over about 5 spaces."
    End If
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
End Sub

Private Sub mnuExamine_Click()
Dim ArrayNum As Integer
    IsThing SelectX, SelectY, ArrayNum 'this is only to get arraynum; we already know that there is something there
    If TypeOfThing(ArrayNum) = OBJ Then 'it needs to be an object before we can show a description, right?
        MsgBox Description(Things(ArrayNum).Desc), vbInformation
    End If
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
End Sub

Private Sub mnuGet_Click()
Dim Temp As Thing
Dim ArrayNum As Integer
Dim i As Integer
    i = 0   'a simple increment variable. That is initialized to 0.
    IsThing SelectX, SelectY, ArrayNum  'get what number this is
    If TypeOfThing(ArrayNum) = OBJ Then 'make sure you're trying to take objects!
        'save the value of the thing that you're taking
        Temp = Things(ArrayNum)
        'take the 'Thing' away from the map
        RemoveThingArray ArrayNum

        If GiveThing(You, lvwPossessions, Temp) = False Then
            MsgBox "Inventory full!"
            PutThing Temp, ScreenX, ScreenY 'put it back in a hurry
        End If
    End If
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
End Sub

Private Sub mnuTalk_Click()
Dim ArrayNum As Integer
Dim Intro As String
    IsThing SelectX, SelectY, ArrayNum
    LoadScript Things(ArrayNum).Desc, Intro
    RunScript Intro
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
End Sub

Private Sub mnuUse_Click()
'now that the SelectX,Y has been processed properly, we finally let mnuUse do something with it.
Dim thiTemp As Thing    'just a holding spot for the 'Thing' in question.
Dim intArray As Integer 'note this variable is used DIFFERENTLY in several places throughout the function due to its general nature and name. Warning!!
    'this function needs wSelect and SelectX,Y to be set to work
    If SelectX = NONE Then 'we know the player is using something in his possessions.
        'therefore get the index, and then the object from his Possessions array.
        intArray = CInt(Right(lvwPossessions.SelectedItem.Key, Len(lvwPossessions.SelectedItem.Key) - 1)) 'SelectedItem is set to the first one if none selected
        thiTemp = You.Possessions(intArray)
    ElseIf SelectX > NONE Then    'then he's using something on the map.
        'so get its data into thiTemp
        IsThing SelectX, SelectY, intArray
        thiTemp = Things(intArray)
    'Else oops! this is a game, and we don't include elses; we just let it crash and then debug it...it's *faster* that way!
    End If
    'OK, now we have to figure out which Type this is and do something accordingly
    Select Case thiTemp.Desc
        Case 1
            MsgBox "You can't find the needle no matter how much you try! Try again later."
        Case 4
            MsgBox "You look a fool grubbing around in the pot of the potted palm, but you hope that you'll find something. Unfortunately, you're not so lucky."
        Case 6
            MsgBox "As you fiddle with the table, it solidifies temporarily...Wow!"
        Case 7
            MsgBox "You look a fool grubbing around in the pot of the potted palm, but you hope that you'll find something. Unfortunately, you're not so lucky."
        Case 8
            MsgBox "The iron pot simply refuses to stop following you!"
        Case 9
            MsgBox "You see 3 K-rangs inside the pot. You take them from inside the pot and discreetly 'pocket' them....Maybe you shouldn't have done that."
            Dim thiGivee As Thing
            thiGivee.Desc = 1
            thiGivee.Movement = STILL
            thiGivee.Type = 10 'a "brick"
            thiGivee.X = 0
            thiGivee.Y = 0
            For intArray = 0 To 2 Step 1
                If GiveThing(You, Me.lvwPossessions, thiGivee) = False Then
                    MsgBox "You don't have enough room to hold them, however."
                    Exit For
                End If
            Next
        Case 10
            'put a talkbox here with two choices: "Mikey" and "Macky" (until we find a better name for him)
            'then set the thread to a certain number to change the storyline.
            'facxila, cxu ne?
        Case 11
            'nothing here yet either, but we'd need more talkboxes and then a select case and a havething(for points) and then a givething
            'and then a takething(for points)(or whatever we decide to make Moneys).
        Case 12
            'nothing here yet either, but we'd need more talkboxes and then a select case and a havething(for the obj) and then a givething(for points)
            'and then a takething(for the obj)
        Case 13, 14
            'K-rang, brick: here we need to run through all the people and monsters on the screen and then give the user a 'Select Target' TalkBox
            'Then we would DeleteThing(I think that's what it's called) whatever they chose. Or maybe we could do something different, but DeleteThing sounds
            'like the thing that came off the top of my head. Also TakeThing or DeleteThing the brick.
        Case 2 To 3, 15 To 16 'no action yet!
            'and also I'm using MsgBox instead of TalkBox; I'll switch later :)
            MsgBox "No matter what you try to do to it, the object refuses to co-operate"
    End Select
    wSelect = UNSELECTED    'I love VB's autocaps :) (I have been working with DevStudio lately, or couldn't
    SelectX = NONE  'you tell?
    SelectY = NONE
End Sub

Private Sub mnuWhatsThis_Click()
Dim ArrayNum As Integer, Result As Integer
Dim msg As String
    msg = "You see"
    If IsThing(SelectX, SelectY, ArrayNum) Then
        Result = TypeOfThing(ArrayNum)
        If Result <> PERSON Then
            msg = msg & " a"
        End If
        msg = msg & " " & imlThings.ListImages(Things(ArrayNum).Type).Key
        If Result = OBJ Then
            msg = msg & " on"
        Else
            msg = msg & " standing on"
        End If
    End If
    If imlTerrain.ListImages(Map(((SelectY - ScreenY) * MAP_ARRAYX) + (SelectX - ScreenX))).Tag = "" Then
        msg = msg & " a"
    Else
        msg = msg & " a patch of"
    End If
    msg = msg & " " & imlTerrain.ListImages(Map(((SelectY - ScreenY) * MAP_ARRAYX) + (SelectX - ScreenX))).Key & "."
    MsgBox msg, vbInformation
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
End Sub

Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XCell As Integer, YCell As Integer
Dim ArrayNum As Integer
Dim bResult As Boolean
    XCell = X \ MAP_TILEXSIZE
    YCell = Y \ MAP_TILEYSIZE
    Select Case Button
        Case vbKeyLButton   'check to see if we're selecting something. If not, ignore the click.
            If wSelect >= DROPPING Then
                bResult = IsThing(SelectX, SelectY, ArrayNum)
                Select Case wSelect
                    Case GETTING
                        If bResult = True And Abs(You.X - SelectX) < 2 And Abs(You.Y - SelectY) < 2 Then
                            If TypeOfThing(ArrayNum) = OBJ Then mnuGet_Click
                        End If
                    Case TALKING
                        If bResult = True Then
                            If TypeOfThing(ArrayNum) = PERSON Then mnuTalk_Click
                        End If
                    Case USING
                        If bResult = True And Abs(You.X - SelectX) < 2 And Abs(You.Y - SelectY) < 2 Then
                            If TypeOfThing(ArrayNum) = OBJ Then mnuUse_Click
                        End If
                    Case DROPPING
                        If bResult = False Then mnuDrop_Click
                    Case EXAMINING
                        If bResult = True Then
                            If TypeOfThing(ArrayNum) = OBJ Then mnuExamine_Click
                        End If
                    Case WHATSTHIS
                        If bResult = True Then mnuWhatsThis_Click
                End Select
                PaintViewport
            End If
        Case Else   'usually right button
        Dim Result As Integer
            wSelect = SELECTED  'set select status so that when we call the context menu
            SelectX = XCell + ScreenX + TopX    'functions they'll know what the player
            SelectY = YCell + ScreenY + TopY    'is pointing at.

            Result = IsThing(SelectX, SelectY, ArrayNum)
            'set or reset all the menu values(the If statement is to determine if the space is empty
            mnuSep1.Visible = False
            mnuTalk.Visible = False
            mnuGet.Visible = False
            If lvwPossessions.ListItems.Count = 0 Or Result = True Or imlTerrain.ListImages(Map(((SelectY - ScreenY) * MAP_ARRAYX + (SelectX - ScreenX)))).Tag = "" Then
                mnuDrop.Visible = False
            Else
                mnuDrop.Visible = True
                mnuSep1.Visible = True
            End If
            mnuUse.Visible = False
            mnuExamine.Visible = False
            mnuWhatsThis.Visible = True
            If Result = True Then   'see IsThing call above
                If TypeOfThing(ArrayNum) = PERSON Then
                    mnuTalk.Visible = True
                    mnuGet.Visible = False
                    mnuUse.Visible = False
                    mnuExamine.Visible = False
                    mnuSep1.Visible = True
                ElseIf TypeOfThing(ArrayNum) = MONSTER Then
                    mnuTalk.Visible = False
                    mnuGet.Visible = False
                    mnuUse.Visible = False
                    mnuExamine.Visible = False
                Else    'we hope OBJ
                    mnuTalk.Visible = False
                    'make sure you're close enough to something to pick it up or use it.
                    'Otherwise leave it invisible.
                    If Abs(You.X - SelectX) < 2 And Abs(You.Y - SelectY) < 2 Then
                        mnuGet.Visible = True
                        mnuUse.Visible = True
                    End If
                    mnuExamine.Visible = True
                    mnuSep1.Visible = True
                End If
            End If
            PopupMenu mnuContext
            PaintViewport
        End Select
End Sub

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MoveStateX As Integer, MoveStateY As Integer
    If Button = vbKeyLButton And wSelect = UNSELECTED Then   'We're dragging. But are we within range?
        If X < (You.X - ScreenX - TopX) * MAP_TILEXSIZE Then 'we're dragging left
            MoveStateX = Left
        ElseIf X > (You.X - ScreenX - TopX) * MAP_TILEXSIZE + MAP_TILEXSIZE Then    'we're dragging right
            MoveStateX = RGHT
        End If
        If Y < (You.Y - ScreenY - TopY) * MAP_TILEYSIZE Then 'we're dragging up
            MoveStateY = UP
        ElseIf Y > (You.Y - ScreenY - TopY) * MAP_TILEYSIZE + MAP_TILEYSIZE Then 'we're dragging down
            MoveStateY = DOWN
        End If
        'now process the movestatex,y and combine them into one variable:MoveState
        If MoveStateX = Left Then
            If MoveStateY = UP Then
                MoveState = UPLEFT
            ElseIf MoveStateY = DOWN Then
                MoveState = DOWNLEFT
            Else
                MoveState = Left
            End If
        ElseIf MoveStateX = RGHT Then
            If MoveStateY = UP Then
                MoveState = UPRIGHT
            ElseIf MoveStateY = DOWN Then
                MoveState = DOWNRIGHT
            Else
                MoveState = RGHT
            End If
        Else        'no x movement, so we must just be doing simple Y movement, so just set movestate to movestatey
            MoveState = MoveStateY
        End If
        tmrMove.Enabled = True
    Else    'set the movestate to STILL so we'll stop moving
        If wSelect >= DROPPING Then
            SelectX = ScreenX + TopX + (X \ MAP_TILEXSIZE)
            SelectY = ScreenY + TopY + (Y \ MAP_TILEYSIZE)
            PaintViewport
        Else
            MoveState = STILL
            tmrMove.Enabled = False
        End If
    End If
End Sub

Private Sub picViewport_Paint()
    PaintViewport
End Sub

Private Sub tmrMove_Timer()
    MoveViewport
End Sub
Private Function LoadScript(ScriptNum As Integer, ByRef Intro As String) As Boolean
'***debugged***(so far...)
'this function loads a script in from a PeopleScript file(*.scr). You must pass it the number
'of the script to load(usually stored in a Thing.Desc variable of a Person). It returns False
'if it cannot find the script number in the jump table. This means that you MUST include a
'(complete) jumptable in your script files. The reason that I am using them is to speed up the read of
'the file(so I don't have to search for each 'Script', then check each 'Script' to see if the number
'is the script number.
Dim Temp As String
Dim ScriptOpen As String    'the filename of the script
Dim ScriptFileNo As Integer 'the filenumber of the script
Dim Count As Integer        'in this function, mainly just a incremented variable
Dim LineNumber As Long  'this is the line number where our script starts
Dim Found As Boolean    'this is a flag that tells whether or not the jump was found inside the jumptable
Dim intStart As Integer 'the start positions of parameters for scripting statements.

    'init the counter variable to the start of the file(note that that is 1, not 0)
    Count = 1
    'open the file
    ScriptOpen = Left(Opener, Len(Opener) - 3) & "scr"
    ScriptFileNo = FreeFile
    Open ScriptOpen For Input As ScriptFileNo
    'now loop through the jumptable until we find our 'script'.
    'Note:the jumptable must have NO lines ahead of it, and must be solid jumps until it reaches the end of the jumptable
    'this means no comments or anything at the beginning before the jumptable.
    Do
        Line Input #ScriptFileNo, Temp
        intStart = InStr(Temp, ",")
        If (CInt(Mid$(Temp, 2, intStart - 1))) = ScriptNum Then    'j0, 100
            Found = True            '0 means the script number, 100 means the line number
            Exit Do                 'for the script.
        End If
        Count = Count + 1
    Loop While Left(Temp, 1) = "j"  'until we reach the end of the jump table
    
    If Found = False Then   'not in jump table, bad script-writer
        LoadScript = False
        Exit Function   'so we bail out.
    End If
    'otherwise, keep going
    
    'init the counter variable
    LineNumber = CInt(Trim$(Right(Temp, Len(Temp) - (intStart))))   'this is the number given to us by the jumptable.
    For Count = Count To (LineNumber - 1) Step 1
        Line Input #ScriptFileNo, Temp
    Next Count
    
    intStart = InStr(Temp, ",")
    'make sure we get the intro to display to the user when we call opentalkbox the first time
    Intro = Right(Temp, Len(Temp) - (intStart + 1))  'Script:0, Hello!
    
    Count = 0   're-init count variable to 0 relative to the script array instead of the
    'script file.
    ReDim Preserve Script(0 To INITIALSCRIPTSIZE) As String   '2001 is the current max number of lines per script.
    'we can change this if need be
    Do
        Line Input #ScriptFileNo, Script(Count)
'        Debug.Print Script(Count)
        Count = Count + 1
    Loop Until Trim$(Script(Count - 1)) = "}"
    ReDim Preserve Script(0 To (Count - 1)) As String  'now cut Script down to size so that we aren't
    'wasting memory.
    
    'we're done!
    LoadScript = True   'and we were successful at Loading the Script.
End Function

Private Sub RunScript(Intro As String)
'this sub actually runs the script and is designed to be called right after LoadScript. It
'actually handles all calls to InterpretScriptLine as well, but I am thinking of including a
'function called InterpretScript which would behave exactly as InterpretScriptLine with the
'exception that it would not use the Script() array but rather a string passed to it.
Dim Count As Integer    'this is the linecount with which we keep track of how complete the
'script is.
Dim i As Integer 'this is a simple counter for use in loops unrelated to updating the line
'position of the script.
Dim intStart As Integer, intEnd As Integer
Dim HeadLineno(0 To ((NUMTHREADS + 1) * (NUMHEADINGS + 1))) As Integer
'these are bookmarks of the headings(i.e. command buttons
'on opentalkbox)    'there are NUMHEADINGS + 1 per thread. If we have 6 current threads,
'then that means we have 18 of them.
Dim HeadText(0 To ((NUMTHREADS + 1) * (NUMHEADINGS + 1))) As String
    Count = 0
    Do Until Trim$(Script(Count)) = "}"    'this loop gets all of the thread line numbers.
    
        'Script:0, You see a stupid looking mouse.
        '.
        '.
        '.
        '}
        intStart = InStr(Script(Count), ":")
        If intStart = 0 Then GoTo ThreadContinue  'skip to the next iteration
        If Trim$(LCase$(Mid$(Script(Count), 1, intStart - 1))) = "thread" Then     'beginning of a thread.
            intEnd = InStr(Script(Count), "-")
            If intEnd <> 0 Then  'the scripter is specifying a range of
            'threads 'Thread:0-999 instead of just Thread:0
                For i = 0 To NUMTHREADS Step 1
                    If Threads(i) >= CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - (intStart + 1)))) And Threads(i) <= CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intEnd)))) Then
                    'we have to be >= than the first number and <= than the second.
                        ThreadLineno(i) = Count 'a match!!
                    End If
                Next i
            Else    'the thread structure doesn't specify a range, just a single one...
                For i = 0 To NUMTHREADS Step 1
                    If Threads(i) = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart)))) Then
                        ThreadLineno(i) = Count 'a match!!
                    End If
                Next i
            End If
        End If
ThreadContinue:  'a label to simulate C's continue statement
        Count = Count + 1   'make sure we increment the awful thing.
    Loop
    'OK: now we have the line #'s of all the threads.
    
Dim j As Integer    'another counter since we're already using i. This one keeps track
    j = 0           'of how many headings we've found.
    For i = 0 To NUMTHREADS Step 1   'loop through all the threads to find all the headings
        'reset count to the start point of each thread.
        Count = ThreadLineno(i)
        Do Until Trim$(Script(Count)) = "]"    'end of thread identifier
            intStart = InStr(Script(Count), ":")
            If intStart = 0 Then GoTo HeadingContinue   'a comment or sub-structural bracket
            If LTrim(LCase$(Left$(Script(Count), intStart - 1))) = "heading" Then    'beginning of heading; we must save it and
            'its text...    (Heading:Miney the Mouse)
                HeadLineno(j) = Count   'set the bookmark
                HeadText(j) = Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart)))  'get the heading
                'to show the user
                j = j + 1 'long hand for j++®   'and remember to tell the loop that we've
                'found another Heading.
            End If
HeadingContinue:
            Count = Count + 1
        Loop
    Next i  'OK: we've found all the headings...
    'now what we want to do is loop endlessly until the user gets bored and clicks 'Bye'.
    'we do this by first calling OpenTalkBox, then processing the result and interpreting
    'the appropriate heading. (or quitting)
Dim intResult As Integer
    LoadTalkBox 'load talk box and associated data.
    Do
        intResult = OpenTalkBox(Intro, HeadText(0), HeadText(1), HeadText(2), _
        HeadText(3), HeadText(4), HeadText(5), "Bye")
        
        If intResult = SCRIPT_BYE Then 'he pressed 'Bye' so we can quit.
            Exit Sub
        End If
        'now set Count to the correct heading line #
        Count = HeadLineno(intResult)
        Do Until Trim$(Script(Count)) = ")"    'and loop until we're done with this heading
            If InterpretScriptLine(Count) = False Then  'this means that we should quit
                intResult = SCRIPT_BYE 'usually results from the (End) command inside the script.
                Exit Do
            End If
        Loop
    Loop Until intResult = SCRIPT_BYE
    'so now unload the talk box.
    UnloadTalkBox
End Sub
Private Function InterpretScriptLine(ByRef Count As Integer) As Boolean
'this function actually interprets and executes each line of the Script as it is passed to it.
'Script() is the array with the script in it, and count is the value that tells us where
'we are in reading the current script. This is because this function is recursive and calls
'itself for interpreting Question and Have instructions. Every time this function
'interprets a value it ups the count by one
'If the return value is False, it means that the script should stop executing. This is
'usually because the script has the End command in it.
Dim OldThread As Integer    'this keeps track of the old thread which is changed in the 'Thread'
'command
Dim NewThread As Integer    'do you really want me to tell you about this??
Dim i As Integer    'a simple counter
Dim intStart As Integer, intEnd As Integer
Dim ArrayNum As Integer
Dim intX As Integer, intY As Integer
Dim Temp As Thing   'a temporary thing to store stuff in before modifying it.
Dim intObjectScreen As Integer, intDesc As Integer
Dim bFound As Boolean
Dim Result As Integer
Dim intType As Integer   'what the type of map is that we're supposed to change to in 'chmap'
Dim sngEndTime As Single
Dim StoreScript As StoredScript
    intStart = InStr(Script(Count), ":")
    If intStart <> 0 Then   'else, it's just a comment or some other kind of bad command, so skip it.
    
    Select Case Trim$(LCase$(Left$(Script(Count), intStart - 1))) 'figure out what command the all-powerful scripter
    'is commanding us to carry out.
        Case "chat"    'a simple chatty OpenTalkBox
            '(Chat:Hello, my name is Mikey and I want you to vote for ME!)
            OpenTalkBox Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart)))
'these are the movie functions; they will make our game able to have our characters pace around, walk away, jump on tables, etc. when you talk to them.
'now our game will be like(Oh, no!) Final Fantasy if we let it!
'Warning on all relative movement commands: if the move goes off the map() array, it will be stored using StoreScriptCommand. That array has only 99
'elements currently, so it CAN become full. Therefore, be sparing if at all possible. Just move out of sight of the player, then 'warp' to where you want to go.
'params: right now I am contemplating this format:
' x+: [repeatnum], [sleeptime], [objdesc]
'repeatnum is the number of times to repeat the particular command. specify 1 for default action(when affecting another person besides currently talked-to one)
'sleeptime is the time in ms to sleep between the repeated commands. specify 0 for default action
'objdesc is the description number of an person that you specify so you can move other characters besides the one that the player is talking to.
'all params are optional. By default they execute once upon the person the player is talking to.
        Case "x+"
        'programmers' note: intType here indicates the number of params: 0 = 0(surprise), 1 = 1, um... I don't think I need to continue with this... '='
            intEnd = Len(Script(Count)) - intStart  'see if there IS a first param.
            If intEnd = 0 Then  'no params! just move once.
                intType = 0
            Else    'there are at least one params; check again.
                intEnd = InStr(Script(Count), ",")  'get the first parameter's end comma
                If intEnd = 0 Then  'oops! only one param.
                    intType = 1
                    'now parse the param.
                    intX = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))    'here, intX means the number of times to move.
                Else    'uh-oh! two params! Parse, then check again.
                    intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
                    intStart = intEnd
                    intEnd = InStr(intStart + 1, Script(Count), ",") 'get the second parameter's end comma
                    If intEnd = 0 Then  'only two params
                        intType = 2
                        intY = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart))) 'here, intY means the length of time in ms to sleep.
                    Else    'all three! wow!
                        intType = 3
                        intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))    'here, intY means the length of time in ms to sleep.
                        intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))   'here intDesc is the Desc of the other Thing to move.
                    End If
                End If
            End If
            'now figure out what to do according to the param level
            Select Case intType
                Case 0
                    'just move the Person and bounds check against the Map array, saving if need be, then PaintViewport; DON'T forget to PiantViePort
                    'it's code will change as the type does.
                    'first get person using Isthing passed selectx, selecty
                    IsThing SelectX, SelectY, i
                    If MoveThingTo(i, 1) = False Then
                        'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                        
                        'first save it to a Temp thing
                        Temp = Things(i)
                        Temp.X = Temp.X + 1 'CHANGEIT
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = PUTTED
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.X = Temp.X
                            StoreScript.Y = Temp.Y
                            StoreScriptCommand StoreScript
                            InterpretScriptLine = False
                            Exit Function
                        ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                            MsgBox "Screen Full. Person that was walking is dead."
                        End If
                    End If
                    SelectX = SelectX + 1   'CHANGEIT
                    'finally, update the screen.
                    PaintViewport
                Case 1  'warning: hereafter, only the new stuff different from the original will be commented: be warned!!!
                    'do a for loop of the previous, only painting the viewport at the end.
                    IsThing SelectX, SelectY, i
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, 1) = False Then
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.X = Temp.X + 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX + 1, SelectY, i
                        End If
                        SelectX = SelectX + 1

                    Next intStart
                    'finally, update the screen.
                    PaintViewport

                Case 2
                    'do a for loop of the directly above, painting in the loop and sleeping in the loop
                    IsThing SelectX, SelectY, i
                    sngEndTime = Timer + (intY / 1000)
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, 1) = False Then
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.X = Temp.X + 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX + 1, SelectY, i
                        End If
                        SelectX = SelectX + 1
                        'update the screen every time now.
                        PaintViewport

                        'now sleep:
                        Do Until Timer > sngEndTime
                            DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                        Loop

                    Next intStart
                    sngEndTime = Timer + (intY / 1000)
                    Do Until Timer > sngEndTime
'                        DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                    Loop
                Case 3  'I realize this section has  lot of repeated code, etc. But it might get cleaned up eventually.
                    If intDesc = NONE Then  'we're moving You
                    
                    '*** Warning! Possible bug! It might be that when we move the arrays, the person talked to would move off the arrays. Then some of the script
                    'commands might not work. however, I will not deal with that just yet.
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            'warning! we need to change the movething If to an If that tests if we went out of range of our little movement range(3-8 or something)
                            'and scroll the screen. Then we need an if that sees If we're crossing a screen boundary and that Loads/Saves Things() and Map().
                            You.X = You.X + 1 '*** CHANGEIT
                            If (You.X - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
                                    TopX = TopX + 1                     'screen a little
                            End If
                            'finally, update the screen.
                            PaintViewport
                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                'this conversation is over
                                SelectX = NONE
                                SelectY = NONE
                                InterpretScriptLine = False
                                Exit Function
                            End If
                            
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                                You.X = You.X + 1
                                If (You.X - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopX = TopX + 1                     'screen a little
                                        If TopX = (MAP_ARRAYX - MAP_SCREENX) + 1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                                        'Things() arrays.
                                            TopX = MAP_ARRAYX - MAP_SCREENX
                                            You.X = You.X - 1 'end ***
                                            PaintViewport   'this so the function will scroll the Map() and Things() arrays.
                                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                                'this conversation is over
                                                SelectX = NONE
                                                SelectY = NONE
                                                InterpretScriptLine = False
                                                Exit Function
                                            End If
                                            
                                        End If
                                End If
                            Next intStart
                            'finally, update the screen.
                            PaintViewport
                            
                        Else    'full move sequence.
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                                You.X = You.X + 1   '***
                                If (You.X - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopX = TopX + 1                     'screen a little
                                End If
                                'update the screen every time now.
                                PaintViewport
                                If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                    'this conversation is over
                                    SelectX = NONE
                                    SelectY = NONE
                                    InterpretScriptLine = False
                                    Exit Function
                                End If
                                
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop

                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                        
                    Else    'we've got to find a person of that Description in the vicinity...
                        'note that once again, I've re-used some int variable names...
                        intStart = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
                        'IsThingDesc.
                        intEnd = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
                        i = IsThingDesc(intDesc, intStart, intEnd)
                        If i = NONE Then MsgBox "No person found! Check your script code!": GoTo Continue
                        'WARNING!!!! MUST find a way to figure out if the person that we have found is the original person talked to(i.e. the one with SelectX,Y
                        'pointing to him) and gracefully(i.e. not a LOT of extra ifs scattered through my code) handle it. I think the best way is to call InterpretScriptLine
                        'and tell it to use the standard style of x+, etc. and then continue ASAP from this section of code!! Anyway, it will hopefully happen very little,
                        'because with a real game, there will be a great diversity of 'Desc' values--almost nothing will have the same--except maybe a troop of soldiers
                        'or something, and they would probably have a different Desc for the person who does most of the talking.
                        
                        'now just paste in the code from above
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            If MoveThingTo(i, 1) = False Then
                                'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                
                                'first save it to a Temp thing
                                Temp = Things(i)
                                Temp.X = Temp.X + 1
                                'next delete it from the current screen.
                                RemoveThingArray i
                                'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = PUTTED
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.X = Temp.X
                                    StoreScript.Y = Temp.Y
                                    StoreScriptCommand StoreScript
                                ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                    MsgBox "Screen Full. Person that was walking is dead."
                                End If
                            End If
                            'finally, update the screen.
                            PaintViewport
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                    
                                If MoveThingTo(i, 1) = False Then
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.X = Temp.X + 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X + 1, Temp.Y, i
                                End If

                            Next intStart
                            'finally, update the screen.
                            PaintViewport

                        Else    'full move sequence.(case 2)
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                            
                                If MoveThingTo(i, 1) = False Then
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.X = Temp.X + 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X + 1, Temp.Y, i
                                End If
                                'update the screen every time now.
                                PaintViewport
                                'now sleep:
                                sngEndTime = Timer + (intY / 1000)
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop
        
                            Next intStart

                        End If
                    End If
                    
            End Select
        Case "x-"
            intEnd = Len(Script(Count)) - intStart  'see if there IS a first param.
            If intEnd = 0 Then  'no params! just move once.
                intType = 0
            Else    'there are at least one params; check again.
                intEnd = InStr(Script(Count), ",")  'get the first parameter's end comma
                If intEnd = 0 Then  'oops! only one param.
                    intType = 1
                    'now parse the param.
                    intX = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))    'here, intX means the number of times to move.
                Else    'uh-oh! two params! Parse, then check again.
                    intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
                    intStart = intEnd
                    intEnd = InStr(intStart + 1, Script(Count), ",") 'get the second parameter's end comma
                    If intEnd = 0 Then  'only two params
                        intType = 2
                        intY = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart))) 'here, intY means the length of time in ms to sleep.
                    Else    'all three! wow!
                        intType = 3
                        intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))    'here, intY means the length of time in ms to sleep.
                        intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))   'here intDesc is the Desc of the other Thing to move.
                    End If
                End If
            End If
            'now figure out what to do according to the param level
            Select Case intType
                Case 0
                    'just move the Person and bounds check against the Map array, saving if need be, then PaintViewport; DON'T forget to PiantViePort
                    'it's code will change as the type does.
                    'first get person using Isthing passed selectx, selecty
                    IsThing SelectX, SelectY, i
                    If MoveThingTo(i, -1) = False Then
                        'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                        
                        'first save it to a Temp thing
                        Temp = Things(i)
                        Temp.X = Temp.X - 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = PUTTED
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.X = Temp.X
                            StoreScript.Y = Temp.Y
                            StoreScriptCommand StoreScript
                            InterpretScriptLine = False
                            Exit Function
                        ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                            MsgBox "Screen Full. Person that was walking is dead."
                        End If
                    End If
                    SelectX = SelectX - 1
                    'finally, update the screen.
                    PaintViewport
                Case 1  'warning: hereafter, only the new stuff different from the original will be commented: be warned!!!
                    'do a for loop of the previous, only painting the viewport at the end.
                    IsThing SelectX, SelectY, i
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, -1) = False Then
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.X = Temp.X - 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX + 1, SelectY, i
                        End If
                        SelectX = SelectX - 1

                    Next intStart
                    'finally, update the screen.
                    PaintViewport

                Case 2
                    'do a for loop of the directly above, painting in the loop and sleeping in the loop
                    IsThing SelectX, SelectY, i
                    sngEndTime = Timer + (intY / 1000)
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, -1) = False Then
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.X = Temp.X - 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX + 1, SelectY, i
                        End If
                        SelectX = SelectX - 1
                        'update the screen every time now.
                        PaintViewport
                        'now sleep:
                        Do Until Timer > sngEndTime
                            'DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                        Loop

                    Next intStart
                    sngEndTime = Timer + (intY / 1000)
                    Do Until Timer > sngEndTime
                        DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                    Loop

                Case 3  'I realize this section has  lot of repeated code, etc. But it might get cleaned up eventually.
                    If intDesc = NONE Then  'we're moving You
                    
                    '*** Warning! Possible bug! It might be that when we move the arrays, the person talked to would move off the arrays. Then some of the script
                    'commands might not work. however, I will not deal with that just yet.
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            'warning! we need to change the movething If to an If that tests if we went out of range of our little movement range(3-8 or something)
                            'and scroll the screen. Then we need an if that sees If we're crossing a screen boundary and that Loads/Saves Things() and Map().
                            You.X = You.X - 1
                            If (You.X - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement, so scroll the
                                    TopX = TopX - 1                     'screen a little
                            End If
                            'finally, update the screen.
                            PaintViewport
                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                'this conversation is over
                                SelectX = NONE
                                SelectY = NONE
                                InterpretScriptLine = False
                                Exit Function
                            End If
                            
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                                You.X = You.X - 1
                                If (You.X - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopX = TopX - 1                     'screen a little
                                        If TopX = -1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                                        'Things() arrays.
                                            TopX = 0
                                            You.X = You.X + 1
                                            PaintViewport   'this so the function will scroll the Map() and Things() arrays.
                                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                                'this conversation is over
                                                SelectX = NONE
                                                SelectY = NONE
                                                InterpretScriptLine = False
                                                Exit Function
                                            End If
                                            
                                        End If
                                End If
                            Next intStart
                            'finally, update the screen.
                            PaintViewport
                            
                        Else    'full move sequence.
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                                You.X = You.X - 1
                                If (You.X - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopX = TopX - 1                     'screen a little
                                End If
                                'update the screen every time now.
                                PaintViewport
                                If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                    'this conversation is over
                                    SelectX = NONE
                                    SelectY = NONE
                                    InterpretScriptLine = False
                                    Exit Function
                                End If
                                
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop

                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                        
                    Else    'we've got to find a person of that Description in the vicinity...
                        'note that once again, I've re-used some int variable names...
                        intStart = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
                        'IsThingDesc.
                        intEnd = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
                        i = IsThingDesc(intDesc, intStart, intEnd)
                        If i = NONE Then MsgBox "No person found! Check your script code!": GoTo Continue
                        'now just paste in the code from above
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            If MoveThingTo(i, -1) = False Then
                                'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                
                                'first save it to a Temp thing
                                Temp = Things(i)
                                Temp.X = Temp.X - 1
                                'next delete it from the current screen.
                                RemoveThingArray i
                                'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = PUTTED
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.X = Temp.X
                                    StoreScript.Y = Temp.Y
                                    StoreScriptCommand StoreScript
                                ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                    MsgBox "Screen Full. Person that was walking is dead."
                                End If
                            End If
                            'finally, update the screen.
                            PaintViewport
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                    
                                If MoveThingTo(i, -1) = False Then
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.X = Temp.X - 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X - 1, Temp.Y, i
                                End If

                            Next intStart
                            'finally, update the screen.
                            PaintViewport

                        Else    'full move sequence.(case 2)
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                            
                                If MoveThingTo(i, -1) = False Then
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.X = Temp.X - 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X - 1, Temp.Y, i
                                End If
                                'update the screen every time now.
                                PaintViewport
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop
        
                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                    End If
                    
            End Select
        
        Case "y+"           '***still gotta convert the cut&pasted 'y' commands.****
        'programmers' note: intType here indicates the number of params: 0 = 0(surprise), 1 = 1, um... I don't think I need to continue with this... '='
            intEnd = Len(Script(Count)) - intStart  'see if there IS a first param.
            If intEnd = 0 Then  'no params! just move once.
                intType = 0
            Else    'there are at least one params; check again.
                intEnd = InStr(Script(Count), ",")  'get the first parameter's end comma
                If intEnd = 0 Then  'oops! only one param.
                    intType = 1
                    'now parse the param.
                    intX = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))    'here, intX means the number of times to move.
                Else    'uh-oh! two params! Parse, then check again.
                    intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
                    intStart = intEnd
                    intEnd = InStr(intStart + 1, Script(Count), ",") 'get the second parameter's end comma
                    If intEnd = 0 Then  'only two params
                        intType = 2
                        intY = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart))) 'here, intY means the length of time in ms to sleep.
                    Else    'all three! wow!
                        intType = 3
                        intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))    'here, intY means the length of time in ms to sleep.
                        intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))   'here intDesc is the Desc of the other Thing to move.
                    End If
                End If
            End If
            'now figure out what to do according to the param level
            Select Case intType
                Case 0
                    'just move the Person and bounds check against the Map array, saving if need be, then PaintViewport; DON'T forget to PiantViePort
                    'it's code will change as the type does.
                    'first get person using Isthing passed selectx, selecty
                    IsThing SelectX, SelectY, i
                    If MoveThingTo(i, , 1) = False Then
                        'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                        
                        'first save it to a Temp thing
                        Temp = Things(i)
                        Temp.Y = Temp.Y + 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = PUTTED
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.X = Temp.X
                            StoreScript.Y = Temp.Y
                            StoreScriptCommand StoreScript
                            InterpretScriptLine = False
                            Exit Function
                        ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                            MsgBox "Screen Full. Person that was walking is dead."
                        End If
                    End If
                    SelectY = SelectY + 1
                    'finally, update the screen.
                    PaintViewport
                Case 1  'warning: hereafter, only the new stuff different from the original will be commented: be warned!!!
                    'do a for loop of the previous, only painting the viewport at the end.
                    IsThing SelectX, SelectY, i
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, , 1) = False Then
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.Y = Temp.Y + 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX, SelectY + 1, i
                        End If
                        SelectY = SelectY + 1

                    Next intStart
                    'finally, update the screen.
                    PaintViewport

                Case 2
                    'do a for loop of the directly above, painting in the loop and sleeping in the loop
                    IsThing SelectX, SelectY, i
                    sngEndTime = Timer + (intY / 1000)
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, , 1) = False Then
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.Y = Temp.Y + 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX, SelectY + 1, i
                        End If
                        SelectY = SelectY + 1
                        'update the screen every time now.
                        PaintViewport
                        'now sleep:
                        Do Until Timer > sngEndTime
                            DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                        Loop

                    Next intStart
                    sngEndTime = Timer + (intY / 1000)
                    Do Until Timer > sngEndTime
                        DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                    Loop

                Case 3  'I realize this section has  lot of repeated code, etc. But it might get cleaned up eventually.
                    If intDesc = NONE Then  'we're moving You
                    
                    '*** Warning! Possible bug! It might be that when we move the arrays, the person talked to would move off the arrays. Then some of the script
                    'commands might not work. however, I will not deal with that just yet.
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            'warning! we need to change the movething If to an If that tests if we went out of range of our little movement range(3-8 or something)
                            'and scroll the screen. Then we need an if that sees If we're crossing a screen boundary and that Loads/Saves Things() and Map().
                            You.Y = You.Y + 1
                            If (You.Y - TopY - ScreenY) > CHAR_MAXYRANGE Then    'we've gone out of our range of movement, so scroll the
                                    TopY = TopY + 1                     'screen a little
                            End If
                            'finally, update the screen.
                            PaintViewport
                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                'this conversation is over
                                SelectX = NONE
                                SelectY = NONE
                                InterpretScriptLine = False
                                Exit Function
                            End If
                            
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                                You.Y = You.Y + 1
                                If (You.Y - TopY - ScreenY) > CHAR_MAXYRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopY = TopY + 1                     'screen a little
                                        If TopY = (MAP_ARRAYY - MAP_SCREENY) + 1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                                        'Things() arrays.
                                            TopY = MAP_ARRAYY - MAP_SCREENY
                                            You.Y = You.Y - 1
                                            PaintViewport   'this so the function will scroll the Map() and Things() arrays.
                                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                                'this conversation is over
                                                SelectX = NONE
                                                SelectY = NONE
                                                InterpretScriptLine = False
                                                Exit Function
                                            End If
                                            
                                        End If
                                End If
                            Next intStart
                            'finally, update the screen.
                            PaintViewport
                            
                        Else    'full move sequence.
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                                You.Y = You.Y + 1
                                If (You.Y - TopY - ScreenY) > CHAR_MAXYRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopY = TopY + 1                     'screen a little
                                End If
                                'update the screen every time now.
                                PaintViewport
                                If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                    'this conversation is over
                                    SelectX = NONE
                                    SelectY = NONE
                                    InterpretScriptLine = False
                                    Exit Function
                                End If
                                
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop

                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                        
                    Else    'we've got to find a person of that Description in the vicinity...
                        'note that once again, I've re-used some int variable names...
                        intStart = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
                        'IsThingDesc.
                        intEnd = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
                        i = IsThingDesc(intDesc, intStart, intEnd)
                        If i = NONE Then MsgBox "No person found! Check your script code!": GoTo Continue
                        'now just paste in the code from above
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            If MoveThingTo(i, , 1) = False Then
                                'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                
                                'first save it to a Temp thing
                                Temp = Things(i)
                                Temp.Y = Temp.Y + 1
                                'next delete it from the current screen.
                                RemoveThingArray i
                                'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = PUTTED
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.X = Temp.X
                                    StoreScript.Y = Temp.Y
                                    StoreScriptCommand StoreScript
                                ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                    MsgBox "Screen Full. Person that was walking is dead."
                                End If
                            End If
                            'finally, update the screen.
                            PaintViewport
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                    
                                If MoveThingTo(i, , 1) = False Then
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.Y = Temp.Y + 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X, Temp.Y + 1, i
                                End If

                            Next intStart
                            'finally, update the screen.
                            PaintViewport

                        Else    'full move sequence.(case 2)
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                            
                                If MoveThingTo(i, , 1) = False Then
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.Y = Temp.Y + 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X, Temp.Y + 1, i
                                End If
                                'update the screen every time now.
                                PaintViewport
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop
        
                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                    End If
                    
            End Select
        
        Case "y-"
            intEnd = Len(Script(Count)) - intStart  'see if there IS a first param.
            If intEnd = 0 Then  'no params! just move once.
                intType = 0
            Else    'there are at least one params; check again.
                intEnd = InStr(Script(Count), ",")  'get the first parameter's end comma
                If intEnd = 0 Then  'oops! only one param.
                    intType = 1
                    'now parse the param.
                    intX = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))    'here, intX means the number of times to move.
                Else    'uh-oh! two params! Parse, then check again.
                    intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
                    intStart = intEnd
                    intEnd = InStr(intStart + 1, Script(Count), ",") 'get the second parameter's end comma
                    If intEnd = 0 Then  'only two params
                        intType = 2
                        intY = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart))) 'here, intY means the length of time in ms to sleep.
                    Else    'all three! wow!
                        intType = 3
                        intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))    'here, intY means the length of time in ms to sleep.
                        intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))   'here intDesc is the Desc of the other Thing to move.
                    End If
                End If
            End If
            'now figure out what to do according to the param level
            Select Case intType
                Case 0
                    'just move the Person and bounds check against the Map array, saving if need be, then PaintViewport; DON'T forget to PiantViePort
                    'it's code will change as the type does.
                    'first get person using Isthing passed selectx, selecty
                    IsThing SelectX, SelectY, i
                    If MoveThingTo(i, , -1) = False Then  'CHANGEIT
                        'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                        
                        'first save it to a Temp thing
                        Temp = Things(i)
                        Temp.Y = Temp.Y - 1 'CHANGEIT
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = PUTTED
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.X = Temp.X
                            StoreScript.Y = Temp.Y
                            StoreScriptCommand StoreScript
                            InterpretScriptLine = False
                            Exit Function
                        ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                            MsgBox "Screen Full. Person that was walking is dead."
                        End If
                    End If
                    SelectY = SelectY - 1   'CHANGEIT
                    'finally, update the screen.
                    PaintViewport
                Case 1  'warning: hereafter, only the new stuff different from the original will be commented: be warned!!!
                    'do a for loop of the previous, only painting the viewport at the end.
                    IsThing SelectX, SelectY, i
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, , -1) = False Then 'CHANGEIT
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.Y = Temp.Y - 1 'CHANGEIT
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX, SelectY - 1, i
                        End If
                        SelectY = SelectY - 1   'CHANGEIT

                    Next intStart
                    'finally, update the screen.
                    PaintViewport

                Case 2
                    'do a for loop of the directly above, painting in the loop and sleeping in the loop
                    IsThing SelectX, SelectY, i
                    sngEndTime = Timer + (intY / 1000)
                    For intStart = 1 To intX Step 1
                    
                        If MoveThingTo(i, , -1) = False Then 'CHANGEIT
                            'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                            
                            'first save it to a Temp thing
                            Temp = Things(i)
                            Temp.Y = Temp.Y - 1 'CHANGEIT
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = PUTTED
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.X = Temp.X
                                StoreScript.Y = Temp.Y
                                StoreScriptCommand StoreScript
                                InterpretScriptLine = False
                                Exit Function
                            ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                MsgBox "Screen Full. Person that was walking is dead."
                             End If
                             'now re-get the person so we can keep looping
                             i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                             IsThing SelectX, SelectY - 1, i
                        End If
                        SelectY = SelectY - 1   'CHANGEIT
                        'update the screen every time now.
                        PaintViewport
                        'now sleep:
                        Do Until Timer > sngEndTime
                            DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                        Loop

                    Next intStart
                    sngEndTime = Timer + (intY / 1000)
                    Do Until Timer > sngEndTime
                        DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                    Loop

                Case 3  'I realize this section has  lot of repeated code, etc. But it might get cleaned up eventually.
                    If intDesc = NONE Then  'we're moving You
                    
                    '*** Warning! Possible bug! It might be that when we move the arrays, the person talked to would move off the arrays. Then some of the script
                    'commands might not work. however, I will not deal with that just yet.
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            'warning! we need to change the movething If to an If that tests if we went out of range of our little movement range(3-8 or something)
                            'and scroll the screen. Then we need an if that sees If we're crossing a screen boundary and that Loads/Saves Things() and Map().
                            You.Y = You.Y - 1 '*** CHANGEIT
                            If (You.Y - TopY - ScreenY) < CHAR_MINYRANGE Then    'we've gone out of our range of movement, so scroll the'CHANGEIT
                                    TopY = TopY - 1                     'screen a little
                            End If
                            'finally, update the screen.(and scroll it)
                            PaintViewport
                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                'this conversation is over
                                SelectX = NONE
                                SelectY = NONE
                                InterpretScriptLine = False
                                Exit Function
                            End If
                            
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                                You.Y = You.Y - 1 '*** CHANGEIT
                                If (You.Y - TopY - ScreenY) < CHAR_MINYRANGE Then    'we've gone out of our range of movement, so scroll the'CHANGEIT
                                        TopY = TopY - 1                     'screen a little
                                        If TopY = -1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                                        'Things() arrays.
                                            TopY = 0
                                            You.Y = You.Y + 1
                                            'end *** CHANGEIT
                                        End If
                                End If
                                'finally, update the screen.(and scroll it)
                                PaintViewport
                                If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                    'this conversation is over
                                    SelectX = NONE
                                    SelectY = NONE
                                    InterpretScriptLine = False
                                    Exit Function
                                End If
                                    'End If     'why were these in here?!? I don't know...but until the code runs through here, I'm not taking it out, either.
                                'End If
                            Next intStart
                            'finally, update the screen.
                            PaintViewport
                            
                        Else    'full move sequence.
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                                You.Y = You.Y - 1 '*** CHANGEIT
                                If (You.Y - TopY - ScreenY) < CHAR_MINYRANGE Then    'we've gone out of our range of movement, so scroll the'CHANGEIT
                                        TopY = TopY - 1                     'screen a little
                                End If
                                'finally, update the screen every time now.(and scroll it)
                                PaintViewport
                                If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                    'this conversation is over
                                    SelectX = NONE
                                    SelectY = NONE
                                    InterpretScriptLine = False
                                    Exit Function
                                End If
                                
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop

                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                        
                    Else    'we've got to find a person of that Description in the vicinity...
                        'note that once again, I've re-used some int variable names...
                        intStart = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
                        'IsThingDesc.
                        intEnd = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
                        i = IsThingDesc(intDesc, intStart, intEnd)
                        If i = NONE Then MsgBox "No person found! Check your script code!": GoTo Continue
                        'now just paste in the code from above
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            If MoveThingTo(i, , -1) = False Then 'CHANGEIT
                                'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                
                                'first save it to a Temp thing
                                Temp = Things(i)
                                Temp.Y = Temp.Y - 1 'CHANGEIT
                                'next delete it from the current screen.
                                RemoveThingArray i
                                'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = PUTTED
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.X = Temp.X
                                    StoreScript.Y = Temp.Y
                                    StoreScriptCommand StoreScript
                                ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                    MsgBox "Screen Full. Person that was walking is dead."
                                End If
                            End If
                            'finally, update the screen.
                            PaintViewport
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                    
                                If MoveThingTo(i, , -1) = False Then 'CHANGEIT
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.Y = Temp.Y - 1 'CHANGEIT
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X, Temp.Y - 1, i 'CHANGEIT
                                End If

                            Next intStart
                            'finally, update the screen.
                            PaintViewport

                        Else    'full move sequence.(case 2)
                            sngEndTime = Timer + (intY / 1000)
                            For intStart = 1 To intX Step 1
                            
                                If MoveThingTo(i, , -1) = False Then 'CHANGEIT
                                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                    
                                    'first save it to a Temp thing
                                    Temp = Things(i)
                                    Temp.Y = Temp.Y - 1 'CHANGEIT
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = PUTTED
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.X = Temp.X
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.X, Temp.Y - 1, i 'CHANGEIT
                                End If
                                'update the screen every time now.
                                PaintViewport
                                'now sleep:
                                Do Until Timer > sngEndTime
                                    DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                                Loop
        
                            Next intStart
                            sngEndTime = Timer + (intY / 1000)
                            Do Until Timer > sngEndTime
                                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                            Loop

                        End If
                    End If
                    
            End Select
    
        Case "warp"
'repeatnum and sleeptime do not apply to this command instead they are replaced with:
'warp: x, y, [objdesc]      note that x,y are NOT optional!
'*** warning!! you must not enter values that are beyond the edge of the map!! because I don't check for that at all!! ***
            intEnd = InStr(Script(Count), ",")  'get the first parameter's end comma
            intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))  'and figure out where it is inside that space
            intStart = intEnd   'now set the new start to the old end
            intEnd = InStr(intStart + 1, Script(Count), ",")    'get the new end
            If intEnd = 0 Then  'default--the person talked to.
                'it's the last parameter(use intStart)
                intY = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))
                
                IsThing SelectX, SelectY, i
                If MoveThing(i, intX, intY) = False Then
                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                    
                    'first save it to a Temp thing
                    Temp = Things(i)
                    Temp.X = intX
                    Temp.Y = intY
                    'next delete it from the current screen.
                    RemoveThingArray i
                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                    If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                        StoreScript.ScriptType = PUTTED
                        StoreScript.Desc = Temp.Desc
                        StoreScript.Movement = Temp.Movement
                        StoreScript.Type = Temp.Type
                        StoreScript.X = Temp.X
                        StoreScript.Y = Temp.Y
                        StoreScriptCommand StoreScript
                        InterpretScriptLine = False
                        Exit Function
                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                        MsgBox "Screen Full. Person that was warping is dead."
                     End If
                     SelectX = intX
                     SelectY = intY
                     
                End If
                'now show that the person has moved
                PaintViewport
                
            Else    'it's either the player or a person identified by a description.
                intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
                'now for the last parameter(no reset necessary)
                intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))
                If intDesc = -1 Then    'we're warping YOU! more work for ME!
                
'possible OPTIMIZATION here:find a way to elim the redundant Save/Load cycle that occurs if the destination is both +/- X and Y screens
'could use a bool variable and *maybe* (long shot) go back to an if/elseif/elseif/elseif/end if structure instead an if/elseif/endif/if/elseif/endif like we have now
'this would elim the possibility of running the Save/Load twice, but only works if you only need a flag sort of thing to fig the jump...not some info that comes along
'inside the X/Y if...
'NOTE: take this comment block out when I have time; thought it thru and it should work :)
                    If (intX - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement
                        If (intX - MAP_SCREENX) >= (MAP_ARRAYX - MAP_SCREENX) + 1 Then  'OK! Full move sequence...
                            'SaveMap Fileno, ScreenX, ScreenY    'save changes of current position to disk
                            SaveThings ObjFileno, ScreenX, ScreenY
                            ScreenX = ((intX \ MAP_SCREENX) * MAP_SCREENX) - ((MAP_NUMSCREENSX \ 2) * MAP_SCREENX)
                            ScreenY = ((intY \ MAP_SCREENY) * MAP_SCREENY) - ((MAP_NUMSCREENSY \ 2) * MAP_SCREENY)
                            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
                            LoadThings ObjFileno, ScreenX, ScreenY
                            RestoreScriptCommands
                        End If
                        'then just center(+- 1 tile) the screen on you
                        TopX = (intX - MAP_SCREENX / 2) - ScreenX
                        TopY = (intY - MAP_SCREENY / 2) - ScreenY
                    ElseIf (intX - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement
                        If (intX - MAP_SCREENX) <= -1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                            SaveThings ObjFileno, ScreenX, ScreenY
                            ScreenX = ((intX \ MAP_SCREENX) * MAP_SCREENX) - ((MAP_NUMSCREENSX \ 2) * MAP_SCREENX)
                            If ScreenX < 0 Then ScreenX = 0
                            ScreenY = ((intY \ MAP_SCREENY) * MAP_SCREENY) - ((MAP_NUMSCREENSY \ 2) * MAP_SCREENY)
                            If ScreenY < 0 Then ScreenY = 0
                            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
                            LoadThings ObjFileno, ScreenX, ScreenY
                            RestoreScriptCommands
                        End If
                        TopX = (intX - MAP_SCREENX / 2) - ScreenX + 1
                        TopY = (intY - MAP_SCREENY / 2) - ScreenY + 1
                    ElseIf (intY - TopY - ScreenY) > CHAR_MAXYRANGE Then    'we've gone out of our range of movement, so scroll the
                        If (intY - MAP_SCREENY) >= (MAP_ARRAYY - MAP_SCREENY) + 1 Then 'OK! Full move sequence...
                            'SaveMap Fileno, ScreenX, ScreenY    'save changes of current position to disk
                            SaveThings ObjFileno, ScreenX, ScreenY
                            ScreenX = ((intX \ MAP_SCREENX) * MAP_SCREENX) - ((MAP_NUMSCREENSX \ 2) * MAP_SCREENX)
                            If ScreenX < 0 Then ScreenX = 0
                            ScreenY = ((intY \ MAP_SCREENY) * MAP_SCREENY) - ((MAP_NUMSCREENSY \ 2) * MAP_SCREENY)
                            If ScreenY < 0 Then ScreenY = 0
                            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
                            LoadThings ObjFileno, ScreenX, ScreenY
                            RestoreScriptCommands
                        End If
                        'then just center(+- 1 tile) the screen on you
                        TopX = (intX - MAP_SCREENX / 2) - ScreenX
                        TopY = (intY - MAP_SCREENY / 2) - ScreenY
                    ElseIf (intY - TopY - ScreenY) < CHAR_MINYRANGE Then    'we've gone out of our range of movement
                        If (intY - MAP_SCREENY) <= -1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                            SaveThings ObjFileno, ScreenX, ScreenY
                            ScreenX = ((intX \ MAP_SCREENX) * MAP_SCREENX) - ((MAP_NUMSCREENSX \ 2) * MAP_SCREENX)
                            If ScreenX < 0 Then ScreenX = 0
                            ScreenY = ((intY \ MAP_SCREENY) * MAP_SCREENY) - ((MAP_NUMSCREENSY \ 2) * MAP_SCREENY)
                            If ScreenY < 0 Then ScreenY = 0
                            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
                            LoadThings ObjFileno, ScreenX, ScreenY
                            RestoreScriptCommands
                        End If
                        TopX = (intX - MAP_SCREENX / 2) - ScreenX
                        TopY = (intY - MAP_SCREENY / 2) - ScreenY
                    End If
                    You.X = intX
                    You.Y = intY
                    'finally, update the screen.
                    PaintViewport
                    If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                        'this conversation is over
                        SelectX = NONE
                        SelectY = NONE
                        InterpretScriptLine = False
                        Exit Function
                    End If

                Else
                    'note that once again, I've re-used some int variable names...
                    intStart = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
                    'IsThingDesc.
                    intEnd = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
                    i = IsThingDesc(intDesc, intStart, intEnd)
                    If MoveThing(i, intX, intY) = False Then
                        'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                       
                        'first save it to a Temp thing
                        Temp = Things(i)
                        Temp.X = Temp.X + 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            StoreScript.ScriptType = PUTTED
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.X = Temp.X
                            StoreScript.Y = Temp.Y
                            StoreScriptCommand StoreScript
                        ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                            MsgBox "Screen Full. Person that was walking is dead."
                        End If
                    End If
                    'finally, update the screen.
                    PaintViewport
                
                End If
            End If
            
        Case "give"    'Give the player something
            '(Give:1, 2, 3)
            '1 = Type, 2 = Desc, 3 = movement(note:must be a number, not a constant)
            Temp.X = 0 'init the x,y to some innocuous but non-NULL value.
            Temp.Y = 0
            intEnd = InStr(Script(Count), ",")  'get the first parameter's end comma
            Temp.Type = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))  'and figure out where it is inside that space
            intStart = intEnd   'now set the new start to the old end
            intEnd = InStr(intStart + 1, Script(Count), ",")    'get the new end
            Temp.Desc = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))  'and figure the parameter again.
            'now for the last parameter(no reset necessary)
            Temp.Movement = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))
            'now call the function that 'gives' the 'Thing'
            If GiveThing(You, lvwPossessions, Temp) = False Then   'uh-oh, the player is out of room.
                'do nothing now, but maybe code some action here.
                MsgBox "Error: Player's possessions full!"
            End If
        Case "take"    'taKe something away from the player
            '(Take: 1)   1 = description number of the object
            
            intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))
            If HaveThing(You, intDesc, ArrayNum) = True Then
                TakeThing You, lvwPossessions, ArrayNum
            End If
        Case "put"    'Put something on the world map
            '(Put:111, 999, 2, 3, 0)
            '111 = X position on the map
            '999 = Y position on the map
            '2 = the type(i.e. picture and 'Thing' type
            '3 = the description number
            '0 = the movement value(no constants allowed)
            intEnd = InStr(Script(Count), ",")
            intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))

            intStart = intEnd
            intEnd = InStr(intStart + 1, Script(Count), ",")
            intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
            
            intStart = intEnd
            intEnd = InStr(intStart + 1, Script(Count), ",")
            Temp.Desc = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
            
            intStart = intEnd
            intEnd = InStr(intStart + 1, Script(Count), ",")
            Temp.Type = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
            
            Temp.Movement = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))
            Temp.X = intX
            Temp.Y = intY
            If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                'location out of range of Map(); store it for future use...
                StoreScript.ScriptType = PUTTED
                StoreScript.Desc = Temp.Desc
                StoreScript.Movement = Temp.Movement
                StoreScript.Type = Temp.Type
                StoreScript.X = intX
                StoreScript.Y = intY
                StoreScriptCommand StoreScript
                GoTo Continue   'the break; hack
            End If

            If PutThing(Temp, ScreenX, ScreenY) = False Then    'uh-oh: screen full. cancel changes, but do nothing
            'else(now currently anyway)
                MsgBox "Screen Full! Cannot put down another Thing"
            End If
        Case "remove"    'remove something from the world map
            '(Remove:1, 4) 'optionally: (Remove:1, 4, 625, 512)
            '1 = description number of thing to remove(this is iffy for multiple instances of one thing per screen.
            '4 = type number of thing to remove(i.e. the picture)
            'It could be quite unreliable for killing enemies.)(Or in a place with a LOT of hay)
            'optionally: 625 = X pos, 512 = Y pos. Note: if there is no object matching the desc num at the position,
            'remove will search the whole screen for an object with that desc num.
            intEnd = InStr(Script(Count), ",")
            intDesc = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
            intStart = intEnd
            intEnd = InStr(intStart + 1, Script(Count), ",")
            If intEnd = 0 Then  'just a standard(no X,Y) remove command)
                intType = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))   'parse the value, then
                'jump down and check the current screen.
            Else
                intType = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1))) 'else parse X,Y value
                'and check them first.
                intStart = intEnd
                intEnd = InStr(intStart + 1, Script(Count), ",")
                intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
                intY = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))
                'make sure that the co-ordinates are in the range of Map(); if not, store using StoreScriptCommand
                If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                    'location out of range of Map(); store it for future use...
                    StoreScript.ScriptType = REMOVE
                    StoreScript.Desc = Temp.Desc
                    StoreScript.Type = Temp.Type
                    StoreScript.X = intX
                    StoreScript.Y = intY
                    StoreScriptCommand StoreScript
                    GoTo Continue   'use the break; hack
                End If
                'now check the X,Y position first
                If IsThing(intX, intY) Then RemoveThing intX, intY
                'then go through the rigamarole of checking the rest if it's not found
            End If
            
            'this function called checks the screen given, then the whole Thing(array).
            intX = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
            'IsThingDesc.
            intY = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
            i = IsThingDesc(intDesc, intX, intY)
            If i <> NONE Then RemoveThingArray i   'remove it if found.
            
        Case "end"    'end conversation
            '(End:)
            InterpretScriptLine = False 'the only reason that InterpretScriptLine is a Function is for this reason.
            'If there's an End: coded in the script, INterpretScriptLine returns false. the calling function then
            'checks the return value. If it's false, the conversation should terminate. In the case of the recursive
            'scripting commands(have, question, etc.) when the second instance of InterpretScriptLine returns False,
            'they will return False as well.
            Exit Function
        
        Case "have"    'if you haVe something. This requires that the scripter supply a yes block
                        'and a no block.
            '(Have:12) 12 = description number
            intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart))))
            If HaveThing(You, intDesc) Then  'execute the yes portion of instructions
                Do Until Trim$(LCase$(Script(Count))) = "havyes"
                    Count = Count + 1
                Loop
                Count = Count + 1
                'now we're at the start of the yes instructions
                Do Until Trim$(Script(Count)) = ">" 'the signal for the end of a have statement block
                    If InterpretScriptLine(Count) = False Then  'obviously the person saw that you HAVE the
                    'poison coated saber and is running in fear.(ending the conversation)
                        InterpretScriptLine = False
                        Exit Function
                    End If
                Loop

                'now loop to end of have structure, and continue reading script
                Do Until Trim$(Script(Count)) = "/have"
                    Count = Count + 1
                Loop

            Else    'execute the no portion
                Do Until Trim$(LCase$(Script(Count))) = "havno"
                    Count = Count + 1
                Loop
                Count = Count + 1
                'now we're at the start of the no instructions
                Do Until Trim$(Script(Count)) = ">" 'the signal for the end of a have statement block
                    If InterpretScriptLine(Count) = False Then  'obviously the person saw that you didn't HAVE the
                    'pot o' gold and is leaving in disgust.(ending the conversation)
                        InterpretScriptLine = False
                        Exit Function
                    End If
                Loop
                
                'now loop to end of have structure, and continue reading script
                Do Until Trim$(Script(Count)) = "/have"
                    Count = Count + 1
                Loop

            End If
        Case "question"    'ask the player a Question. This requires that the scripter supply a yes block
                        'and a no block.
            '(Question:Do you want to sell that poison coated saber?)
            Result = OpenTalkBox(Right$(Script(Count), Len(Script(Count)) - (intStart)), "Yes", "No")
            If Result = 0 Then  'yes
                Do Until Trim$(LCase$(Script(Count))) = "ansyes"    'this allows for comments in between
                    Count = Count + 1
                Loop
                'now we're at the start of the yes block, so step through and interpret the instructions there
                Do Until Trim$(Script(Count)) = ">"    'the signal for the end of any kind of if block.
                    If InterpretScriptLine(Count) = False Then  'there is an "end" embedded
                        InterpretScriptLine = False         'somewhere in there.
                        Exit Function
                    End If
                Loop
                'now loop to end of question structure, and continue reading script
                Do Until Trim$(Script(Count)) = "/question"
                    Count = Count + 1
                Loop
            Else    'probably no
                Do Until Trim$(LCase$(Script(Count))) = "ansno"
                    Count = Count + 1
                Loop
                'now we're at the start of the no block
                Do Until Trim$(Script(Count)) = ">"
                    If InterpretScriptLine(Count) = False Then  'we found an "end" somewhere.
                        InterpretScriptLine = False
                        Exit Function
                    End If
                Loop
                'now loop to end of question structure, and continue reading script
                Do Until Trim$(Script(Count)) = "/question"
                    Count = Count + 1
                Loop
            End If
        Case "isthread" 'lets the scripter check to see if another thread currently has a certain value in it.
'        '(IsThread:99)
'        'optionally:(IsThread:88-99) 'to check a range of values.
'        'start value: y|n; end value: >
'       'whole block is looks as such:   isthread: 88-99...y...>...n...>.../isthread
'        '88-99 = value that might be in the thread.
'        'New:Leave thread number out.
            intEnd = InStr(Script(Count), "-")
            bFound = False  'default to False
            If intEnd <> 0 Then  'the scripter is specifying a range of
            'threads (IsThread:88-99)
                For i = 0 To NUMTHREADS Step 1
                    If Threads(i) >= CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - (intStart + 1)))) _
                    And Threads(i) <= CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intEnd)))) Then
                    'IsThread:0-999
                    'we have to be >= than the first number and <= than the second.
                        bFound = True
                    End If
                Next i
            Else    'the thread structure doesn't specify a range, just a single one...
                For i = 0 To NUMTHREADS Step 1
                    If Threads(i) = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart)))) Then
                        bFound = True
                    End If
                Next i
            End If
            
            'now interpret the instructions inside the yes or no blocks.(same code as question and have commands)
            If bFound = True Then  'yes
                Do Until Trim$(LCase$(Script(Count))) = "isyes"    'this allows for comments in between
                    Count = Count + 1
                Loop
                'now we're at the start of the yes block, so step through and interpret the instructions there
                Do Until Trim$(Script(Count)) = ">"    'the signal for the end of any kind of if block.
                    If InterpretScriptLine(Count) = False Then  'there is an "end" embedded
                        InterpretScriptLine = False         'somewhere in there.
                        Exit Function
                    End If
                Loop

                'now loop to end of isthread structure, and continue reading script
                Do Until Trim$(Script(Count)) = "/isthread"
                    Count = Count + 1
                Loop
            Else    'probably no
                Do Until Trim$(LCase$(Script(Count))) = "isno"
                    Count = Count + 1
                Loop
                'now we're at the start of the no block
                Do Until Trim$(Script(Count)) = ">"
                    If InterpretScriptLine(Count) = False Then  'we found an "end" somewhere.
                        InterpretScriptLine = False
                        Exit Function
                    End If
                Loop
                'now loop to end of isthread structure, and continue reading script
                Do Until Trim$(Script(Count)) = "/isthread"
                    Count = Count + 1
                Loop

            End If
        Case "chmap"
'            (ChMap:111, 999, 22)   'NOTE: X and Y START AT 0!!! This is VERY important!
'            111 = X pos, 999 = y pos, 22 = map type
            intEnd = InStr(Script(Count), ",")
            intX = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
            intStart = intEnd
            intEnd = InStr(intStart + 1, Script(Count), ",")
            intY = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - intStart - 1)))
            intType = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))
            If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                StoreScript.ScriptType = CHMAP
                StoreScript.X = intX
                StoreScript.Y = intY
                StoreScript.Type = intType
                StoreScriptCommand StoreScript
                GoTo Continue:
            End If
            Map(((intY - ScreenY) * MAP_ARRAYX) + (intX - ScreenX)) = intType    'voila!, it's changed
        Case "sleep"
            '(Sleep:250)
            '250 = number of ms to pause everything(useful perhaps for animation type stuff)
            'NOTE:Used intDesc here to avoid creating yet another variable
            '!!DESC HAS NOTHING TO DO WITH THE SLEEP COMMAND!!
            intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intStart)))
            sngEndTime = Timer + (intDesc / 1000)
            Do Until Timer > sngEndTime
                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
            Loop
            'as an aside, my original code looked a little different as developed in QBasic/C style enviroment, but
            'I changed the method after seeing it done this way a couple of times by Mastering VB 5.
            'The moral of the story? Mastering VB 5 shouldn't be ignored because it can come in handy.

'        Case "paint"
'            PaintViewport   'duh: this is simple(and unneeded: commented out for now)
        Case "chthread"    'change a thread
            '(Thread:0, 1)
            ' 0 = number of thread to change, 1 = number to which to change the thread.
            intEnd = InStr(Script(Count), ",")
            OldThread = CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - (intStart + 1))))
            NewThread = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - intEnd)))
            Threads(OldThread) = NewThread  'this is Ryan's idea... a revised method.
    End Select
    
    End If  'end if there's a ':' in the line
Continue:       'this is a hack to emulate the 'break;' statement from C programming.(carelessly omitted by the very
'structure of 'Select Case')
    Count = Count + 1   'make sure we increment the line count.(even if it wasn't a command)
    InterpretScriptLine = True  'tell the user that it's OK to continue.  (0_0) Lil' Arf an' Nonnie
End Function
Private Function StoreScriptCommand(StoreScript As StoredScript) As Boolean
Dim i As Integer
Dim bFound As Boolean
    bFound = False
    For i = 0 To MAXSTOREDCOMM Step 1
        If StoredScriptCommands(i).X = NONE Then
            bFound = True
            Exit For
        End If
    Next i
    
    If bFound = True Then
        StoredScriptCommands(i) = StoreScript
    End If
    
    StoreScriptCommand = bFound
End Function
Private Sub RestoreScriptCommands()
Dim i As Integer
    'WARNING!! WARNING!!: the value of + 29 COULD be incorrect.
    'It might be + 30 instead. This is because + 30 has been used time out of mind
    'in PaintViewPort to detect if we're stepping off the edge of the map...
    For i = 0 To MAXSTOREDCOMM Step 1
        If (StoredScriptCommands(i).X > ScreenX) And (StoredScriptCommands(i).X < ScreenX + (MAP_ARRAYX - 1)) _
        And (StoredScriptCommands(i).Y > ScreenY) And (StoredScriptCommands(i).Y < ScreenY + (MAP_ARRAYY - 1)) Then   'it's a hit
            InterpretStoredScript StoredScriptCommands(i)
            StoredScriptCommands(i).X = NONE  'make sure we invalidate this record so that it can be overwritten and
            'will not be executed again.
        End If
    Next i
End Sub
Private Sub InterpretStoredScript(StoreScript As StoredScript)
    Dim OldThread As Integer    'this keeps track of the old thread which is changed in the 'Thread'
'command
Dim NewThread As Integer    'do you really want me to tell you about this??
Dim i As Integer    'a simple counter
Dim intStart As Integer, intEnd As Integer
Dim ArrayNum As Integer
Dim intX As Integer, intY As Integer
Dim Temp As Thing   'a temporary thing to store stuff in before modifying it.
Dim intObjectScreen As Integer, intDesc As Integer
Dim bFound As Boolean
Dim Result As Integer
Dim intType As Integer   'what the type of map is that we're supposed to change to in 'chmap'
Dim sngEndTime As Single
Dim StoreScriptComm As StoredScript
    Select Case StoreScript.ScriptType 'figure out what command the all-powerful scripter
    'is commanding us to carry out.
        Case CHAT    'a simple chatty OpenTalkBox(note: pass this command a value bigger than 0 in the X value to
            'let it know that you have already called LoadTalkBox
            '(Chat:Hello, my name is Mikey and I want you to vote for ME!)
            If StoreScript.X < 1 Then
                LoadTalkBox
                OpenTalkBox StoreScript.Tag
                UnloadTalkBox
            Else
                OpenTalkBox StoreScript.Tag
            End If
        Case GIVE    'Give the player something
            '(Give:1, 2, 3)
            '1 = Type( .type), 2 = Desc( .desc), 3 = movement (.movement)
            Temp.X = 0 'init the x,y to some innocuous but non-NULL value.
            Temp.Y = 0
            Temp.Type = StoreScript.Type
            Temp.Desc = StoreScript.Desc
            Temp.Movement = StoreScript.Movement
            'now call the function that 'gives' the 'Thing'
            If GiveThing(You, lvwPossessions, Temp) = False Then   'uh-oh, the player is out of room.
                'do nothing now, but maybe code some action here.
                MsgBox "Error: Player's possessions full!"
            End If
        Case TAKE    'taKe something away from the player
            '(Take: 1)   1 = description number of the object( .desc)

            intDesc = StoreScript.Desc
            If HaveThing(You, intDesc, ArrayNum) = True Then
                TakeThing You, lvwPossessions, ArrayNum
            End If
        Case PUTTED    'Put something on the world map
            '(Put:111, 999, 2, 3, 0)
            '111 = X position on the map
            '999 = Y position on the map
            '2 = the type(i.e. picture and 'Thing' type (.type)
            '3 = the description number( .desc)
            '0 = the movement value (.movement)
            Temp.Type = StoreScript.Type
            Temp.Desc = StoreScript.Desc
            Temp.Movement = StoreScript.Movement
            Temp.X = StoreScript.X
            Temp.Y = StoreScript.Y
            If (Temp.X < ScreenX) Or (Temp.X > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYX - 1)) Then
                'location out of range of Map(); store it for future use...
                'REPLACE WITH StoreScriptCommand(StoredCommand as StoredCommand)
                MsgBox "Warning! Check call stack and above If for possible bugs!"
                StoreScriptComm.ScriptType = PUTTED
                StoreScriptComm.Desc = Temp.Desc
                StoreScriptComm.Movement = Temp.Movement
                StoreScriptComm.Type = Temp.Type
                StoreScriptComm.X = intX
                StoreScriptComm.Y = intY
                StoreScriptCommand StoreScriptComm
            End If

            If PutThing(Temp, ScreenX, ScreenY) = False Then    'uh-oh: screen full. cancel changes, but do nothing
            'else(now currently anyway)
                MsgBox "Screen Full! Cannot put down another Thing"
            End If
            
        Case REMOVE    'remove something from the world map
        'NOTE: X MUST be -99(NOT_GIVEN) if you tell this command to use the default value(this screen, then the current
        'whole Map() array.
            '(Remove:1, 4) 'optionally: (Remove:1, 4, 625, 512)
            '1 = description number of thing to remove(this is iffy for multiple instances of one thing per screen. (.desc)
            '4 = type number of thing to remove(i.e. the picture) ( .type)
            'It could be quite unreliable for killing enemies.)(Or in a place with a LOT of hay)
            'optionally: 625 = X pos( .x), 512 = Y pos( .y). Note: if there is no object matching the desc num at the position,
            'remove will search the whole screen for an object with that desc num.
            intDesc = StoreScript.Desc
            intType = StoreScript.Type
            If StoreScript.X = NOT_GIVEN Then  'just a standard(no X,Y) remove command)
                intObjectScreen = (((You.Y - ScreenY) \ MAP_SCREENX) * MAP_ARRAYX) + (((You.X - ScreenX) \ MAP_SCREENX) * OBJ_MAXTHINGSSCREEN)
                'just jump down and check the current screen.(skip the precise X,Y check)
            Else
                intObjectScreen = (((StoreScript.Y - ScreenY) \ MAP_SCREENX) * MAP_ARRAYX) + (((StoreScript.X - ScreenX) \ MAP_SCREENX) * OBJ_MAXTHINGSSCREEN)
                intX = StoreScript.X
                intY = StoreScript.Y
                'make sure that the co-ordinates are in the range of Map(); if not, store using StoreScriptCommand
                If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                    'location out of range of Map(); store it for future use...
                    StoreScriptComm.ScriptType = REMOVE
                    StoreScriptComm.Desc = Temp.Desc
                    StoreScriptComm.Type = Temp.Type
                    StoreScriptComm.X = intX
                    StoreScriptComm.Y = intY
                    StoreScriptCommand StoreScriptComm
                End If
                'now check the X,Y position first
                If IsThing(intX, intY) Then RemoveThing intX, intY
                'then go through the rigamarole of checking everything again.
            End If
            For i = intObjectScreen To intObjectScreen + 9 Step 1 'loop through all the 'Things' on this particular screen.
                If Things(i).Desc = intDesc And Things(i).Type = intType And Things(i).X <> NONE Then   'a match!!
                    RemoveThingArray i
                    bFound = True
                    Exit For
                End If
            Next i
            If bFound = False Then   'check the whole array because we didn't find the object on the same screen
            'as the person being talked to.
                For i = 0 To (OBJ_MAXTHINGSARRAY - 1) Step 1 'the WHOLE object array.
                    If Things(i).Desc = intDesc And Things(i).Type = intType And Things(i).X <> NONE Then RemoveThingArray i
                    'remove it from the array. Note that this approach may unduly weight the incidence of removed objects
                    'on the topleft screen, but this cannot be helped.
                Next i
            End If
'*** UNSUPPORTED ***
'        Case "end"    'end conversation
'            '(End:)
'            InterpretScriptLine = False 'the only reason that InterpretScriptLine is a Function is for this reason.
'            'If there's an End: coded in the script, INterpretScriptLine returns false. the calling function then
'            'checks the return value. If it's false, the conversation should terminate. In the case of the recursive
'            'scripting commands(have, question, etc.) when the second instance of InterpretScriptLine returns False,
'            'they will return False as well.
'            Exit Sub
''Brainstorm!! Instead of the programmer providing a yes/no block somehow, why don't we just stick a value meaning 'True' in one of the StoredScript
''slots that we're given? Then the programmer can check that since we don't have a return value
'        Case "have"    'if you haVe something. This requires that the scripter supply a yes block
'                        'and a no block. Since you can't do that with one shot, this isn't supported (yet).
'            '(Have:12) 12 = description number
'            intDesc = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart))))
'            If HaveThing(You, intDesc) Then  'execute the yes portion of instructions
'                Do Until Trim$(LCase$(Script(Count))) = "havyes"
'                    Count = Count + 1
'                Loop
'                Count = Count + 1
'                'now we're at the start of the yes instructions
'                Do Until Trim$(Script(Count)) = ">" 'the signal for the end of a have statement block
'                    If InterpretScriptLine(Count) = False Then  'obviously the person saw that you HAVE the
'                    'poison coated saber and is running in fear.(ending the conversation)
'                        InterpretScriptLine = False
'                        Exit Sub
'                    End If
'                Loop
'
'                'now loop to end of have structure, and continue reading script
'                Do Until Trim$(Script(Count)) = "/have"
'                    Count = Count + 1
'                Loop
'
'            Else    'execute the no portion
'                Do Until Trim$(LCase$(Script(Count))) = "havno"
'                    Count = Count + 1
'                Loop
'                Count = Count + 1
'                'now we're at the start of the no instructions
'                Do Until Trim$(Script(Count)) = ">" 'the signal for the end of a have statement block
'                    If InterpretScriptLine(Count) = False Then  'obviously the person saw that you didn't HAVE the
'                    'pot o' gold and is leaving in disgust.(ending the conversation)
'                        InterpretScriptLine = False
'                        Exit Sub
'                    End If
'                Loop
'
'                'now loop to end of have structure, and continue reading script
'                Do Until Trim$(Script(Count)) = "/have"
'                    Count = Count + 1
'                Loop
'
'            End If
'        Case "question"    'ask the player a Question. This requires that the scripter supply a yes block
'                        'and a no block.
'            '(Question:Do you want to sell that poison coated saber?)
'            Result = OpenTalkBox(Right$(Script(Count), Len(Script(Count)) - (intStart)), "Yes", "No")
'            If Result = 0 Then  'yes
'                Do Until Trim$(LCase$(Script(Count))) = "ansyes"    'this allows for comments in between
'                    Count = Count + 1
'                Loop
'                'now we're at the start of the yes block, so step through and interpret the instructions there
'                Do Until Trim$(Script(Count)) = ">"    'the signal for the end of any kind of if block.
'                    If InterpretScriptLine(Count) = False Then  'there is an "end" embedded
'                        InterpretScriptLine = False         'somewhere in there.
'                        Exit Sub
'                    End If
'                Loop
'                'now loop to end of question structure, and continue reading script
'                Do Until Trim$(Script(Count)) = "/question"
'                    Count = Count + 1
'                Loop
'            Else    'probably no
'                Do Until Trim$(LCase$(Script(Count))) = "ansno"
'                    Count = Count + 1
'                Loop
'                'now we're at the start of the no block
'                Do Until Trim$(Script(Count)) = ">"
'                    If InterpretScriptLine(Count) = False Then  'we found an "end" somewhere.
'                        InterpretScriptLine = False
'                        Exit Sub
'                    End If
'                Loop
'                'now loop to end of question structure, and continue reading script
'                Do Until Trim$(Script(Count)) = "/question"
'                    Count = Count + 1
'                Loop
'            End If
'        Case "isthread" 'lets the scripter check to see if another thread currently has a certain value in it.
''        '(IsThread:99)
''        'optionally:(IsThread:88-99) 'to check a range of values.
''        'start value: y|n; end value: >
''        '88-99 = value that might be in the thread.
''        'New:Leave thread number out.
'            intEnd = InStr(Script(Count), "-")
'            If intEnd <> 0 Then  'the scripter is specifying a range of
'            'threads (IsThread:88-99)
'            Result = 1  'default to 'no'
'                For i = 0 To NUMTHREADS Step 1
'                    If Threads(i) >= CInt(Trim$(Mid$(Script(Count), intStart + 1, intEnd - (intStart + 1)))) And Threads(i) <= CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intEnd)))) Then
'                    't0-999
'                    'we have to be >= than the first number and <= than the second.
'                        Result = 0  'meaning yes
'                    End If
'                Next i
'            Else    'the thread structure doesn't specify a range, just a single one...
'                For i = 0 To NUMTHREADS Step 1
'                    If Threads(i) = CInt(Trim$(Right$(Script(Count), Len(Script(Count)) - (intStart)))) Then
'                        bFound = True
'                    End If
'                Next i
'            End If
'
'            'now interpret the instructions inside the yes or no blocks.(same code as question and have commands)
'            If bFound = True Then  'yes
'                Do Until Trim$(LCase$(Script(Count))) = "isyes"    'this allows for comments in between
'                    Count = Count + 1
'                Loop
'                'now we're at the start of the yes block, so step through and interpret the instructions there
'                Do Until Trim$(Script(Count)) = ">"    'the signal for the end of any kind of if block.
'                    If InterpretScriptLine(Count) = False Then  'there is an "end" embedded
'                        InterpretScriptLine = False         'somewhere in there.
'                        Exit Sub
'                    End If
'                Loop
'
'                'now loop to end of isthread structure, and continue reading script
'                Do Until Trim$(Script(Count)) = "/isthread"
'                    Count = Count + 1
'                Loop
'            Else    'probably no
'                Do Until Trim$(LCase$(Script(Count))) = "isno"
'                    Count = Count + 1
'                Loop
'                'now we're at the start of the no block
'                Do Until Trim$(Script(Count)) = ">"
'                    If InterpretScriptLine(Count) = False Then  'we found an "end" somewhere.
'                        InterpretScriptLine = False
'                        Exit Sub
'                    End If
'                Loop
'                'now loop to end of isthread structure, and continue reading script
'                Do Until Trim$(Script(Count)) = "/isthread"
'                    Count = Count + 1
'                Loop
'
'            End If
'*** END UNSUPPORTED ***
        Case CHMAP
'            (ChMap:111, 999, 22)   'NOTE: X and Y START AT 0!!! This is VERY important!
'            111 = X pos(.x), 999 = y pos (.y), 22 = map type (.type)
            intX = StoreScript.X
            intY = StoreScript.Y
            If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                MsgBox "Warning: Check call stack and above if for possible bug!"
                Exit Sub
            End If
            intType = StoreScript.Type
            Map(((intY - ScreenY) * MAP_ARRAYX) + (intX - ScreenX)) = intType    'voila!, it's changed
'        Case "paint"
'            PaintViewport   'duh: this is simple(and unneeded: commented out for now)
        Case SLEEP
            '(Sleep:250)
            '250 = number of ms to pause everything(useful perhaps for animation type stuff)(.x)
            '(i.e. OH NO!!! Not Final Fantasy movies.... Yaaahh!)
            'NOTE:Used intDesc here to avoid creating yet another variable
            '!!DESC HAS NOTHING TO DO WITH THE SLEEP COMMAND!!
            intDesc = StoreScript.X
            sngEndTime = Timer + (intDesc / 1000)
            Do Until Timer > sngEndTime
                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                'although this could be dangerous if they used the arrow keys to scroll alot.
            Loop
        Case CHTHREAD    'change a thread
            '(Thread:0, 1)
            ' 0 = number of thread to change(storescript.x), 1 = number to which to change the thread.(storescript.y)
            OldThread = StoreScript.X
            NewThread = StoreScript.Y
            Threads(OldThread) = NewThread  'this is Ryan's idea... a revised method.

    End Select
    
'Continue:       'this is a hack to emulate the 'break;' statement from C programming.(carelessly omitted by the very
'structure of 'Select Case') though in this function (excuse me) sub, an Exit Sub will do just as well as there is no
'critical incrementing and returning of values to be done.
End Sub
Private Function HaveThing(Ch As Char, Desc As Integer, Optional ByRef ArrayNum As Integer = NOT_GIVEN) As Boolean
    'ch is the person to search, search should contain the info on the Thing that we're
    'searching for
Dim i As Integer    'a simple incrementing variable
Dim Found As Boolean
    For i = 0 To CHAR_MAXPOSSESSIONS Step 1
        If Desc = Ch.Possessions(i).Desc Then
            Found = True
            Exit For
        End If
    Next i
    If Found = True And ArrayNum <> NOT_GIVEN Then
        ArrayNum = i
    Else
        ArrayNum = NONE
    End If
    HaveThing = Found
End Function
Private Function GiveThing(Ch As Char, lvwCh As ListView, Givee As Thing) As Boolean
    'ch is the person we're giving to, givee is the thing which we are giving
Dim i As Integer    'a simple incrementable variable, unwise in the ways of the world
Dim Found As Boolean
    Found = False   'init Found to be absolutely sure.
    'find a free space in the person's possessions
    For i = 0 To CHAR_MAXPOSSESSIONS Step 1
        If Ch.Possessions(i).X = NONE Then
            Found = True
            Exit For
        End If
    Next i
    
    If Found = True Then
        'give the thing to you
        Ch.Possessions(i) = Givee
        With lvwCh
            Dim itmX As ListItem
            Set itmX = .ListItems.Add()
            'ListNum(i) = itmX.Index    'this would save the index in an array... but the index can change
            'if you remove 1 or two. so we would want to refer to something by its 'key'.
            itmX.Key = "#" & CStr(i)
            itmX.Icon = Ch.Possessions(i).Type
            itmX.Text = imlThings.ListImages(Ch.Possessions(i).Type).Key
            itmX.SubItems(LVW_DESC) = CStr(Ch.Possessions(i).Desc)
            If Ch.Possessions(i).Movement = STILL Then
                itmX.SubItems(LVW_MOVEMENT) = "Still"
            ElseIf Ch.Possessions(i).Movement = RANDOM Then
                itmX.SubItems(LVW_MOVEMENT) = "Random"
            Else    'give it a guess and say follow
                itmX.SubItems(LVW_MOVEMENT) = "Follow"
            End If
        End With
    End If
    
    GiveThing = Found   'True or False, respectively
End Function
Private Sub TakeThing(Ch As Char, lvwCh As ListView, ArrayRemove As Integer)
    'Ch is the the person to take from. lvwCh is the listview to remove from, and Removenum is the number in the
    'possessions array to remove
        Ch.Possessions(ArrayRemove).X = NONE
        lvwCh.ListItems.REMOVE lvwCh.ListItems("#" & CStr(ArrayRemove)).Index
End Sub

