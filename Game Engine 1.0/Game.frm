VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
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
   Begin VB.Timer tmrMoveThings 
      Interval        =   500
      Left            =   4200
      Top             =   5040
   End
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
      _Version        =   327682
      Icons           =   "imlThings"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
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
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Weight"
         Object.Width           =   2646
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
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   120
      MousePointer    =   4  'Icon
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   4
      Top             =   480
      Width           =   4800
   End
   Begin ComctlLib.ImageList imlThings 
      Left            =   4440
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   82
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":0000
            Key             =   "Purina Table"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":0C52
            Key             =   "Potted Bush"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":18A4
            Key             =   "blank(do not use)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":24F6
            Key             =   "Potted Palm"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3148
            Key             =   "Clay Pot"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3D9A
            Key             =   "Haystack"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":49EC
            Key             =   "Iron Pot"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":563E
            Key             =   "Inscribed Pot"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6290
            Key             =   "Ballot Box"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6EE2
            Key             =   "Brick"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":7B34
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":8786
            Key             =   "Bottle"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":8FD8
            Key             =   "Bottles"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":982A
            Key             =   "Dictionary"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":A47C
            Key             =   "Park Bench"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":ACCE
            Key             =   "TGP"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":B920
            Key             =   "Deluxe R Pizza"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":C572
            Key             =   "Karrot"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":D1C4
            Key             =   "Ridiculous Pizza"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":DE16
            Key             =   "Ridiculous Pizza Box"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":EA68
            Key             =   "Deluxe R Pizza Box"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":F6BA
            Key             =   "NoteinBottle"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1030C
            Key             =   "BottleStack"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":10F5E
            Key             =   "L5 Door"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":11BB0
            Key             =   "WoodenDoor"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":12802
            Key             =   "DeepHole"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":13454
            Key             =   "Monster Pit"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":140A6
            Key             =   "TellyBooth"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":148F8
            Key             =   "BoatFloat"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1554A
            Key             =   "BoatIcon"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1619C
            Key             =   "FF1Well"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":16DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":17A40
            Key             =   "Mikey le Mouse"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":17CD2
            Key             =   "The Chatty Lady"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":18924
            Key             =   "Miney le Mouse"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":18BB6
            Key             =   "Professor"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":19808
            Key             =   "Mega Mouse"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1A45A
            Key             =   "Fred the Freeloader"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1B0AC
            Key             =   "Macky Le Mouse"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1B33E
            Key             =   "Kat"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1BF90
            Key             =   "Easy Fix"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1C7E2
            Key             =   "Shuffler"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1D034
            Key             =   "Grandpa Clone"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1D886
            Key             =   "GrandPa #13"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1E0D8
            Key             =   "Bottle Stacker"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1E92A
            Key             =   "Bunny"
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1F17C
            Key             =   "Durty Kurt"
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":1F9CE
            Key             =   "Ears the Rabbit"
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":20220
            Key             =   "Expendable Crewman"
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":20A72
            Key             =   "Eyes the Rabbit"
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":212C4
            Key             =   "Fat Rabbit"
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":21B16
            Key             =   "Feet the Rabbit"
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":22368
            Key             =   "Ganwa"
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":22BBA
            Key             =   "Klown"
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2340C
            Key             =   "Kangaroo Rat"
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":23C5E
            Key             =   "Old Rabbit"
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":244B0
            Key             =   "Sun Glassed Rat"
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":24D02
            Key             =   "Pierre"
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":25554
            Key             =   "Rat"
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":25DA6
            Key             =   "Shipwrecked Guy"
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":265F8
            Key             =   "Solo"
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":26E4A
            Key             =   "Spork"
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2769C
            Key             =   "Spotted Rabbit"
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":27EEE
            Key             =   "Stinky"
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":28740
            Key             =   "Blind Rat"
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":28F92
            Key             =   "Hermit #1"
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":29BE4
            Key             =   "KittyKat"
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2A836
            Key             =   "Stephen"
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2B088
            Key             =   "Tough Rat"
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2B8DA
            Key             =   "W C Rabbit"
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2C12C
            Key             =   "Yogi Rat"
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2C97E
            Key             =   "Flame Warpher"
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2D1D0
            Key             =   "MegaMighty"
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2DE22
            Key             =   "MegaSailor"
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2EA74
            Key             =   "Bun'rab'"
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":2F6C6
            Key             =   "Kat2"
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":30318
            Key             =   "ProfAllTiedUp"
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":30F6A
            Key             =   "MineyAllTiedUp"
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":311FC
            Key             =   "Lady Bug"
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":31E4E
            Key             =   "Bad Spider"
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":32AA0
            Key             =   "Live Flower"
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":336F2
            Key             =   "Giant Roach"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlTerrain 
      Left            =   3600
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   85
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":34344
            Key             =   "Bush"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3465E
            Key             =   "Cave Floor"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":34978
            Key             =   "Pool"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":34C92
            Key             =   "Cave Wall"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":34FAC
            Key             =   "Fire"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":352C6
            Key             =   "Forest"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":355E0
            Key             =   "Lawn"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":358FA
            Key             =   "Gravel"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":35C14
            Key             =   "House"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":35F2E
            Key             =   "Mountain"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":36248
            Key             =   "Stalagmite"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":36562
            Key             =   "Fruit Tree"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3687C
            Key             =   "Water"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":36B96
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":377E8
            Key             =   "Brick Wall"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3801A
            Key             =   "Carpet"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3886C
            Key             =   "Door"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":390BE
            Key             =   "Windowed Door"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":39910
            Key             =   "Cobbles"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3A562
            Key             =   "Krops"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3B1B4
            Key             =   "Dead Krops"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3BE06
            Key             =   "CFlower"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3CA58
            Key             =   "MFlower"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3D6AA
            Key             =   "KFlower"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3E2FC
            Key             =   "SeaSandUp"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3EF4E
            Key             =   "SeaSandLt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":3FBA0
            Key             =   "SeaSandRt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":407F2
            Key             =   "SeaSandDn"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":41444
            Key             =   "Sand"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":42096
            Key             =   "Sea"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":42CE8
            Key             =   "Tile"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4393A
            Key             =   "Dirt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4458C
            Key             =   "Dandelions"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":451DE
            Key             =   "Grass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":45E30
            Key             =   "Tracks Left"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":46A82
            Key             =   "Tracks right"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":476D4
            Key             =   "Tracks up"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":48326
            Key             =   "Tracks down"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":48F78
            Key             =   "Stone Walk"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":49BCA
            Key             =   "Signature"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4A81C
            Key             =   "Leafy Bush"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4B46E
            Key             =   "Berry Bush"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4C0C0
            Key             =   "Boring Grass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4CD12
            Key             =   "Pomarbo"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4D964
            Key             =   "FF1TreeBot"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4E5B6
            Key             =   "FF1TreeTop"
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4F208
            Key             =   "FirBot"
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":4FE5A
            Key             =   "FirMiddle"
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":50AAC
            Key             =   "FirTop"
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":516FE
            Key             =   "FirLBot"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":52350
            Key             =   "FirLEdge"
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":52FA2
            Key             =   "FirLEdgeTop"
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":53BF4
            Key             =   "FirRBot"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":54846
            Key             =   "FirREdge"
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":55498
            Key             =   "FirREdgeTop"
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":560EA
            Key             =   "FirMidLBot"
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":56D3C
            Key             =   "FirMidRBot"
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5798E
            Key             =   "FirMidLTop"
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":585E0
            Key             =   "FirMidRTop"
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":59232
            Key             =   "ElevatorL"
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":59A84
            Key             =   "ElevatorR"
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5A2D6
            Key             =   "Goodgrass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5AF28
            Key             =   "Small Karrots"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5BB7A
            Key             =   "CSand BL"
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5C7CC
            Key             =   "CSand BR"
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5D41E
            Key             =   "CSand TL"
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5E070
            Key             =   "CSand TR"
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5ECC2
            Key             =   "Stonewall"
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":5F914
            Key             =   "CSand InvBL"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":60566
            Key             =   "CSand InvBR"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":611B8
            Key             =   "CSand InvTL"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":61E0A
            Key             =   "CSand Inv TR"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":62A5C
            Key             =   "Cement"
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":636AE
            Key             =   "CementEngraved"
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":64300
            Key             =   "CementMsg"
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":64F52
            Key             =   "CementWriting"
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":65BA4
            Key             =   "CementSt"
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":667F6
            Key             =   "CementLtSt"
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":67448
            Key             =   "CementTTT"
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6809A
            Key             =   "CementCracked"
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":68CEC
            Key             =   "CementCracked2"
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6993E
            Key             =   "CementCracked3"
         EndProperty
         BeginProperty ListImage83 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6A590
            Key             =   "CCement"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage84 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6B1E2
            Key             =   "CCementEngr"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage85 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Game.frx":6BE34
            Key             =   "CCementSt"
            Object.Tag             =   "C"
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
'Note: The in-game comments are usually about obvious things but I like to write comments because it proves my superiority in English.
'That's why there are lots of them. I also like to type so the two create lots of comments in my code.
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


'   11. OK, this is a proposed method for getting onto a ship, car, etc. and using the normal navigation keys to move it around, while hiding
'the player, thus simulating the player getting on a vehicle and then riding around on it.
'first off, we'll implement this by creating a whole bunch of new MOVEMENT types. Each one would be named SHIP, CAR or something.
'Then we create in the
'Type Char '(that's You)
'   State as Integer
'   :
'   :
'End Type
'The .State would indicate what the player is doing at the time...right now we only care about using it to see if the player is RIDING, but
'it could later be used to indicate all sorts of useful stuff. That's why I decided on making it an int.
'Anyway, to start RIDING, the player uses the vehicle. In the use menu code, we flip his state to RIDING. This means in the Paint sub, we
'don't render him.*Code to here*  Second, we set his X,Y equal to the vehicle's X,Y. Next, whenever the player presses a key, it sets off a series of events
'that include moving all the objects on the map array. So when a vehicle is moved, it checks to see if the player is RIDING.
'If he is, then it sets its X,Y to the players. Three issues arise here.
'1) If there's more than one vehicle on the array, both will
'suddenly appear at the X,Y of the player.
'2)moving the vehicle across screen boundaries involves a little extra work, but not a whole lot.
'3) Each vehicle will have a set off tile(s) that they can travel across. Each time it moves while the player is RIDING, it needs to
'check to see if the new position is on a travelable tile type. If not, then we turn RIDING
'off, and stop moving the vehicle as well.
'That's it(as soon as I figure out how to fix that multiple vehicle bug. The way might be to make yet another property for Char and have it
'point to the Object array number of the vehicle being ridden. Then we just check it in the moveobject code to make sure he's riding the
'right vehicle.)
'   12. This is it. I'm about to make the big jump to having a timer event control the movement of the Things. This is accomplish by commenting
'the MoveThings out in MoveViewport and putting
'   MoveThings
'   PaintViewport
'inside the timer. Of course, we'll have to pause the timer anytime the player 1)presses a game key(u,d,x,l, etc.) or 2)right-clicks the picbox;
'then we'll unpause it at the end of every mnuUse/Drop etc. function when we're done with whatever it is. I'll write more after it's implemented.
'(06/03/1999)   13.Added a 'backbuffer' picture box to remove all(almost) flickering. The picture was starting to flicker badly when it was
'updated every 500 ms. I had to uncomment the BitBlt definition and I found out the hard way that the PicBox's AutoRedraw property must
'be trueif you want the picture to be hidden :). By the way, I tried implementing this in MapEdit as well, but I messed up some and got rattled
'and it didn't reduce the flickering a whole lot anyway, so I gave up on it. Oh well. *Maybe* later.
'   14. Added weight handling. Now everything has a weight and you can only pick up 250 'weights'
'at a time(this could easily change). The ASCII value of the first character in Thing.Tag holds the weight.
'0 for weight means that the player cannot pick the item up. The limit for a weight for a particular item is 255--the range of a Byte which
'is really what a 1 length string is...
'Note, however that the constants for x+, etc. are spelled out for InterpretStoredScript. In addition, I'm thinking about re-tooling some of these
'as Enums...the ones that only can be useda certain way.(i.e. CharState_riding, normal, whatever)
'   15. Yes, finally I have implemented a save/restore method: SaveState & LoadState. They work--sometimes! I still am not sure what's
'happening...it seems not to work when your position is X > 30...hmmm. well, I'll look at the code again for typos.
'Later: I found out what was happening; the code the figured the ScreenX,Y was buggy--it used your X,Y rather than your X,Y - (the screen
'tile size)[i.e. 10]. So the TopX,Y would frequently be < 0 thereby causing PaintViewport to load a *new* section of map on top
'of the currently loaded one...it was a mess but should be fixed now.
'anyway, as a side note(game map specific), I have found that we need another layer of cement wall around the edge of the level before it'll stop
'crashing :( more work for me.
'*** Constants ***
Private Const CHAR_MAXPOSSESSIONS As Integer = 50
Private Const CHAR_MAXWEIGHTCARRYING As Integer = 250

Private Const CHAR_MAXXRANGE As Integer = 7
Private Const CHAR_MAXYRANGE As Integer = 7
Private Const CHAR_MINXRANGE As Integer = 3
Private Const CHAR_MINYRANGE As Integer = 3

Private Const SCRIPT_BYE As Integer = 6

Private Const LVW_DESC As Integer = 1
Private Const LVW_MOVEMENT As Integer = 2
Private Const LVW_WEIGHT As Integer = 3

Enum SelectState
    UNSELECTED = 0 'means that no'thing' is selected right now
    Selected = 1 'means that the user has selected an X,Y and filled SelectX,Y with them.
    DROPPING = 2
    EXAMINING = 3
    GETTING = 4
    TALKING = 5
    USING = 6
    WHATSTHIS = 7
End Enum

Public Enum CharState 'public is default but I typed it anyway.
    Normal = 0
    Riding = 1
End Enum
'stored script constants
Private Const MAXSTOREDCOMM = 99
'Enum StoredScriptType   'to tell the InterpretStoredScript function what type we're passing WARNING: You cannot save enums
'within types to disk so there goes our enum! Oh, well. they will have to be constants again.
Private Const NoAction = 0
Private Const Remove = 1
Private Const Putted = 2
Private Const ChMap = 3
Private Const Chat = 4
Private Const ChThread = 5
Private Const Sleep = 6
Private Const Give = 7
Private Const Take = 8
Private Const Warp = 9
Private Const XPlus = 10
Private Const XMinus = 11
Private Const YPlus = 12
Private Const YMinus = 13      'woo-ooo! unlucky!
'End Enum
'talking and storyboard constants
Private Const NUMTHREADS As Integer = 1 '(0 to 1)
Private Const NUMHEADINGS As Integer = 2 '(0 to 2)
Private Const INITIALSCRIPTSIZE = 2000 '(0 to 2000) this is 2000 right now, but could change in the future if necessary.
Private Enum ScreenMoveType
'Const STILL = 0    'Note that this is already defined inside Packrat.bas so it is redundant to define it twice
     Up = 1    'all these stupid states are good preparation for animation in V2.0
     Down = 2  'plus they're useful now for tracking the mouse.
     Lft = 3       'these two stupid constants misspelled because of naming conflicts
     Rght = 4      'with the Ridiculous®™ functions already provided by VB.
     UpLeft = 5
     UpRight = 6
     DownLeft = 7
     DownRight = 8
End Enum
    'that's all for now, folks.
' *** End Constants ***
'The player's type.
Private Type Char
    x As Integer   'this is figured JUST LIKE CellX, CellY
    Y As Integer
    Possessions(0 To CHAR_MAXPOSSESSIONS - 1) As Thing
    State As CharState
    ThingRef As Integer  'right now the only use for this is to keep track of the vehicle we're riding.
    WeightCarrying As Integer   'how much you're carrying--up to 250!!
    'Health as byte <- the byte is wish full thinking
    'exp as integer
    ':
    ':
    'Whatever else Casey wants to put in here.
End Type

'a type that will store all stored scripting commands so that they can be stored at a future time and restored when
'we change screens.
Private Type StoredScript   'this struct is the exact copy of StoredScript except for the ScriptType extra...well, I don't know how to
'inherit structures from other structures in VB so here it is, a copycat struct, not a child struct.
    ScriptType As Integer 'tells the command that should be executed on this 'Thing'
    
        'after this StoredScript mostly matches struct 'Thing'..........oops! I mean Type Thing.
    Desc As Integer
        'NOTE: if you do not use X,Y with the InterpretScriptCommand() fill them with NOT_GIVEN(just in case)
        'this is mainly because (currently the remove command) the command may be able to use the X,Y optionally
        'and won't be able to tell your empty values from values of (0,0)
    x As Integer
    Y As Integer
    Movement As Byte
    Type As Integer
    Tag As String * 4
End Type
'Stored Script Declares
Dim StoredScriptCommands(0 To MAXSTOREDCOMM) As StoredScript

'Player position Declares
Dim You As Char
Dim ScreenX As Long
Dim ScreenY As Long
Dim TopX As Integer
Dim TopY As Integer

'Selection Declares
'Dim bBlocking As Boolean    'so that you can drag the mouse to paint
Dim wSelect As SelectState 'a flag that tells us what we're doing with the 'selection' of a 'thing'
Dim SelectX As Integer, SelectY As Integer 'whats the square that the user has selected to
'work with for an object?
'Filename declares(more inside individual functions that open and close particular files within one function call.)
Dim Fileno As Integer
Dim ObjFileno As Integer    'this is the File of the 'Things' file: currently 65% of the
'size of the map file, but that will change if we change the structure of 'Thing'.

'Animation Declares
Dim MoveState As ScreenMoveType    'this will keep track of how the player is moving
'Dim bToolTips As Boolean    'this is to keep track of whether the user wants tooltips or not(currently unused)

'Object declares
Dim Description() As String    'this holds all the descriptions of the different 'Things'

'Talking and storyboard declares
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
Private Sub LoadState(strStateFilename As String)
'strStateFilename should be the first part only of the file--not the .gsv part. that's assumed. I'm assuming that long filenames are fine&dandy
'so go ahead.
'function note: I store most things that are trivial(i.e. just one variable used per StoreScript struct) in the X member.
Dim stoTemp As StoredScript 'these two structs are identical, BTW except for the fact that StoredScript has 1 extra member.
Dim intSaveFileNo As Integer
Dim thiTemp As Thing
Dim i As Long   'long is just in case and under Win95(at least on a PII and I assume any Pentium) they really *are* faster.
    'open the file
    strStateFilename = Left(Opener, Len(Opener) - 3) & "gsv"    'gsv' stands for Game SaVe...it's all a big coincidence that it's the same
    intSaveFileNo = FreeFile                                                        'that it's the same as a Genecyst save state ;)
    Open strStateFilename For Random As #intSaveFileNo Len = Len(stoTemp)
    
    Get #intSaveFileNo, 1, stoTemp  'get the bulk of the info on You
    You.x = stoTemp.x
    You.Y = stoTemp.Y
    'now that we have X&Y we need to fig the screenX,Y and TopX,Y(c&p from InterpretScriptLine I think)
    ScreenX = ((You.x - MAP_SCREENX - 1) \ MAP_SCREENX) * MAP_SCREENX 'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
    ScreenY = ((You.Y - MAP_SCREENY - 1) \ MAP_SCREENY) * MAP_SCREENY 'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
    TopX = You.x - ScreenX - 4
    TopY = You.Y - ScreenY - 4  'I hope that these 4 lines of code work...I'm just winging it with no examples to prove my point.
    You.ThingRef = stoTemp.Desc
    You.State = stoTemp.Type    'here's the risky part--I won't know if this works until after I test it.
    
    Get #intSaveFileNo, , stoTemp   'now get the rest of your info
    You.WeightCarrying = stoTemp.x
    'now we get the threads from the file--note that we don't clear it out first since they're just simple ints.(unlike the next two sections)
    For i = 0 To NUMTHREADS Step 1
        Get #intSaveFileNo, , stoTemp
        Threads(i) = stoTemp.x
    Next i
    'now for the next two sections(possessions & remote script commands stack) we have to clear the data out first.
    lvwPossessions.ListItems.Clear  'clear the listview first
    For i = 0 To CHAR_MAXPOSSESSIONS - 1 Step 1
        thiTemp.Desc = NONE
        thiTemp.Movement = 0
        thiTemp.Type = NONE
        thiTemp.x = NONE
        thiTemp.Y = NONE
        You.Possessions(i) = thiTemp
    Next i
'now give to you all the things that are saved.
    For i = 0 To CHAR_MAXPOSSESSIONS - 1 Step 1
        Get #intSaveFileNo, , stoTemp
        thiTemp.Desc = stoTemp.Desc
        thiTemp.Movement = stoTemp.Movement
        thiTemp.Type = stoTemp.Type
        thiTemp.x = stoTemp.x
        thiTemp.Y = stoTemp.Y
        thiTemp.Tag = stoTemp.Tag
   Next i
'now to clear the SavedScriptCommands
    stoTemp.x = NONE
    stoTemp.Y = NONE
    stoTemp.Desc = NONE
    stoTemp.Movement = STILL
    stoTemp.ScriptType = NoAction
    stoTemp.Tag = ""
    stoTemp.Type = NONE
    For i = 0 To MAXSTOREDCOMM Step 1
        'now clear out the current stored script array with the empty StoredScript
        StoredScriptCommands(i) = stoTemp
    Next i
    'now that it's clear let's read in the new ones. Then we'll run RestoreScriptCommands just in case.
    For i = 0 To MAXSTOREDCOMM Step 1
        Get #intSaveFileNo, , stoTemp
        StoredScriptCommands(i) = stoTemp   'the ones that don't have anything in them *should* be empty :) well, really they will be because
        'they will be filled with empty structures when they're written to disk.
    Next i
    'here goes nothing(i.e. RestoreScriptCommands)
    RestoreScriptCommands   'no args(yes, this is a gratituous comment) wa ha ha read every comment in my code and you will go mad!!
    'now we should be done...I hope.
    SelectX = NONE  'o yeah. These lines R from Form_Load() buy they're useful so I'll keep them anyway.
    SelectY = NONE
    
    LoadMap Fileno, ScreenX, ScreenY 'call all of the load functions    'ummm...I knew there was something I was forgetting.
    LoadThings ObjFileno, ScreenX, ScreenY
    Close intSaveFileNo 'close the save file.
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
            'now if your thingref is pointing to something then make sure we refresh it to--else it points to the wrong thing!
            If You.ThingRef <> NONE Then IsThing You.x, You.Y, You.ThingRef
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
            If You.ThingRef <> NONE Then IsThing You.x, You.Y, You.ThingRef
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
            If You.ThingRef <> NONE Then IsThing You.x, You.Y, You.ThingRef
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
            If You.ThingRef <> NONE Then IsThing You.x, You.Y, You.ThingRef
'        Else
'            If bSaved = False Then
''               SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
'                SaveThings ObjFileno, ScreenX, ScreenY
'                bSaved = True
'                RestoreScriptCommands
'            End If
'        End If
    End If
'*** new: I'm changing the code here to paint to the 'back-buffer' picturebox--hopefully this will reduce flickering
    PaintMapFast picBackBuffer, imlTerrain, TopX, TopY
    
    'now paint the objects
    PaintThings picBackBuffer, imlThings, TopX, TopY, ScreenX, ScreenY 'paint the 'Things' onto picBackBuffer as
    'well.
    
    'now paint 'you' on screen.
    If You.State = Normal Then  'if you're riding (or something else later) don't paint you to the screen.
        imlThings.ListImages("Professor").Draw picBackBuffer.hDC, (You.x - ScreenX - TopX) * MAP_TILEXSIZE, (You.Y - ScreenY - TopY) * MAP_TILEYSIZE, imlTransparent
    End If
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
        DrawHighLight picBackBuffer, (SelectX - TopX - ScreenX) * MAP_TILEXSIZE, (SelectY - TopY - ScreenY) * MAP_TILEYSIZE, _
        MAP_TILEXSIZE, MAP_TILEYSIZE
    End If
    'now draw the completed image to the picViewport in one fell swoop.
    '(using BitBlt since the other method didn't work like I thought it would)
    BitBlt picViewport.hDC, 0, 0, picViewport.Width, picViewport.Height, picBackBuffer.hDC, 0, 0, SRCCOPY
    'tell the user his position(for debugging purposes currently)
    lblPosition.Caption = "X: " & You.x & " Y: " & You.Y '& " TopX = " & TopX & " TopY = " & TopY
    'don't need TopX,Y info any more.(but for heavy duty debugging uncomment it)
End Sub

Private Sub MoveViewport()
Dim iOldYouX As Integer
Dim iOldYouY As Integer
Dim iThingNum As Integer
    iOldYouX = You.x 'this so it is very easy to restore your settings if you walked through walls.
    iOldYouY = You.Y
    Select Case MoveState   'move 'you' in the appropriate direction
        Case STILL
            Exit Sub
        Case Up
            You.Y = You.Y - 1
        Case Down
            You.Y = You.Y + 1
        Case Lft
            You.x = You.x - 1
        Case Rght
            You.x = You.x + 1
        Case UpLeft
            You.x = You.x - 1
            You.Y = You.Y - 1
        Case UpRight
            You.x = You.x + 1
            You.Y = You.Y - 1
        Case DownLeft
            You.x = You.x - 1
            You.Y = You.Y + 1
        Case DownRight
            You.x = You.x + 1
            You.Y = You.Y + 1
    End Select
    If You.State = Normal Then  'otherwise the clipping is custom--right now the only alternative is Riding, where the clipping is handled
    'by the vehicle itself(the clipping is more permissive, that's why it's easier to handle)
        If imlTerrain.ListImages(Map(((You.Y - ScreenY) * MAP_ARRAYX) + (You.x - ScreenX))).Tag = "" Then
            'oops, we hit a solid rock.
            You.x = iOldYouX
            You.Y = iOldYouY
            MoveState = STILL
        ElseIf IsThing(You.x, You.Y, iThingNum) Then
            'we bumped into an irate person. here we should init battles, maybe move people out of your
            'way, and maybe move objects out of your way...But for now, we'll just stop you flat in your
            'tracks
            You.x = iOldYouX
            You.Y = iOldYouY
            MoveState = STILL
        End If
    Else    'right now only other state is You.Riding...so that's what this else is.
        MoveThings  'we need to call this when you're riding or else the ship gets left behind.
    End If
        'now that we've moved and clipped you, we need to move the Things, AND clip them.
'    MoveThings 'this is not called from here anymore; it's controlled by tmrMoveThings.
    
    If (You.x - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
        TopX = TopX + 1                     'screen a little
        If TopX = (MAP_ARRAYX - MAP_SCREENX) + 1 Then 'we're on the edge of the map, so bounce the player back
            TopX = MAP_ARRAYX - MAP_SCREENX
            You.x = You.x - 1
        End If
    ElseIf (You.x - TopX - ScreenX) < CHAR_MINXRANGE Then
        TopX = TopX - 1
        If TopX = -1 Then
            TopX = 0
            You.x = You.x + 1
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
            If Things(i).x > NONE Then    'make sure we've got a valid 'Thing'
            newX = Things(i).x
            newY = Things(i).Y
            Select Case Things(i).Movement
                Case RANDOM 'here we generate a random movement
                    Direction = CInt(Rnd * 8)   '0 to 8(i think)
                    'note:0 to 8 are based on the 8 movement constants used below
                    Select Case Direction
                        Case STILL
                            GoTo Continue
                        Case Up
                            newY = Things(i).Y - 1
                        Case Down
                            newY = Things(i).Y + 1
                        Case Lft   'i hope this is the right value NEW:no it's not! I've had to misspell both LFT and RGHT because of naming
                            newX = Things(i).x - 1      'conflicts. It was bugging the movement up and Stephen found it.
                        Case Rght
                            newX = Things(i).x + 1
                        Case UpLeft
                            newY = Things(i).Y - 1
                            newX = Things(i).x - 1
                        Case UpRight
                            newY = Things(i).Y - 1
                            newX = Things(i).x + 1
                        Case DownLeft
                            newY = Things(i).Y + 1
                            newX = Things(i).x - 1
                        Case DownRight
                            newY = Things(i).Y + 1
                            newX = Things(i).x + 1
                    End Select
                Case FOLLOW 'here they try to follow you until they run into the edge of the screen.
                    If You.x < Things(i).x Then
                        newX = Things(i).x - 1
                    ElseIf You.x > Things(i).x Then
                        newX = Things(i).x + 1
                    End If
                    If You.Y < Things(i).Y Then
                        newY = Things(i).Y - 1
                    ElseIf You.Y > Things(i).Y Then
                        newY = Things(i).Y + 1
                    End If
                Case ESCAPE 'opposite of FOLLOW
                    'NOTE: This untested as yet. However, as the FOLLOW works perfectly I expect no problems.
                    'Later:this has been tested but there seems to be a very small inconsistency in the way it works...not worth worrying about
                    'because its extremely trivial. In fact, it may not strictly exist ??... :)
                    If You.x < Things(i).x Then
                        newX = Things(i).x + 1
                    ElseIf You.x > Things(i).x Then
                        newX = Things(i).x - 1
                    End If
                    If You.Y < Things(i).Y Then
                        newY = Things(i).Y + 1
                    ElseIf You.x > Things(i).Y Then
                        newY = Things(i).Y - 1
                    End If
                Case SHIP
                'note that there is code in PaintViewport that regrabs the correct ThingRef when the screen changes--this code may now be
                'unworking since I moved the call of MoveThings to a timer...I think I'll have to call MoveThings every time that player moves
                'now so that the ship will keep up with him. That way the code will not get confused or behind on moving.
                    If You.State = Riding And You.ThingRef = i Then '1)the player is riding 2)the player is riding THIS object
                    'if you can't satisfy these requirements, then don't do anything(sim STILL)
                        'anyway, now we check to make sure that the player is still on SEA.(30)
                        If imlTerrain.ListImages(Map(((You.Y - ScreenY) * MAP_ARRAYX) + (You.x - ScreenX))).Index <> 30 Then
                            'now switch state back to normal
                            You.State = Normal
                            You.ThingRef = NONE
                            GoTo Continue   'skip to the next object.
                        End If
                        'now move us since we know that we're still sailing around happily.
                        If MoveThing(i, You.x, You.Y) = False Then  'oh, no, we're crossing a screen boundary meaning we have to do the move
                            'code ourselves(this code c&p from XPLUS, etc. code in InterpretStoredScript)
Dim Temp As Thing   'what does it look like?
                            Temp = Things(i)    'store it (lucky us; VB doesn't require a overloaded = operator for Types)
                            Temp.x = You.x
                            Temp.Y = You.Y
                            'next delete it from the current screen.
                            RemoveThingArray i
                            If PutThing(Temp, ScreenX, ScreenY) = False Then    'check to make sure that the thing has space to be put!
                                MsgBox "Screen Full. Ship capsized. Game over."
                            Else
                                'now we have to get the ship's new array number.(You.ThingRef is passed so the function will fill it with the new number)
                                IsThing Temp.x, Temp.Y, You.ThingRef
                            End If
                        End If
                    End If
                    GoTo Continue
                Case Else
                    GoTo Continue   'this hack works like the C++ statement 'continue' which
                    'VB carelessly never implemented.
            End Select
            'now clip the object--whoops! this is much later, but I recently reordered the If block for better performance...however, I have
            'found the hard way that If(newX - ScreenX) = -1 etc. test must be first or the other tests can crash.
            If (newX - ScreenX) = -1 Or (newY - ScreenY) = -1 Or (newX - ScreenX) = 30 Or (newY - ScreenY) = 30 Then
                GoTo Continue 'the continue; hack
            ElseIf imlTerrain.ListImages(Map(((newY - ScreenY) * 30) + (newX - ScreenX))).Tag = "" Then  'the 'Thing' hit a solid tile.
                GoTo Continue   'the continue; hack
            ElseIf IsThingExclude(newX, newY, i) = True Then    'you bumped into another 'Thing'
                GoTo Continue 'the continue; hack
            ElseIf newX = You.x And newY = You.Y Then  'the player is already here...Don't get in his way!!
                GoTo Continue   'the continue; hack
            'now if the 'Thing' survived the clipping, then we'll move it. But MoveThing has inherent screen clipping, so
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
            MoveState = Lft
        Case vbKeyDown
            MoveState = Down
        Case vbKeyUp
            MoveState = Up
        Case vbKeyRight
            MoveState = Rght
        Case vbKeyNumpad2
            MoveState = Down
        Case vbKeyNumpad4
            MoveState = Lft
        Case vbKeyNumpad6
            MoveState = Rght
        Case vbKeyNumpad8
            MoveState = Up
        Case vbKeyEnd
            MoveState = DownLeft
        Case vbKeyPageDown
            MoveState = DownRight
        Case vbKeyHome
            MoveState = UpLeft
        Case vbKeyPageUp
            MoveState = UpRight
        Case vbKeyNumpad1
            MoveState = DownLeft
        Case vbKeyNumpad3
            MoveState = DownRight
        Case vbKeyNumpad7
            MoveState = UpLeft
        Case vbKeyNumpad9
            MoveState = UpRight
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
                tmrMoveThings.Enabled = True
                PaintViewport
            Else    'maybe they're trying to quit
                'so unload the form in the standard way if they do, in fact, want to quit.
                If vbYes = MsgBox("Are you sure you want to quit?", vbYesNo, "Pressed ESC") Then Form_Unload 0
            End If
        Case vbKeyD 'drop
            tmrMoveThings.Enabled = False
            wSelect = DROPPING
            SelectX = You.x
            SelectY = You.Y
        Case vbKeyG 'get
            tmrMoveThings.Enabled = False
            wSelect = GETTING
            SelectX = You.x
            SelectY = You.Y
        Case vbKeyE, vbKeyX 'both E and X should work for EXamine
            tmrMoveThings.Enabled = False
            wSelect = EXAMINING
            SelectX = You.x
            SelectY = You.Y
        Case vbKeyT 'talk
            tmrMoveThings.Enabled = False
            wSelect = TALKING
            SelectX = You.x
            SelectY = You.Y
        Case vbKeyU 'use
            tmrMoveThings.Enabled = False
            wSelect = USING
            SelectX = You.x
            SelectY = You.Y
        Case vbKeyW, vbKeyL 'what's this should also work for What's this AND Look
            tmrMoveThings.Enabled = False
            wSelect = WHATSTHIS
            SelectX = You.x
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
                Case Up
                    SelectY = SelectY - 1
                Case Down
                    SelectY = SelectY + 1
                Case Lft
                    SelectX = SelectX - 1
                Case Rght
                    SelectX = SelectX + 1
                Case UpLeft
                    SelectX = SelectX - 1
                    SelectY = SelectY - 1
                Case UpRight
                    SelectX = SelectX + 1
                    SelectY = SelectY - 1
                Case DownLeft
                    SelectX = SelectX - 1
                    SelectY = SelectY + 1
                Case DownRight
                    SelectX = SelectX + 1
                    SelectY = SelectY + 1
            End Select

            If wSelect = GETTING Or wSelect = USING Then    'don't let the cursor get more than one space away.
                If Abs(You.x - SelectX) > 1 Or Abs(You.Y - SelectY) > 1 Then SelectX = OldX: SelectY = OldY
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
    You.x = TopX + 4
    You.Y = TopY + 4
    You.State = Normal
    You.ThingRef = NONE
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
    DummyThing.x = NONE
    DummyThing.Y = NONE
    'init your possessions to nothing
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
    LoadDescriptions    'this off of yet another open file...we just don't need the Fileno for later, therefore we don't have to pass it ByVal
    'like the others.
    
    'now init stored script commands array.
    For i = 0 To MAXSTOREDCOMM Step 1
        StoredScriptCommands(i).x = NONE
    Next i
    PaintViewport
    AdLabel 'generate free advertising for various copyrighted products.
    'check if there is a saved game; if so, does the player want to load it?
    If Dir(App.Path & "\" & "Hello.gsv") <> "" Then 'shockingly, this game will not work now if you install on the root folder!!(because of the presence
    'of the '\' on the root drives which I (again, shockingly) do not test for and correct.
    'now ask to make sure they want load the old game(they're stupid if they don't because the map/character movement changes are saved
    'inside the map independently of the save file ;)
        If MsgBox("Do you want to restore the saved game?", vbYesNo, "Game Restore") = vbYes Then
            LoadState "Hello" 'umm...is that it? I hope so.
            Exit Sub 'make sure to skip the intro. ha ha. no, really!
        End If
    End If
    'now play the intro:
    LoadTalkBox
    'these lines commented out so that the game 1. doesn't require a bmp file in the zip and 2. so that it doesn't take so
    'long to start the game... (NEW: Most of the lines now commented are from episode 1. The good ones are for episode 2.)
    SetTalkBoxFont "Arial", 24  'wish we could do 3d...
    OpenTalkBox "EPISODE II:" & vbCrLf & "LEVEL FIVE"
'    OpenTalkBox "The Adventures of You, Mikey, and Miney"
    SetTalkBoxFont "Arial", 14
'    SetTalkBoxBackGround "C:\My Documents\Visual Basic\Game Engine 1.0\green hills.bmp"
    OpenTalkBox "It is the eve of the great election!  After much campaigning, Mikey is assured of the office of mayor until Macro le Mouse, claiming to be his cousin, appears on the scene.  Now there are two candidates for mayor, and an argument is brewing between them...."
    OpenTalkBox "StoryLine: Nathan Sanders" & vbCrLf & "Rachel Sanders"
    OpenTalkBox "Programming: Nathan Sanders"
    OpenTalkBox "Art: Rachel Sanders"
'    OpenTalkBox "Mikey as Himself" 'OK, I admit it. This is stolen from the Muppets movies(but I won't tell you where any of the rest came from! (Maybe just from my tortured imagination. How would you know ha ha)
    SetTalkBoxFont "Times New Roman", 14    'set the game's standard talk font
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
    SaveState "Hello"   'save whether they want to or not!
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
            SelectX = You.x
            SelectY = You.Y
            picViewport.SetFocus
        Case vbKeySpace
            lvwPossessions_DblClick
    End Select
End Sub

Private Sub lvwPossessions_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbKeyRButton Then lvwPossessions_DblClick
End Sub

Private Sub mnuDrop_Click()
    'wSelect should be set to selected and selectx,y initialized for this function.
Dim Temp As Thing
Dim Count As Integer
    If lvwPossessions.ListItems.Count = 0 Then Exit Sub 'can't drop anything when you don't have anything to drop
    Count = CInt(Right(lvwPossessions.SelectedItem.Key, Len(lvwPossessions.SelectedItem.Key) - 1)) 'SelectedItem is set to the first one if none selected
    Temp = You.Possessions(Count)
    Temp.x = SelectX
    Temp.Y = SelectY
    
    If PutThing(Temp, ScreenX, ScreenY) = True Then
        TakeThing You, lvwPossessions, Count
    Else
        MsgBox "Screen Full! Try dropping over about 5 spaces."
    End If
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
    PaintViewport 'refresh screen. else you get confusing things like dropping something and not have it appear yet it's not your inventory
    'anymore and what are you to do?!?
    '.........
    'That's right!! Call for Zackman, super-hero from the far past!!!!!!!!
    tmrMoveThings.Enabled = True    'start everything moving again.
End Sub

Private Sub mnuExamine_Click()
Dim ArrayNum As Integer
Dim strWeightDesc As String
    IsThing SelectX, SelectY, ArrayNum 'this is only to get arraynum; we already know that there is something there
    If TypeOfThing(ArrayNum) = OBJ Then 'it needs to be an object before we can show a description, right?
        If Asc(Mid$(Things(ArrayNum).Tag, 1, 1)) > 0 Then
            strWeightDesc = "It weighs " & Asc(Mid$(Things(ArrayNum).Tag, 1, 1)) & " Weights."
        Else    'it cannot be picked up!
            strWeightDesc = "It cannot be picked up."
        End If
        MsgBox Description(Things(ArrayNum).Desc) & vbCrLf & strWeightDesc, vbInformation
    End If
    
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
    tmrMoveThings.Enabled = True    'start everything moving again.
End Sub

Private Sub mnuGet_Click()
Dim Temp As Thing
Dim ArrayNum As Integer
Dim i As Integer
    i = 0   'a simple increment variable. That is initialized to 0.
    If You.State = Normal Or SelectX <> You.x Or SelectY <> You.Y Then  'we have to make sure that you cannot pick up the vehicle
    'in which you are riding!!
        IsThing SelectX, SelectY, ArrayNum  'get what number this is
        If TypeOfThing(ArrayNum) = OBJ Then 'make sure you're trying to take objects!
            'save the value of the thing that you're taking
            Temp = Things(ArrayNum)
            'take the 'Thing' away from the map
            RemoveThingArray ArrayNum
    
            If GiveThing(You, lvwPossessions, Temp) = False Then
                MsgBox "Inventory full! or You can't pick this up!"
                PutThing Temp, ScreenX, ScreenY 'put it back in a hurry
            End If  'end if inventory full
        End If  'end if taking non-object
    End If  'end if getting vehicle currently being used
    wSelect = UNSELECTED
    SelectX = NONE
    SelectY = NONE
    PaintViewport   'refresh screen.
    tmrMoveThings.Enabled = True    'start everything moving again.
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
    tmrMoveThings.Enabled = True    'start everything moving again.
End Sub

Private Sub mnuUse_Click()
'now that the SelectX,Y has been processed properly, we finally let mnuUse do something with it.
Dim thiTemp As Thing    'just a holding spot for the 'Thing' in question.
Dim intArray As Integer 'note this variable is used DIFFERENTLY according to where the object is being used from. Warning!!
Dim i As Integer
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
            thiGivee.x = 0
            thiGivee.Y = 0
            For i = 0 To 2 Step 1
                If GiveThing(You, Me.lvwPossessions, thiGivee) = False Then
                    MsgBox "You don't have enough room to hold them, however."
                    Exit For
                End If
            Next
        Case 10 'a ballot box
            'put a talkbox here with two choices: "Mikey" and "Macky" (until we find a better name for him)
            'then set the thread to a certain number to change the storyline.
            'facxila, cxu ne?
        Case 11 'a automated storekeeper that sells
            'nothing here yet either, but we'd need more talkboxes and then a select case and a havething(for points) and then a givething
            'and then a takething(for points)(or whatever we decide to make Moneys).
        Case 12 'a automated storekeeper that buys
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
    'now for special actions: right now just start RIDING a SHIP
    If thiTemp.Movement = SHIP Then
        You.x = thiTemp.x
        You.Y = thiTemp.Y
        You.State = Riding  'very cool...I think I like Enums. Anyway, this means that we don't draw You.
        You.ThingRef = intArray 'this so we won't suddenly be riding two ships at once :)
        PaintViewport
    End If
    wSelect = UNSELECTED    'I love VB's autocaps :) (I have been working with DevStudio lately, or couldn't
    SelectX = NONE  'you tell?
    SelectY = NONE
    tmrMoveThings.Enabled = True    'start everything moving again.
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
    tmrMoveThings.Enabled = True    'start everything moving again.
End Sub

Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim XCell As Integer, YCell As Integer
Dim ArrayNum As Integer
Dim bResult As Boolean
    XCell = x \ MAP_TILEXSIZE
    YCell = Y \ MAP_TILEYSIZE
    Select Case Button
        Case vbKeyLButton   'check to see if we're selecting something. If not, ignore the click.
            If wSelect >= DROPPING Then
                bResult = IsThing(SelectX, SelectY, ArrayNum)
                Select Case wSelect
                    Case GETTING
                        If bResult = True And Abs(You.x - SelectX) < 2 And Abs(You.Y - SelectY) < 2 Then
                            If TypeOfThing(ArrayNum) = OBJ Then mnuGet_Click
                        End If
                    Case TALKING
                        If bResult = True Then
                            If TypeOfThing(ArrayNum) = PERSON Then mnuTalk_Click
                        End If
                    Case USING
                        If bResult = True And Abs(You.x - SelectX) < 2 And Abs(You.Y - SelectY) < 2 Then
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
            tmrMoveThings.Enabled = False
            wSelect = Selected  'set select status so that when we call the context menu
            SelectX = XCell + ScreenX + TopX    'functions they'll know what the player
            SelectY = YCell + ScreenY + TopY    'is pointing at.

            Result = IsThing(SelectX, SelectY, ArrayNum)
            'set or reset all the menu values(the If statement is to determine if the space is empty
            mnuSep1.Visible = False
            mnuTalk.Visible = False
            mnuGet.Visible = False
            If lvwPossessions.ListItems.Count = 0 Or Result = True Or imlTerrain.ListImages(Map(((SelectY - ScreenY) * MAP_ARRAYX + (SelectX - ScreenX)))).Tag = "" Then
            'if you don't have anything or the tile is solid(there should be an exception here for moving ships around though) or there's
            'somebody here already, don't let them drop anything
                mnuDrop.Visible = False
            Else
            'show Drop plus the 1st divider since it *is* possible to drop something here.
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
                    If Abs(You.x - SelectX) < 2 And Abs(You.Y - SelectY) < 2 Then
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

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim MoveStateX As Integer, MoveStateY As Integer
    If Button = vbKeyLButton And wSelect = UNSELECTED Then   'We're dragging. But are we within range?
        If x < (You.x - ScreenX - TopX) * MAP_TILEXSIZE Then 'we're dragging left
            MoveStateX = Lft
        ElseIf x > (You.x - ScreenX - TopX) * MAP_TILEXSIZE + MAP_TILEXSIZE Then    'we're dragging right
            MoveStateX = Rght
        End If
        If Y < (You.Y - ScreenY - TopY) * MAP_TILEYSIZE Then 'we're dragging up
            MoveStateY = Up
        ElseIf Y > (You.Y - ScreenY - TopY) * MAP_TILEYSIZE + MAP_TILEYSIZE Then 'we're dragging down
            MoveStateY = Down
        End If
        'now process the movestatex,y and combine them into one variable:MoveState
        If MoveStateX = Lft Then
            If MoveStateY = Up Then
                MoveState = UpLeft
            ElseIf MoveStateY = Down Then
                MoveState = DownLeft
            Else
                MoveState = Lft
            End If
        ElseIf MoveStateX = Rght Then
            If MoveStateY = Up Then
                MoveState = UpRight
            ElseIf MoveStateY = Down Then
                MoveState = DownRight
            Else
                MoveState = Rght
            End If
        Else        'no x movement, so we must just be doing simple Y movement, so just set movestate to movestatey
            MoveState = MoveStateY
        End If
        tmrMove.Enabled = True
    Else    'set the movestate to STILL so we'll stop moving
        If wSelect >= DROPPING Then
            SelectX = ScreenX + TopX + (x \ MAP_TILEXSIZE)
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
                        Temp.x = Temp.x + 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = Putted
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.x = Temp.x
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
                            Temp.x = Temp.x + 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                            Temp.x = Temp.x + 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                    'commands might not work. however, I will not deal with that just yet. However, I *do* hope that they all still work as
                    'it would be nice to be able to warp you 'inside' a hole and the script still continue executing(e.g. you're captured outside by
                    'rabbits and they take you inside and push you into a jail area)
                        If intX = 1 Then    'just one move--no sleeps either(cut&paste case 0)
                            'warning! we need to change the movething If to an If that tests if we went out of range of our little movement range(3-8 or something)
                            'and scroll the screen. Then we need an if that sees If we're crossing a screen boundary and that Loads/Saves Things() and Map().
                            You.x = You.x + 1
                            If (You.x - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
                                    TopX = TopX + 1                     'screen a little
                            End If
                            'finally, update the screen.
                            PaintViewport
                            'here is the if in question from above(the moving array causing bugs problem)
                            If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                                'this conversation is over
                                SelectX = NONE
                                SelectY = NONE
                                InterpretScriptLine = False
                                Exit Function
                            End If
                            
                        ElseIf intY = 0 Then    'more than one move, but no sleeps or screenpaints in between(cut&paste case 1)
                            For intStart = 1 To intX Step 1
                                You.x = You.x + 1
                                If (You.x - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopX = TopX + 1                     'screen a little
                                        If TopX = (MAP_ARRAYX - MAP_SCREENX) + 1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                                        'Things() arrays.
                                            TopX = MAP_ARRAYX - MAP_SCREENX
                                            You.x = You.x - 1 'end ***
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
                                You.x = You.x + 1   '***
                                If (You.x - TopX - ScreenX) > CHAR_MAXXRANGE Then    'we've gone out of our range of movement, so scroll the
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
                                Temp.x = Temp.x + 1
                                'next delete it from the current screen.
                                RemoveThingArray i
                                'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = Putted
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.x = Temp.x
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
                                    Temp.x = Temp.x + 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x + 1, Temp.Y, i
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
                                    Temp.x = Temp.x + 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x + 1, Temp.Y, i
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
                        Temp.x = Temp.x - 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = Putted
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.x = Temp.x
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
                            Temp.x = Temp.x - 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                            Temp.x = Temp.x - 1
                            'next delete it from the current screen.
                            RemoveThingArray i
                            'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                            You.x = You.x - 1
                            If (You.x - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement, so scroll the
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
                                You.x = You.x - 1
                                If (You.x - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement, so scroll the
                                        TopX = TopX - 1                     'screen a little
                                        If TopX = -1 Then 'this code signals PaintViewport that we need to scroll the Map() and
                                        'Things() arrays.
                                            TopX = 0
                                            You.x = You.x + 1
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
                                You.x = You.x - 1
                                If (You.x - TopX - ScreenX) < CHAR_MINXRANGE Then    'we've gone out of our range of movement, so scroll the
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
                                Temp.x = Temp.x - 1
                                'next delete it from the current screen.
                                RemoveThingArray i
                                'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = Putted
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.x = Temp.x
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
                                    Temp.x = Temp.x - 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x - 1, Temp.Y, i
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
                                    Temp.x = Temp.x - 1
                                    'next delete it from the current screen.
                                    RemoveThingArray i
                                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x - 1, Temp.Y, i
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
                        If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = Putted
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.x = Temp.x
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
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                                If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = Putted
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.x = Temp.x
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
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x, Temp.Y + 1, i
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
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x, Temp.Y + 1, i
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
                        If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            'SelectX,Y go out of focus and the conversation is over.
                            SelectX = NONE
                            SelectY = NONE
                            StoreScript.ScriptType = Putted
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.x = Temp.x
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
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                'SelectX,Y go out of focus and the conversation is over.
                                SelectX = NONE
                                SelectY = NONE
                                StoreScript.ScriptType = Putted
                                StoreScript.Desc = Temp.Desc
                                StoreScript.Movement = Temp.Movement
                                StoreScript.Type = Temp.Type
                                StoreScript.x = Temp.x
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
                                If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                    StoreScript.ScriptType = Putted
                                    StoreScript.Desc = Temp.Desc
                                    StoreScript.Movement = Temp.Movement
                                    StoreScript.Type = Temp.Type
                                    StoreScript.x = Temp.x
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
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x, Temp.Y - 1, i 'CHANGEIT
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
                                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                                        StoreScript.ScriptType = Putted
                                        StoreScript.Desc = Temp.Desc
                                        StoreScript.Movement = Temp.Movement
                                        StoreScript.Type = Temp.Type
                                        StoreScript.x = Temp.x
                                        StoreScript.Y = Temp.Y
                                        StoreScriptCommand StoreScript
                                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                                        MsgBox "Screen Full. Person that was walking is dead."
                                     End If
                                     'now re-get the person so we can keep looping
                                     i = 0  'just in case...this has been holding the index of another Thing in it, so better safe(and slow) than sorry
                                     IsThing Temp.x, Temp.Y - 1, i 'CHANGEIT
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
                    Temp.x = intX
                    Temp.Y = intY
                    'next delete it from the current screen.
                    RemoveThingArray i
                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                        StoreScript.ScriptType = Putted
                        StoreScript.Desc = Temp.Desc
                        StoreScript.Movement = Temp.Movement
                        StoreScript.Type = Temp.Type
                        StoreScript.x = Temp.x
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
                    You.x = intX
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
                        Temp.x = Temp.x + 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            StoreScript.ScriptType = Putted
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.x = Temp.x
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
            Temp.x = 0 'init the x,y to some innocuous but non-NULL value.
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
            Temp.x = intX
            Temp.Y = intY
            If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                'location out of range of Map(); store it for future use...
                StoreScript.ScriptType = Putted
                StoreScript.Desc = Temp.Desc
                StoreScript.Movement = Temp.Movement
                StoreScript.Type = Temp.Type
                StoreScript.x = intX
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
                    StoreScript.ScriptType = Remove
                    StoreScript.Desc = Temp.Desc
                    StoreScript.Type = Temp.Type
                    StoreScript.x = intX
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
                StoreScript.ScriptType = ChMap
                StoreScript.x = intX
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
        If StoredScriptCommands(i).x = NONE Then
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
        If (StoredScriptCommands(i).x > ScreenX) And (StoredScriptCommands(i).x < ScreenX + (MAP_ARRAYX - 1)) _
        And (StoredScriptCommands(i).Y > ScreenY) And (StoredScriptCommands(i).Y < ScreenY + (MAP_ARRAYY - 1)) Then   'it's a hit
            InterpretStoredScript StoredScriptCommands(i)
            StoredScriptCommands(i).x = NONE  'make sure we invalidate this record so that it can be overwritten and
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
        Case Chat    'a simple chatty OpenTalkBox(note: pass this command a value bigger than 0 in the X value to
            'let it know that you have already called LoadTalkBox
            '(Chat:Hello, my name is Mikey and I want you to vote for ME!)
            If StoreScript.x < 1 Then
                LoadTalkBox
                OpenTalkBox StoreScript.Tag
                UnloadTalkBox
            Else
                OpenTalkBox StoreScript.Tag
            End If
        Case Warp
'WARNING! WARNING! untested code here! it was just pasted in from the InterpretScriptLine function and has no gaurantees
'to work now that it's here.
            intX = StoreScript.x
            intY = StoreScript.Y
            'the choices for intType are: you, the 'Thing' passed by desc, and the 'Thing' passed by array number
            intType = StoreScript.Type  'this is NONE for U, NOT_GIVEN for RA val, and the desc for desc num.
            If StoreScript.Type = NOT_GIVEN Then   'default--by RA val
                i = StoreScript.Desc
                If MoveThing(i, intX, intY) = False Then
                    'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                    
                    'first save it to a Temp thing
                    Temp = Things(i)
                    Temp.x = intX
                    Temp.Y = intY
                    'next delete it from the current screen.
                    RemoveThingArray i
                    'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                    If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                        StoreScript.ScriptType = Putted
                        StoreScript.Desc = Temp.Desc
                        StoreScript.Movement = Temp.Movement
                        StoreScript.Type = Temp.Type
                        StoreScript.x = Temp.x
                        StoreScript.Y = Temp.Y
                        StoreScriptCommand StoreScript
                        'InterpretStoredScript = False  'not needed in a sub!!(InterpretScriptLine's a function, whichis why this line's still in here;
                        'it's been pasted from there.)
                        Exit Sub
                    ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                        MsgBox "Screen Full. Person that was warping is dead."
                     End If
                     'now we need to return the new values of the 'Thing' in Storscript.X,Y. This is so if an object is moved, it can update the SelectX,SelectY.
                     'Otherwise, the caller ignores the values
                     StoreScript.x = intX
                     StoreScript.Y = intY
                End If
                'now show that the person has moved
                PaintViewport
                
            ElseIf intType = NONE Then  'it's the player
                
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
'                    TopX = (intX - MAP_SCREENX / 2) - ScreenX
'                    TopY = (intY - MAP_SCREENY / 2) - ScreenY
                    You.x = intX
                    You.Y = intY
                    'finally, update the screen.
                    PaintViewport
                    If (SelectX < ScreenX) Or (SelectX > ScreenX + (MAP_ARRAYX - 1)) Or (SelectY < ScreenY) Or (SelectY > ScreenY + (MAP_ARRAYY - 1)) Then
                        'this conversation is over
                        SelectX = NONE
                        SelectY = NONE
                        Exit Sub
                    End If

            Else    'it's a 'Thing'
                    'note that once again, I've re-used some int variable names...
                    intStart = ((SelectX - ScreenX) \ MAP_SCREENX) * MAP_SCREENX   'this complex operation snaps the selectx to the nearest 0, 10, or 20 to pass to
                    'IsThingDesc.
                    intEnd = ((SelectY - ScreenY) \ MAP_SCREENX) * MAP_SCREENX    'same here
                    i = IsThingDesc(intDesc, intStart, intEnd)
                    If MoveThing(i, intX, intY) = False Then
                        'OK, we need to remove it and move it ourselves...if it's out of range of the array, remove and make a put equiv. to be stored at the new pos.
                                       
                        'first save it to a Temp thing
                        Temp = Things(i)
                        Temp.x = Temp.x + 1
                        'next delete it from the current screen.
                        RemoveThingArray i
                        'now see if it should be just put down on the next screen(still in the Map array) or stored to be put down later(going off the Map array)
                        If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYY - 1)) Then
                            StoreScript.ScriptType = Putted
                            StoreScript.Desc = Temp.Desc
                            StoreScript.Movement = Temp.Movement
                            StoreScript.Type = Temp.Type
                            StoreScript.x = Temp.x
                            StoreScript.Y = Temp.Y
                            StoreScriptCommand StoreScript
                        ElseIf PutThing(Temp, ScreenX, ScreenY) = False Then
                            MsgBox "Screen Full. Person that was walking is dead."
                        End If
                    End If
                    'finally, update the screen.
                    PaintViewport
                
                End If
'            End If

        Case Give    'Give the player something
            '(Give:1, 2, 3)
            '1 = Type( .type), 2 = Desc( .desc), 3 = movement (.movement)
            Temp.x = 0 'init the x,y to some innocuous but non-NULL value.
            Temp.Y = 0
            Temp.Type = StoreScript.Type
            Temp.Desc = StoreScript.Desc
            Temp.Movement = StoreScript.Movement
            'now call the function that 'gives' the 'Thing'
            If GiveThing(You, lvwPossessions, Temp) = False Then   'uh-oh, the player is out of room.
                'do nothing now, but maybe code some action here.
                MsgBox "Error: Player's possessions full!"
            End If
        Case Take    'taKe something away from the player
            '(Take: 1)   1 = description number of the object( .desc)

            intDesc = StoreScript.Desc
            If HaveThing(You, intDesc, ArrayNum) = True Then
                TakeThing You, lvwPossessions, ArrayNum
            End If
        Case Putted    'Put something on the world map
            '(Put:111, 999, 2, 3, 0)
            '111 = X position on the map
            '999 = Y position on the map
            '2 = the type(i.e. picture and 'Thing' type (.type)
            '3 = the description number( .desc)
            '0 = the movement value (.movement)
            Temp.Type = StoreScript.Type
            Temp.Desc = StoreScript.Desc
            Temp.Movement = StoreScript.Movement
            Temp.x = StoreScript.x
            Temp.Y = StoreScript.Y
            If (Temp.x < ScreenX) Or (Temp.x > ScreenX + (MAP_ARRAYX - 1)) Or (Temp.Y < ScreenY) Or (Temp.Y > ScreenY + (MAP_ARRAYX - 1)) Then
                'location out of range of Map(); store it for future use...
                'REPLACE WITH StoreScriptCommand(StoredCommand as StoredCommand)
                MsgBox "Warning! Check call stack and above If for possible bugs!"
                StoreScriptComm.ScriptType = Putted
                StoreScriptComm.Desc = Temp.Desc
                StoreScriptComm.Movement = Temp.Movement
                StoreScriptComm.Type = Temp.Type
                StoreScriptComm.x = intX
                StoreScriptComm.Y = intY
                StoreScriptCommand StoreScriptComm
            End If

            If PutThing(Temp, ScreenX, ScreenY) = False Then    'uh-oh: screen full. cancel changes, but do nothing
            'else(now currently anyway)
                MsgBox "Screen Full! Cannot put down another Thing"
            End If
            
        Case Remove    'remove something from the world map
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
            If StoreScript.x = NOT_GIVEN Then  'just a standard(no X,Y) remove command)
                intObjectScreen = (((You.Y - ScreenY) \ MAP_SCREENX) * MAP_ARRAYX) + (((You.x - ScreenX) \ MAP_SCREENX) * OBJ_MAXTHINGSSCREEN)
                'just jump down and check the current screen.(skip the precise X,Y check)
            Else
                intObjectScreen = (((StoreScript.Y - ScreenY) \ MAP_SCREENX) * MAP_ARRAYX) + (((StoreScript.x - ScreenX) \ MAP_SCREENX) * OBJ_MAXTHINGSSCREEN)
                intX = StoreScript.x
                intY = StoreScript.Y
                'make sure that the co-ordinates are in the range of Map(); if not, store using StoreScriptCommand
                If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                    'location out of range of Map(); store it for future use...
                    StoreScriptComm.ScriptType = Remove
                    StoreScriptComm.Desc = Temp.Desc
                    StoreScriptComm.Type = Temp.Type
                    StoreScriptComm.x = intX
                    StoreScriptComm.Y = intY
                    StoreScriptCommand StoreScriptComm
                End If
                'now check the X,Y position first
                If IsThing(intX, intY) Then RemoveThing intX, intY
                'then go through the rigamarole of checking everything again.
            End If
            For i = intObjectScreen To intObjectScreen + 9 Step 1 'loop through all the 'Things' on this particular screen.
                If Things(i).Desc = intDesc And Things(i).Type = intType And Things(i).x <> NONE Then   'a match!!
                    RemoveThingArray i
                    bFound = True
                    Exit For
                End If
            Next i
            If bFound = False Then   'check the whole array because we didn't find the object on the same screen
            'as the person being talked to.
                For i = 0 To (OBJ_MAXTHINGSARRAY - 1) Step 1 'the WHOLE object array.
                    If Things(i).Desc = intDesc And Things(i).Type = intType And Things(i).x <> NONE Then RemoveThingArray i
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
        Case ChMap
'            (ChMap:111, 999, 22)   'NOTE: X and Y START AT 0!!! This is VERY important!
'            111 = X pos(.x), 999 = y pos (.y), 22 = map type (.type)
            intX = StoreScript.x
            intY = StoreScript.Y
            If (intX < ScreenX) Or (intX > ScreenX + (MAP_ARRAYX - 1)) Or (intY < ScreenY) Or (intY > ScreenY + (MAP_ARRAYY - 1)) Then
                MsgBox "Warning: Check call stack and above if for possible bug!"
                Exit Sub
            End If
            intType = StoreScript.Type
            Map(((intY - ScreenY) * MAP_ARRAYX) + (intX - ScreenX)) = intType    'voila!, it's changed
'        Case "paint"
'            PaintViewport   'duh: this is simple(and unneeded: commented out for now)
        Case Sleep
            '(Sleep:250)
            '250 = number of ms to pause everything(useful perhaps for animation type stuff)(.x)
            '(i.e. OH NO!!! Not Final Fantasy movies.... Yaaahh!)
            'NOTE:Used intDesc here to avoid creating yet another variable
            '!!DESC HAS NOTHING TO DO WITH THE SLEEP COMMAND!!
            intDesc = StoreScript.x
            sngEndTime = Timer + (intDesc / 1000)
            Do Until Timer > sngEndTime
                DoEvents    'make sure we let Windows® 95™ continue doing its good work(spreading Microsoft everywhere)
                'although this could be dangerous if they used the arrow keys to scroll alot.
            Loop
        Case ChThread    'change a thread
            '(Thread:0, 1)
            ' 0 = number of thread to change(storescript.x), 1 = number to which to change the thread.(storescript.y)
            OldThread = StoreScript.x
            NewThread = StoreScript.Y
            Threads(OldThread) = NewThread  'this is Ryan's idea... a revised method.

    End Select
    
'Continue:       'this is a hack to emulate the 'break;' statement from C programming.(carelessly omitted by the very
'structure of 'Select Case') though in this function (excuse me) sub, an Exit Sub will do just as well as there is no
'critical incrementing and returning of values to be done.
End Sub
Private Function HaveThing(Ch As Char, Desc As Integer, Optional ByRef ArrayNum As Integer = NOT_GIVEN) As Boolean
    'ch is the person to search, search should contain the info on the Thing that we're
    'searching for; ArrayNum is the number that we put the actual number in if they want it.
Dim i As Integer    'a simple incrementing variable
Dim Found As Boolean
    For i = 0 To CHAR_MAXPOSSESSIONS - 1 Step 1
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
Dim intWeight As Integer
    Found = False   'init Found to be absolutely sure.
    'find a free space in the person's possessions
    For i = 0 To CHAR_MAXPOSSESSIONS Step 1
        If Ch.Possessions(i).x = NONE Then
            Found = True
            Exit For
        End If
    Next i
    If Found = True Then
        intWeight = Asc(Mid$(Givee.Tag, 1, 1))
        If intWeight = 0 Or intWeight + You.WeightCarrying > CHAR_MAXWEIGHTCARRYING + 1 Then 'immovable or too heavy!
            GiveThing = False   'bail out but quick.
            Exit Function
        End If

        'give the thing to you
        Ch.Possessions(i) = Givee
        Ch.WeightCarrying = Ch.WeightCarrying + intWeight
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
            ElseIf Ch.Possessions(i).Movement = ESCAPE Then
                itmX.SubItems(LVW_MOVEMENT) = "Escape"
            ElseIf Ch.Possessions(i).Movement = SHIP Then
                itmX.SubItems(LVW_MOVEMENT) = "Ship"
            Else    'give it a guess and say follow
                itmX.SubItems(LVW_MOVEMENT) = "Follow"
            End If
            itmX.SubItems(LVW_WEIGHT) = Asc(Mid$(Ch.Possessions(i).Tag, 1, 1))
        End With
    End If
    
    GiveThing = Found   'True or False, respectively
End Function
Private Sub TakeThing(Ch As Char, lvwCh As ListView, ArrayRemove As Integer)
    'Ch is the the person to take from. lvwCh is the listview to remove from, and Removenum is the number in the
    'possessions array to remove
        Ch.Possessions(ArrayRemove).x = NONE
        lvwCh.ListItems.Remove lvwCh.ListItems("#" & CStr(ArrayRemove)).Index
        Ch.WeightCarrying = Ch.WeightCarrying - Asc(Mid$(Ch.Possessions(ArrayRemove).Tag, 1, 1))
End Sub

Private Sub tmrMoveThings_Timer()
    MoveThings
    PaintViewport
End Sub

Private Sub SaveState(strStateFilename As String)
'strStateFilename should be the first part only of the file--not the .gsv part. that's assumed. I'm assuming that long filenames are fine&dandy
'so go ahead.
'function note: I store most things that are trivial(i.e. just one variable used per StoreScript struct) in the X member.
Dim stoTemp As StoredScript 'these two structs are identical, BTW except for the fact that StoredScript has 1 extra member.
Dim intSaveFileNo As Integer
Dim thiTemp As Thing
Dim i As Long   'long is just in case and under Win95(at least on a PII and I assume any Pentium) they really *are* faster.
    'open the file
    strStateFilename = Left(Opener, Len(Opener) - 3) & "gsv"    'gsv' stands for Game SaVe...it's all a big coincidence that it's the same
    intSaveFileNo = FreeFile                                                        'that it's the same as a Genecyst save state ;)
    Open strStateFilename For Random As #intSaveFileNo Len = Len(stoTemp)

    'first we c&p the code from form_load:
    stoTemp.x = You.x
    stoTemp.Y = You.Y
    stoTemp.Desc = You.ThingRef
    stoTemp.Type = You.State    'here's where we may have problems; these two used to be both ints(or maybe one was a byte)
                                                'anyway, now they are both enums and I'm not sure how much enums retain their 'intness' after being
                                                'reformed.
    'here we need another stoTemp structure or else we need to dump the first to disk and move on(i.e. overwrite the current one) because
    'because of disk constraints the StoreScript struct is rather small, containing large amounts of Bytes when the Char struct contains lots of
    'Ints. Actually right now the only left-over member of Char is WeightCarrying but it is an unfortunate int compared to 2 | 3 extra bytes
    'hanging around unused on our first stoTemp. I don't want to be bothered with a lot of huey of conversion from Byte to Int and Int to Byte
    'while loading, so there it is.
    Put #intSaveFileNo, 1, stoTemp  'I'm hard coding the 1 even though I don't think it's really necessary.
    'now put weightcarrying in:
    stoTemp.x = You.WeightCarrying
    'and write *it*
    Put #intSaveFileNo, , stoTemp   'OK, that's all for You. now for the threads.
    For i = 0 To NUMTHREADS Step 1  'right now this is a rather stupid 0 to 1. But we'll get bigger soon!
        stoTemp.x = Threads(i)
        Put #intSaveFileNo, , stoTemp
    Next i
    'now for the 'stuff' you're carrying.
    For i = 0 To CHAR_MAXPOSSESSIONS - 1 Step 1
        stoTemp.x = You.Possessions(i).x
        stoTemp.Y = You.Possessions(i).Y
        stoTemp.Desc = You.Possessions(i).Desc
        stoTemp.Movement = You.Possessions(i).Movement
        stoTemp.Tag = You.Possessions(i).Tag
        stoTemp.Type = You.Possessions(i).Type
        Put #intSaveFileNo, , stoTemp
    Next i
    'now for the stack of things to be restored
    For i = 0 To MAXSTOREDCOMM Step 1
        stoTemp.x = StoredScriptCommands(i).x
        stoTemp.Y = StoredScriptCommands(i).Y
        stoTemp.Desc = StoredScriptCommands(i).Desc
        stoTemp.Movement = StoredScriptCommands(i).Movement
        stoTemp.Tag = StoredScriptCommands(i).Tag
        stoTemp.Type = StoredScriptCommands(i).Type
        Put #intSaveFileNo, , stoTemp
    Next i
    
    SaveMap Fileno, ScreenX, ScreenY    'this may not be needed but I'll put it in anyway.
    SaveThings ObjFileno, ScreenX, ScreenY
    'and now we're done...that was actually the easy part! now I have to code the Load State code.
    Close intSaveFileNo
End Sub
