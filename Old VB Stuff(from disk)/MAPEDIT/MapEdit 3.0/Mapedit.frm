VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{7CDAE33A-0321-11D3-ADB9-646109C10000}#1.0#0"; "SmartScrollBar.ocx"
Object = "{7CDAE34A-0321-11D3-ADB9-646109C10000}#1.0#0"; "VSmartScrollBar.ocx"
Begin VB.Form frmMapEdit 
   Caption         =   "Map Editor for Chrysalis"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   9735
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
   Icon            =   "Mapedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   StartUpPosition =   2  'CenterScreen
   Begin VSmartScrollBar.VSScroll vsbJump 
      Height          =   3360
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5927
   End
   Begin SmartScrollBar.HSScroll hsbJump 
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   0
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   661
   End
   Begin VB.CheckBox chkWest 
      Caption         =   "<"
      DragIcon        =   "Mapedit.frx":0442
      Height          =   1455
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox chkSouth 
      Caption         =   "\/"
      DragIcon        =   "Mapedit.frx":074C
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox chkNorth 
      Caption         =   "/\"
      DragIcon        =   "Mapedit.frx":0A56
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   1455
   End
   Begin VB.CheckBox chkEast 
      Caption         =   ">"
      DragIcon        =   "Mapedit.frx":0D60
      Height          =   1455
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox picThumbnail 
      BorderStyle     =   0  'None
      Height          =   1350
      Left            =   3810
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   13
      Top             =   5160
      Width           =   1350
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1560
      Picture         =   "Mapedit.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Remove an Object"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   720
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   720
      Picture         =   "Mapedit.frx":116C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Alter an Object's Properties"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "&Move"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   105
      Picture         =   "Mapedit.frx":126E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Move an Object"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   600
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1230
      Picture         =   "Mapedit.frx":1370
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add Object"
      Top             =   5625
      UseMaskColor    =   -1  'True
      Width           =   600
   End
   Begin ComctlLib.ListView lvwTerrain 
      Height          =   6240
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   11007
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "imlTerrain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdEast 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3240
      Picture         =   "Mapedit.frx":1472
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5820
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton cmdSouth 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2820
      Picture         =   "Mapedit.frx":1574
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.CommandButton cmdNorth 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2820
      Picture         =   "Mapedit.frx":1676
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.CommandButton cmdWest 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      Picture         =   "Mapedit.frx":1778
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5820
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.PictureBox picViewport 
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   360
      MousePointer    =   4  'Icon
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   360
      Width           =   4800
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   480
      Picture         =   "Mapedit.frx":187A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5625
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   2880
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin ComctlLib.ImageList imlThings 
      Left            =   3360
      Top             =   4920
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
            Picture         =   "Mapedit.frx":197C
            Key             =   "Purina Table"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":25CE
            Key             =   "Potted Bush"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3220
            Key             =   "blank(do not use)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3E72
            Key             =   "Potted Palm"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4AC4
            Key             =   "Clay Pot"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5716
            Key             =   "Haystack"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6368
            Key             =   "Iron Pot"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6FBA
            Key             =   "Inscribed Pot"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":7C0C
            Key             =   "Ballot Box"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":885E
            Key             =   "Brick"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":94B0
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":A102
            Key             =   "Bottle"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":A954
            Key             =   "Bottles"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":B1A6
            Key             =   "Dictionary"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":BDF8
            Key             =   "Park Bench"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":C64A
            Key             =   "TGP"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":D29C
            Key             =   "Deluxe R Pizza"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":DEEE
            Key             =   "Karrot"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":EB40
            Key             =   "Ridiculous Pizza"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":F792
            Key             =   "Ridiculous Pizza Box"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":103E4
            Key             =   "Deluxe R Pizza Box"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":11036
            Key             =   "NoteinBottle"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":11C88
            Key             =   "BottleStack"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":128DA
            Key             =   "L5 Door"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1352C
            Key             =   "WoodenDoor"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1417E
            Key             =   "DeepHole"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":14DD0
            Key             =   "Monster Pit"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":15A22
            Key             =   "TellyBooth"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":16274
            Key             =   "BoatFloat"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":16EC6
            Key             =   "BoatIcon"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":17B18
            Key             =   "FF1Well"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1876A
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":193BC
            Key             =   "Mikey le Mouse"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1964E
            Key             =   "The Chatty Lady"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1A2A0
            Key             =   "Miney le Mouse"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1A532
            Key             =   "Professor"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1B184
            Key             =   "Mega Mouse"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1BDD6
            Key             =   "Fred the Freeloader"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1CA28
            Key             =   "Macky Le Mouse"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1CCBA
            Key             =   "Kat"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1D90C
            Key             =   "Easy Fix"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1E15E
            Key             =   "Shuffler"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1E9B0
            Key             =   "Grandpa Clone"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1F202
            Key             =   "GrandPa #13"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":1FA54
            Key             =   "Bottle Stacker"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":202A6
            Key             =   "Bunny"
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":20AF8
            Key             =   "Durty Kurt"
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2134A
            Key             =   "Ears the Rabbit"
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":21B9C
            Key             =   "Expendable Crewman"
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":223EE
            Key             =   "Eyes the Rabbit"
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":22C40
            Key             =   "Fat Rabbit"
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":23492
            Key             =   "Feet the Rabbit"
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":23CE4
            Key             =   "Ganwa"
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":24536
            Key             =   "Klown"
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":24D88
            Key             =   "Kangaroo Rat"
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":255DA
            Key             =   "Old Rabbit"
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":25E2C
            Key             =   "Sun Glassed Rat"
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2667E
            Key             =   "Pierre"
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":26ED0
            Key             =   "Rat"
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":27722
            Key             =   "Shipwrecked Guy"
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":27F74
            Key             =   "Solo"
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":287C6
            Key             =   "Spork"
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":29018
            Key             =   "Spotted Rabbit"
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2986A
            Key             =   "Stinky"
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2A0BC
            Key             =   "Blind Rat"
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2A90E
            Key             =   "Hermit #1"
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2B560
            Key             =   "KittyKat"
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2C1B2
            Key             =   "Stephen"
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2CA04
            Key             =   "Tough Rat"
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2D256
            Key             =   "W C Rabbit"
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2DAA8
            Key             =   "Yogi Rat"
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2E2FA
            Key             =   "Flame Warpher"
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2EB4C
            Key             =   "MegaMighty"
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":2F79E
            Key             =   "MegaSailor"
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":303F0
            Key             =   "Bun'rab'"
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":31042
            Key             =   "Kat2"
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":31C94
            Key             =   "ProfAllTiedUp"
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":328E6
            Key             =   "MineyAllTiedUp"
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":32B78
            Key             =   "Lady Bug"
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":337CA
            Key             =   "Bad Spider"
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3441C
            Key             =   "Live Flower"
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3506E
            Key             =   "Giant Roach"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "lblStatus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   18
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblPosition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblPosition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   5280
      Width           =   705
   End
   Begin ComctlLib.ImageList imlTerrain 
      Left            =   2280
      Top             =   4920
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
            Picture         =   "Mapedit.frx":35CC0
            Key             =   "Bush"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":35FDA
            Key             =   "Cave Floor"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":362F4
            Key             =   "Pool"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3660E
            Key             =   "Cave Wall"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":36928
            Key             =   "Fire"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":36C42
            Key             =   "Forest"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":36F5C
            Key             =   "Lawn"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":37276
            Key             =   "Gravel"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":37590
            Key             =   "House"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":378AA
            Key             =   "Mountain"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":37BC4
            Key             =   "Stalagmite"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":37EDE
            Key             =   "Fruit Tree"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":381F8
            Key             =   "Water"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":38512
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":39164
            Key             =   "Brick Wall"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":39996
            Key             =   "Carpet"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3A1E8
            Key             =   "Door"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3AA3A
            Key             =   "Windowed Door"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3B28C
            Key             =   "Cobbles"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3BEDE
            Key             =   "Krops"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3CB30
            Key             =   "Dead Krops"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3D782
            Key             =   "CFlower"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3E3D4
            Key             =   "MFlower"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3F026
            Key             =   "KFlower"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":3FC78
            Key             =   "SeaSandUp"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":408CA
            Key             =   "SeaSandLt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4151C
            Key             =   "SeaSandRt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4216E
            Key             =   "SeaSandDn"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":42DC0
            Key             =   "Sand"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":43A12
            Key             =   "Sea"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":44664
            Key             =   "Tile"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":452B6
            Key             =   "Dirt"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":45F08
            Key             =   "Dandelions"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":46B5A
            Key             =   "Grass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":477AC
            Key             =   "Tracks Left"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":483FE
            Key             =   "Tracks right"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":49050
            Key             =   "Tracks up"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":49CA2
            Key             =   "Tracks down"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4A8F4
            Key             =   "Stone Walk"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4B546
            Key             =   "Signature"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4C198
            Key             =   "Leafy Bush"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4CDEA
            Key             =   "Berry Bush"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4DA3C
            Key             =   "Boring Grass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4E68E
            Key             =   "Pomarbo"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4F2E0
            Key             =   "FF1TreeBot"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":4FF32
            Key             =   "FF1TreeTop"
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":50B84
            Key             =   "FirBot"
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":517D6
            Key             =   "FirMiddle"
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":52428
            Key             =   "FirTop"
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5307A
            Key             =   "FirLBot"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":53CCC
            Key             =   "FirLEdge"
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5491E
            Key             =   "FirLEdgeTop"
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":55570
            Key             =   "FirRBot"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":561C2
            Key             =   "FirREdge"
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":56E14
            Key             =   "FirREdgeTop"
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":57A66
            Key             =   "FirMidLBot"
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":586B8
            Key             =   "FirMidRBot"
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5930A
            Key             =   "FirMidLTop"
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":59F5C
            Key             =   "FirMidRTop"
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5ABAE
            Key             =   "ElevatorL"
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5B400
            Key             =   "ElevatorR"
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5BC52
            Key             =   "Goodgrass"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5C8A4
            Key             =   "Small Karrots"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5D4F6
            Key             =   "CSand BL"
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5E148
            Key             =   "CSand BR"
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5ED9A
            Key             =   "CSand TL"
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":5F9EC
            Key             =   "CSand TR"
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6063E
            Key             =   "Stonewall"
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":61290
            Key             =   "CSand InvBL"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":61EE2
            Key             =   "CSand InvBR"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":62B34
            Key             =   "CSand InvTL"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":63786
            Key             =   "CSand Inv TR"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":643D8
            Key             =   "Cement"
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6502A
            Key             =   "CementEngraved"
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":65C7C
            Key             =   "CementMsg"
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":668CE
            Key             =   "CementWriting"
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":67520
            Key             =   "CementSt"
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":68172
            Key             =   "CementLtSt"
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":68DC4
            Key             =   "CementTTT"
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":69A16
            Key             =   "CementCracked"
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6A668
            Key             =   "CementCracked2"
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6B2BA
            Key             =   "CementCracked3"
         EndProperty
         BeginProperty ListImage83 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6BF0C
            Key             =   "CCement"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage84 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6CB5E
            Key             =   "CCementEngr"
            Object.Tag             =   "C"
         EndProperty
         BeginProperty ListImage85 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mapedit.frx":6D7B0
            Key             =   "CCementSt"
            Object.Tag             =   "C"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAd 
      Caption         =   "Your Ad Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   6720
      Width           =   5175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCleanThingsArray 
         Caption         =   "&Clean Thing Array"
      End
      Begin VB.Menu mnuCleanThingsFile 
         Caption         =   "Clean &Thing File"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuShowSolid 
         Caption         =   "Sho&w Solid Tiles"
      End
      Begin VB.Menu mnuToolTips 
         Caption         =   "&Tool Tips"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRefreshThumbnail 
         Caption         =   "Refresh Thumb&nail"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCursorSize 
         Caption         =   "Cursor &Size..."
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset Cursor Size"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScrollFaster 
         Caption         =   "Scroll &Faster"
      End
      Begin VB.Menu mnuScrollSlower 
         Caption         =   "Scroll S&lower"
      End
      Begin VB.Menu mnuRequireClick 
         Caption         =   "Require &Click to Scroll"
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "&Context"
      Visible         =   0   'False
      Begin VB.Menu mnuWhatsThis 
         Caption         =   "&What's This?"
      End
      Begin VB.Menu mnuUnselect 
         Caption         =   "&Unselect All"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frmMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Beginning of version 2
'I have left almost all of the legacy code in here just in case my overhaul
'doesn't work out. (New(v3.0): Have removed most of it by now)

'By Nathan Sanders. I have changed a lot of things. It is the 15th of July(1998),
'and I am still not completely done. However, here is a list of the things that I've
'changed:
    '1.Changed the file system from Sequential with an array(MapYSize) of strings of length
'MapXSize to a Random file on disk of size MapYsize * MapXSize
'in addition a small portion of the map is repeatedly saved and loaded from disk in
'a 30 * 30 array to be edited and viewed. When the user moves 10 spaces(1 screen), the
'array is again saved to disk and loaded from a different portion.
    '2.Changed the startup sequence so that the user has a choice of opening an existing
'file and creating a new one.
    '3.Added support for mouse dragging to paint tiles as opposed to simply having to
'click repeatedly.
    '4.Added a highlight box(backed up by a modularized sub) that shows the exact tile
'at which the user is pointing. In addition, the box sizes to accomodate the size of
'the cursor(in tiles) according to what the user has chosen
    '5.Currently I am changing the file access code to only open and close the file
'at startup and shutdown of the program. This is accomplished by passing Load,SaveMap a Fileno
'parameter. I anticipate that this will increase performance somewhat.
'Note: Completed! But it doesn't have any noticable impact on my performance.
    '6.I changed the colors on the controls to system colors.
    '7.(07/21)Added just about all of the error checking. Note: the program still
'crashes when you move the mouse outside the picture box while still drawing if
'the cursor is outside the range of the 30*30 Map array(i.e. near the top, left,
'bottom, or right; especially when using a large cursor size).(08/15)Fixed this incidentally
    '8.(07/20)Have changed the picture boxes over to an ImageList and a linked Listview.
'PaintMap now uses the ImageList.ListImages(Index).Draw (hDC,x, y, style) method.
    '9.(07/21)All I have left is to move PaintMap to the explorer.bas module. This could cause some
'problems, but I think that I'm going to pass alot of arguments instead of making all of them
'global.(I know, I'm starting to sound like the programming books, but hey, why do they
'tell you that, anyway? Note:Completed with no problems.
    '10.Fixed a bug wherein if the map initialization was too slow, PaintMap would catch up
'and give a bug when Map(899) was not initialized from within LoadMap because it caught up
'before mnuNew was called.
'(after being called from inside Form_Paint when the form was shown after IT was called from
'within frmOpen_Ok(). So I simply just Loaded frmMapEdit from within frmOpen_Ok() insteead
'of Showing it. Then I Showed it AFTER calling LoadMap(so that that Map(899) was initialized)
    '11.Fixed a bug wherein, because I had typed a Y instead of an X in the if clause of the
'mouse move. But the bug wasn't in the Click, so I found out what the problem is that way.
'The behavior was that you couldn't drag-to-draw past a certain X value(usually 50).
    '12.Removed the Clear function as superfluous. Use Explorer, for crying out loud.
    '13. Changed the declaration of Map(900) to Map(899) because the last element
'wasn't being used.
    '14. (11/07)Fixed bug which prevented you from making maps bigger than 181^2(32767)
'because the temporary variable reserved for multiplying the two ints(MapX,YSize) was an int
'also...that caused Overflow problems, so I had to CLng both ints(MapX,YSize) before mult-
'iplying them.
    '15. (11/08)Fixed program design error by moving edge clip logic back into frmMapEdit
'into a sub called PaintPicViewPort. This calls PaintMap and does all edge clipping, as well
'as showing the user his position, drawing the Highlight box, and calling all future
'Paintxxx functions. Also simplified PaintMap's parameter list greatly and fixed a small bug
'wherein if you moved to the edge of the map, the program slowed down greatly because it was
'saving the map to disk repeatedly and loading it from the saved position. This was because
'I forgot to make the Boolean keeping track of whether or not we had saved Static. Now it
'runs much faster on the edge of the map.
    '16.(12/08) Added a Jump button which allows the user to instantaneously jump anywhere
'in the world(map). Included some hilarious Star Truck feedback for the benefit of the user.
'(Ha; ha)
    '17.(20/08) Have added many new features related to 'Things'. Added the Select, Add, Move
'Properties, and Delete buttons. Added the frmThingEdit which changes properties for 'Things'.
'(Still have not figured out how to set the icon to be selected in the ListView when
'I load frmThingEdit. Also integrated PaintThings into PaintPicViewPort. Now the only
'problem is that the objects FLICKER. I think this must be solved by adding another 320 ^ 2
'picture box, then Painting everything to it, then Blitting the contents of the extra onto
'PicViewPort.
    '18.Also changed the size of the form to c. 640 x 480, but also added sizing capabilities
'wherein you can size the form if you have a bigger resolution. The only thing that changes
'is the size of lvwTerrain, though, that being the main thing that the user would want more
'viewing space for.
    '19.Also added a very nice touch with the ToolTipText. The ToolTips now pop up and give
'you a terse description of the Tile you are looking at.(Not the 'Thing') The problem is that
'the ToolTips only disappear when you move the cursor to a new kind of tile, and then they
'leave a blank space until you move the cursor again(causing a PaintPicViewPort)
    '20. Added a context menu to picViewport which simplifies some of the command button
'clicking to manipulate objects. Actually, I have found that this almost makes the command
'buttons superfluous.
    '21.(08/29) Finally removed Mikey and Miney from the Tile imagelist. Also added a(buggy)
'command to turn off the tooltips. Can't figure out why it's not working.
    '22.Added a command that blits a Red 'X' on top of all solid tiles. Also added the
'convention that all solid tiles have an imagelist tag of "" and all clear tiles have a tag
'of 'C'.
    '23.Added the Key text to the items in the listview on frmThingEdit. A nice touch.
    '24.(08/31) Fixed ToolTip bug wherein the ToolTips would not turn off.
    '24(09/03) Fixed bad bug with the 'fixed' ToolTips that arose because I cut and pasted the code
'from PaintMap
    '25.(12/22) Fixed slight bug where the right edge check in PaintPicViewPort never saved its state as 'false' when it moved right;
'I never *noticed* this bug, but it's a good thing I found it nevertheless.
    '26 (12/22) Cool! Added a 90x90 picture box that displays a 3x3 pixel representation of every tile on the map...but since I had to come up with an algorithm
'that would use *exactly* 32767 colors, I simply increment through every color available on a 16-bit display. This produces a somewhat similar effect to night
'vision. Like I said, Cool! Also, it's *really* slow to repaint completely. However, I *have* stuck it inside the MouseUp for picViewport. This could be a really
'cool radar in the Game Engine...
'           *** 1999 ***
    '27 (04/23) Note: fixed bug about a month ago wherein the state of the 'Thing' was not loaded into the
'frmThingEdit correctly(unfortunately this is still not completely fixed; it just looks like it is)
    '28 Added mouseover checkboxes(I used them because they have the cool 3-D up/down toggle)
'that auto scroll and took out that stupid drag-on-the-edge code that RJ and Ryan recommended(it
'didn't work the way *anyone* wanted it to because we don't do full screen)
    '29 (05/<5) After the fact: wrote a custom ActiveX control that basically takes a scroll bar and doesn't throw a Change event when
'the programmer does 'hsbJump.Value = 999' like a normal scroll bar does. Then I took it and 1)synced the position of the scroll bar with the
'current TopX,Y (this gives a good overall 'where am I?' effect) and 2)let the user drag the thumb to instantly warp to that position. Then I
'deleted with joy the warp command button. One small thing; the code currently in there is fractious and insulting because you can't jump to
'exactly X|Y=0 or anywhere within 20 tiles of the far edge of the map...oh well I guess I should take out the insulting msgbox at least but
'I'm not going to right now.
    '30 (05/10) Changed the code for the mouseover checkboxes so that now they init a drag when they detect a mouseover, then when the
'chkNorth_Drag() detects a Leave, it cancels the drag, stops the scroll and unchecks the box. Now the leave code is perfect, except for that
'pesky outline box...there has to be an option to turn that off!(Later) Yes I found it, you change the DragIcon to a mousepointer looking thing.
'the best one I found looks like a mixture of the normal pointer and the 3d pointer that come with windows(it's all white, but still has that over
'hang over the stem of the arrow) Warning: Require click to scroll is buggy right now, but with the improved code, its use should be low...
'so I'm not going to fix it today! Too bad. Later(05/28): I think the reason the mouse pointer looks funny is that VB reduces pointers to 2
'colors or something like that.
    '31(05/27) Added a Clean Array menu. This 'cleans' all the deleted Things so they cannot be restored later by RestoreThingsArray.
'Also added Clean File but it doesn't have any code behind it nor will likely for some time.
Dim TerrainType As Integer
Dim ScreenX As Long 'global offset of Map array from 0,0 of world co-ordinates
Dim ScreenY As Long
Dim TopX As Integer 'offset of picture box from ScreenX,Y
Dim TopY As Integer
Dim CellY As Integer    'offset of cursor(or player in Game Engine) from ScreenX,Y
Dim CellX As Integer
Dim bBlocking As Integer    'whether or not we are drawing(i.e. the mouse is down)
Dim iScrolling As ScrollDir 'so that you can drag the mouse to scroll
Private Enum ScrollDir
    North = 1
    South = 2
    West = 3
    East = 4
    Stopped = 0
End Enum
'Const NORTH = 1 'scrolling constants
'Const SOUTH = 2
'Const WEST = 3
'Const EAST = 4
'Const STOPPED = 0
Dim wSelect As SelectState 'a flag that tells us what we're doing with the 'selection' of a 'thing'
Dim SelectX As Integer, SelectY As Integer 'whats the square that the user has selected to
'work with for an object?
Private Enum SelectState
    UNSELECTED = 0
    SELECTING = 1
    Selected = 2
    MOVING = 3
End Enum
'Const UNSELECTED = 0 'means that no'thing' is selected right now
'Const SELECTING = 1 'means that in PaintPicViewPort, we give the user SelHiLight, not HiLight.
'Const SELECTED = 2 'means that the user has selected an X,Y and filled SelectX,Y with them.
'Const MOVING = 3 'means that we're moving something around
Dim CursorXSize As Integer   'I had forgotten about this cool function
Dim CursorYSize As Integer
Dim Fileno As Integer
Dim ObjFileno As Integer    'this is the File of the 'Things' file: currently 65% of the
'size of the map file, but that will change if we change the structure of 'Thing'.
Dim bToolTips As Boolean    'this is to keep track of whether the user wants tooltips or not
Dim bShowSolid As Boolean 'this is to keep track of whether we should blit a 'blank' tile
'over all the solid tiles.
Dim bRequireClick As Boolean       'this is to see whether we need a *click* to start scrolling or
'just a mouseover.
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
'this function declare and accompanying constant are unused currently, but may be used in the
'real game.
Const SRCCOPY = &HCC0020

Private Sub PaintPicViewport()
    'this will hold both PaintMap(explore.bas) and PaintStuff (packrat.bas) eventually.
'plus all of the edge test code that is currently inside PaintMap. I have needed to clean
'PaintMap up ever since I made it Public, and now is the time.

    'here we have the edge map test code that used to be in PaintMap(explore.bas)
Static bSaved As Boolean    'alert: just fixed the problem wherein the map got VERY slow at
'land's end. I forgot to make bSaved static and it came up as False every time.(Boy do I
'feel stupid.)

'First move the array and clip it to the edges.
    If TopX = 20 Then
        If ScreenX + 30 = MapXSize And bSaved = False Then  'we're at map edge!
            SaveMap Fileno, ScreenX, ScreenY    'save the array to disk but DO NOT move the array
            SaveThings ObjFileno, ScreenX, ScreenY
            'over to next position because it would otherwise go off the edge.
            '(or reset the viewport)
            bSaved = True 'turn on a switch to make sure we don't repeatedly save to disk
            'when moving along the edge of the map(because we don't reset position when
            'moving along edge of map)
        ElseIf ScreenX + 30 = MapXSize And bSaved = True Then
            'do nothing at all(because continually saving degradates! performance.
        Else    'move along now
            TopX = 10   'reset the viewport to center of array
            SaveMap Fileno, ScreenX, ScreenY    'save changes of current position to disk
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenX = ScreenX + 10  'move array over 10 cells to next pos.
            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
            LoadThings ObjFileno, ScreenX, ScreenY
            PaintThumbNail  'this new: paints little thumb at bottom of screen
            bSaved = False
        End If
    End If
    
    If TopY = 20 Then
        If ScreenY + 30 = MapYSize And bSaved = False Then  'same comments here...
            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            bSaved = True
        ElseIf ScreenY + 30 = MapYSize And bSaved = True Then  'skip
        Else    'we're not at the edge of the screen, so business as usual
            TopY = 10
            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenY = ScreenY + 10
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
            PaintThumbNail  'this new: paints little thumb at bottom of screen
            bSaved = False  'fixed bug!! This was missing!
        End If
    End If
    'oops, forgot to add top, left checking(I was really tired last night)
    If TopX = 0 Then
        If ScreenX = 0 And bSaved = False Then 'we're at maps edge(world's end)
            SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
            SaveThings ObjFileno, ScreenX, ScreenY
            bSaved = True
        ElseIf ScreenX = 0 And bSaved = True Then
        Else
            TopX = 10
            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenX = ScreenX - 10
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
            PaintThumbNail  'this new: paints little thumb at bottom of screen
            bSaved = False
        End If
    End If
    If TopY = 0 Then
        If ScreenY = 0 And bSaved = False Then 'we're at maps edge(world's end)
            SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
            SaveThings ObjFileno, ScreenX, ScreenY
            bSaved = True
        ElseIf ScreenY = 0 And bSaved = True Then
        Else
            TopY = 10
            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            ScreenY = ScreenY - 10
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
            PaintThumbNail  'this new: paints little thumb at bottom of screen
            bSaved = False
        End If
    End If

    PaintMap picViewport, imlTerrain, TopX, TopY 'this is the NEW, IMPROVED reduced param list.
    'probably have to leave DrawHighLight inside MouseMove because I don't what the state is
    'inside this function...or do I? Yeah, I do. It's global so I can share it between events
    'inside picViewport
    'now paint the objects
    PaintThings picViewport, imlThings, TopX, TopY, ScreenX, ScreenY 'paint the 'Things' onto picBackBuffer as
    'well.
    If wSelect = SELECTING Then 'if they're selecting, display a depressed hilite! that is
    'always one cell ^ 2.
        DrawSelHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
        32, 32
    ElseIf bBlocking Then   'otherwise, if we're drag-to-drawing, show a depressed hilite!
    'that is the size of the 'cursor'
        DrawSelHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
        CursorXSize * 32, CursorYSize * 32  'this should work. I just stole it from
    Else 'picBackBuffer_MouseMove() verbatim.
    'or just show a normal hilite! that is the size of the 'cursor'.
        DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
        CursorXSize * 32, CursorYSize * 32
    End If
    
    If wSelect = Selected Then  'make sure we keep the cell they selected highlighted as a
    'visual cue.
        DrawSelHighLight picViewport, (SelectX - TopX - ScreenX) * 32, (SelectY - TopY - ScreenY) * 32, _
        32, 32
    End If
    If bShowSolid Then  'if the user requested to see all solid tiles blitted with an extra
    'x' over them. So that's what we'll do(stealing the loop code from PaintMap in explore.bas)
    '***later note:this *should* be in a sub but it isn't...
    Dim XIndex As Integer, YIndex As Integer
    Dim TerrainVal As Integer
    Dim TempX As Integer, TempY As Integer
    
        'for all the tiles on the screen...
        For YIndex = 0 To 9 'change to a constant called YVIEWPORT
            For XIndex = 0 To 9 'change to a constant called XVIEWPORT
                TempX = TopX + XIndex 'figure our current position
                TempY = TopY + YIndex
    '               replace 30 with ARRAY_X_SIZE someday
                TerrainVal = Map((TempY * 30) + TempX) 'Look up the value in the array
                '"" means that it is solid, "C" means that it is 'Clear'
                If TerrainVal > -1 Then 'this must be two separate ifs, because(unlike C)
                'VB checks both conditions before chunking a comparison. Then if TerrainVal
                'is -1, it also checks imlTerrain which(being a image list) stops at 1, has
                'no -1. This generates an error.
                'This also makes me think that Ryan is right that VB is sometimes balky
                'for game programming.
                    If imlTerrain.ListImages(TerrainVal).Tag = "" Then 'blit an 'X'
                    imlThings.ListImages("blank(do not use)").Draw picViewport.hDC, XIndex * 32, YIndex * 32, imlTransparent
                    End If
                End If 'other style possibilities include imlTransparent,imlSelected, and imlFocus
            Next XIndex
        Next YIndex
        
    End If
    'now bitblt everything to picViewport
    'BitBlt picViewport.hDC, 0, 0, 320, 320, picBackBuffer.hDC, 0, 0, SRCCOPY
    'this taken out because it didn't work that well...and for some reason the cursor didn't work
    'anymore.
    
    'tell the user his position 'Note that commented text is NO LONGER NEEDED. However, if you wish to re-enable it for debugging
    'purposes, feel free to do so; the info is somewhat useful for determining the 'screen' edges...
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY '& " TopX = " & TopX & " TopY = " & TopY
    'update the scroll bars
    hsbJump.Value = TopX + ScreenX
    vsbJump.Value = TopY + ScreenY
End Sub
Private Sub chkEast_Click()
    If bRequireClick Then 'start the scroll, Solo!
'yes sir!; setting scroll direction.
        iScrolling = East
'enabling cancel detection!
        chkEast.Drag vbBeginDrag
'executing!!
        tmrScroll.Enabled = True
    End If
'<Star Truck theme here>
End Sub

Private Sub chkEast_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If State = 1 Then '1 means Leave, sir.
'thanks, Spork. Solo, end scroll!
'yes, sir!; preparing visual feedback.
        chkEast.Value = vbUnchecked
'ending cancel detection.
        chkEast.Drag vbCancel
'flagging to stop scrolling
        iScrolling = Stopped
'stop scroll timer!
        tmrScroll.Enabled = False
    End If
'<End song here>
End Sub

Private Sub chkEast_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not bRequireClick And iScrolling = Stopped Then   'start the scroll, Solo!
'yes, sir!; preparing visual feedback.
        chkEast.Value = vbChecked
'setting scroll direction.
        iScrolling = East
'enabling cancel detection!
        chkEast.Drag vbBeginDrag
'executing!!
        tmrScroll.Enabled = True
    End If
'<Star Truck theme here>
End Sub
Private Sub chkNorth_Click()
    If bRequireClick Then
        iScrolling = North
        chkNorth.Drag vbBeginDrag
        tmrScroll.Enabled = True
    End If
End Sub

Private Sub chkNorth_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If State = 1 Then 'leave
        chkNorth.Value = vbUnchecked
        chkNorth.Drag vbCancel
        iScrolling = Stopped
        tmrScroll.Enabled = False
    End If
End Sub

Private Sub chkNorth_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not bRequireClick And iScrolling = Stopped Then
        chkNorth.Value = vbChecked
        iScrolling = North
        chkNorth.Drag vbBeginDrag
        tmrScroll.Enabled = True
    End If
End Sub
Private Sub chkSouth_Click()
'Static bInside As Boolean   'maybe this will help?--hmm but it won't help the problem in the DragOver/Leaving scenario...maybe we'll
'just have to leave Require Click to Scroll buggy for now.
    If bRequireClick = True Then
        'And bInside = False 'this used to be in there in conjuction with the static variable in an attempt to restrict
        'bInside = True 'the number of _Click() events recived and/or acted on.
        'chkSouth.Value = vbChecked  'this causes another _Click() event. why doesn't this cascade out of control? [Stack overflow]
        iScrolling = South
        chkSouth.Drag vbBeginDrag
        tmrScroll.Enabled = True
        'bInside = False
    End If
End Sub
Private Sub chkSouth_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If State = 1 Then 'leave
        chkSouth.Value = vbUnchecked    'this causes a _Click() event
        chkSouth.Drag vbCancel              'it's a cascading type thing where you end up with about 30 items in the
        iScrolling = Stopped                'call stack...
        tmrScroll.Enabled = False
    End If
End Sub

Private Sub chkSouth_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not bRequireClick And iScrolling = Stopped Then
        chkSouth.Value = vbChecked
        iScrolling = South
        chkSouth.Drag vbBeginDrag
        tmrScroll.Enabled = True
    End If
End Sub
Private Sub chkWest_Click()
    If bRequireClick Then
        iScrolling = West
        chkWest.Drag vbBeginDrag
        tmrScroll.Enabled = True
    End If
End Sub
Private Sub chkWest_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If State = 1 Then 'leave
        chkWest.Value = vbUnchecked
        chkWest.Drag vbCancel
        iScrolling = Stopped
        tmrScroll.Enabled = False
    End If
End Sub
Private Sub chkWest_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not bRequireClick And iScrolling = Stopped Then
        chkWest.Value = vbChecked
        iScrolling = West
        chkWest.Drag vbBeginDrag
        tmrScroll.Enabled = True
    End If
End Sub
Private Sub cmdAdd_Click()
Dim Temp As Thing
    If wSelect = Selected Then  'if they've clicked cmdSelect, all is well...
        Temp.x = SelectX
        Temp.Y = SelectY
    Else    'they need to click cmdSelect to choose a cell.
        MsgBox "You need to select a cell to which to add the 'Thing'!"
        wSelect = UNSELECTED    'reset the select vars
        SelectX = -1
        SelectY = -1
        cmdAdd.Enabled = False  'disable all the 'Thing' buttons now that wSelect = UNSELECTED
        cmdMove.Enabled = False
        cmdProperties.Enabled = False
        cmdRemove.Enabled = False
        Exit Sub
    End If
    Load frmThingEdit   'this so we can alter its controls if need be.
    'now change frmThingEdits control properties
    With frmThingEdit
        .lvwType.ListItems.Item(1).Selected = True
        .lstMoveStyle.ListIndex = 0
        .txtDesc.Text = "1"
        .chkImmovable.Value = vbUnchecked
        .txtWeight.Text = "1"
        .txtWeight.Enabled = True
        .Show vbModal 'get the info off of frmThingEdit
        If .Tag = "OK" Then
            Temp.Desc = CInt(.txtDesc.Text)
            Temp.Movement = .lstMoveStyle.ListIndex
            Temp.Type = .lvwType.SelectedItem.Index
            Mid(Temp.Tag, 1, 1) = Chr$(.txtWeight.Text)
        Else    'returned Cancel from frmthingedit
            Exit Sub
        End If
    End With
    If Not (PutThing(Temp, ScreenX, ScreenY)) Then    'call the function from packrat.bas
    'if PutThings returns false, it means that the screen is full
        MsgBox "Screen is full! You must remove some 'Things' from this screen to be able to add more. " _
        & vbCrLf & "(Note: Try adding objects on the boundaries of two screens so that you can get 20 'Things' in an equivalent space)"
    End If
    wSelect = UNSELECTED    'reset the select vars
    SelectX = -1
    SelectY = -1
    cmdAdd.Enabled = False  'disable all the 'Thing' buttons now that wSelect = UNSELECTED
    cmdMove.Enabled = False
    cmdProperties.Enabled = False
    cmdRemove.Enabled = False
End Sub

Private Sub cmdEast_Click()
'Dim Dummy As Integer ' As Long
    TopX = TopX + 1
    CellX = CellX + 1
    If TopX = 21 Then   'note that PaintPicViewPort uses the value 20 to tell when we're at
        TopX = 20   'the edge of the map(it just leaves it at 20 and tests for it next time
        CellX = CellX - 1   'around too.)
    End If
    PaintPicViewport    'new sub that calls all Paintx functions that are in the *.bas files
End Sub


Private Sub cmdJump_Click()
Dim JumpY, JumpX
    JumpX = InputBox("Caption: Enter the X coordinate of the Screen to which you'd like to jump. Note: It must be bigger than 9 and smaller than MapXSize - 20.", "Jump Co-Ordinates", ScreenX + TopX)
    JumpY = InputBox("Caption: Enter the Y coordinate of the Screen to which you'd like to jump. Note: It must be bigger than 9 and smaller than MapYSize - 20.", "Jump Co-Ordinates", ScreenY + TopY)
    
    If JumpX = "" Or JumpY = "" Then Exit Sub   'they canceled!
    If Not (IsNumeric(JumpX)) Or Not (IsNumeric(JumpY)) Then Exit Sub   'they entered text!!
    MsgBox "Computing Jump Co-ordinates"    'give the user some Star Truck Feed-Back
    JumpX = (JumpX \ 10) * 10   'this to make sure that the co-ordinates are evenly
    JumpY = (JumpY \ 10) * 10   'divisible by 10
    If JumpX > 9 And JumpX < (MapXSize - 20) And JumpY > 9 And JumpY < (MapYSize - 20) Then
        MsgBox "Transmitting jump data at 16.6 M/Sec."  'more Warped Star Truck feedback
        SaveMap Fileno, ScreenX, ScreenY
        ScreenX = JumpX - 10
        ScreenY = JumpY - 10
        TopX = 10
        TopY = 10
        LoadMap Fileno, ScreenX, ScreenY
        LoadThings ObjFileno, ScreenX, ScreenY
        PaintPicViewport
        MsgBox "We have arrived safely at (" & JumpX & ", " & JumpY & "), Caption!"
    Else
        MsgBox "Error, Caption! Invalid Jump Co-ordinates!"
    End If
End Sub

Private Sub cmdMove_Click()
Dim Temp As Thing
Static Count As Integer 'this is the address in the Things array of the 'thing' to be moved.
'it is static because cmdmove is called TWICE: once from the Click, when wSelect should be
'selected, and once from picViewport_Click() where wSelect should be MOVING and SelectX,Y
'should have in them the NEW position of Thing(Count) (saved from first time around) which
'we then proceed to set using MoveThing. The complication comes from the fact that they
'MIGHT be trying to move the 'Thing' off-screen(but not off of picViewport). Then we have to
'RemoveThing and PutThing at SelectX,Y. Unfortunately, the NEW screen just might already be
'full -- if it is, we give the whole thing up in disgust(after telling the user).
    If wSelect = Selected Then  'if they've clicked cmdSelect, all is well...
        If IsThing(SelectX, SelectY, Count) = False Then
            MsgBox "You need to select a 'Thing' to move!"
            wSelect = UNSELECTED    'reset the select vars
            SelectX = -1
            SelectY = -1
            cmdAdd.Enabled = False  'disable all the 'Thing' buttons now that wSelect = UNSELECTED
            cmdMove.Enabled = False
            cmdProperties.Enabled = False
            cmdRemove.Enabled = False
            Exit Sub
        End If
        wSelect = MOVING    'tell picViewport that I need a value to move the 'Thing' to.
        cmdAdd.Enabled = False  'disable all the 'Thing' buttons so that the user can't do
        cmdMove.Enabled = False 'something unexpected.
        cmdProperties.Enabled = False
        cmdRemove.Enabled = False
        Exit Sub    'get out of this sub till next time when we actually place the 'Thing'
    ElseIf wSelect = MOVING Then    'this code could just as easily been placed inside
    'picViewport_Click... but I had no way to give it Count from Sub to Sub. So here we are,
    'running it this way.
        If MoveThing(Count, SelectX, SelectY) = False Then  'we need to delete the 'Thing' and
        'add it to the right screen because MoveThing couldn't move it off-screen.
            Temp = Things(Count) 'save Things(Count)
            Temp.x = SelectX
            Temp.Y = SelectY
            RemoveThingArray Count 'delete it.
            If PutThing(Temp, ScreenX, ScreenY) = False Then  'aaaargghhh!!! the other screen is full!!
                MsgBox "You are trying to move a 'Thing' to another screen. " _
                & vbCrLf & "That screen is full. Tough Luck."
            Else 'success!!
                Count = -1  'reset count variable
            End If
        End If
    Else    'they need to click cmdSelect to choose a cell.
        MsgBox "You need to select a 'Thing' to move!"
        Exit Sub
    End If
    wSelect = UNSELECTED    'reset the select vars
    SelectX = -1
    SelectY = -1
    cmdAdd.Enabled = False  'disable all the 'Thing' buttons now that wSelect = UNSELECTED
    cmdMove.Enabled = False
    cmdProperties.Enabled = False
    cmdRemove.Enabled = False
End Sub

Private Sub cmdProperties_Click()
Dim Count As Integer
    If wSelect = Selected Then
        If Not IsThing(SelectX, SelectY, Count) Then
            MsgBox "No 'thing' selected to alter properties thereof!"
        Else
            Load frmThingEdit
            With frmThingEdit
                .lvwType.ListItems(Things(Count).Type).Selected = True
                .lvwType.ListItems(Things(Count).Type).EnsureVisible    'help this listview doesn't have a property that actually
                .lstMoveStyle.ListIndex = Things(Count).Movement    'sets the focus of the listview to a particular value!!!
                .txtDesc.Text = Things(Count).Desc
                If Things(Count).Desc < PERSON Then
                    .txtWeight.Text = Asc(Things(Count).Tag)   ' I don't have to trim the thing with Left because Asc only looks at  first letter anyway
                    .txtWeight.Enabled = True
                    .chkImmovable.Enabled = True
                    If Asc(Mid$(Things(Count).Tag, 1, 1)) > 0 Then 'it's movable
                        .chkImmovable.Value = vbUnchecked
                        .txtWeight.Enabled = True
                    Else
                        .chkImmovable.Value = vbChecked 'let's see if we can get away with not calling _Click
                        .txtWeight.Enabled = False
                    End If
                Else
                    .chkImmovable.Enabled = False
                    .txtWeight.Enabled = False
                End If
                frmThingEdit.Show vbModal
                If frmThingEdit.Tag = "OK" Then
                    Things(Count).Desc = CInt(frmThingEdit.txtDesc.Text)
                    Things(Count).Movement = frmThingEdit.lstMoveStyle.ListIndex
                    Things(Count).Type = frmThingEdit.lvwType.SelectedItem.Index
                    Mid$(Things(Count).Tag, 1, 1) = Chr$(frmThingEdit.txtWeight.Text)
                End If
            End With
        End If
    End If
    wSelect = UNSELECTED    'reset the select vars
    SelectX = -1
    SelectY = -1
    cmdAdd.Enabled = False  'disable all the 'Thing' buttons now that wSelect = UNSELECTED
    cmdMove.Enabled = False
    cmdProperties.Enabled = False
    cmdRemove.Enabled = False
End Sub

Private Sub cmdRemove_Click()
    If wSelect = Selected Then
        If Not RemoveThing(SelectX, SelectY) Then
            MsgBox "Nothing selected to remove!"
        End If
    End If
    wSelect = UNSELECTED    'reset the select vars
    SelectX = -1
    SelectY = -1
    cmdAdd.Enabled = False  'disable all the 'Thing' buttons now that wSelect = UNSELECTED
    cmdMove.Enabled = False
    cmdProperties.Enabled = False
    cmdRemove.Enabled = False

End Sub

Private Sub cmdSelect_Click()
    wSelect = SELECTING
    SelectX = -1
    SelectY = -1
End Sub

Private Sub cmdNorth_Click()
    TopY = TopY - 1
    CellY = CellY - 1
    If TopY = -1 Then
        TopY = 0
        CellY = CellY + 1
    End If
    PaintPicViewport    'new sub that calls all Paintx functions that are in the *.bas files
End Sub

Private Sub cmdSouth_Click()
    TopY = TopY + 1
    CellY = CellY + 1
    If TopY = 21 Then
        TopY = 20
        CellY = CellY - 1
    End If
    PaintPicViewport    'new sub that calls all Paintx functions that are in the *.bas files
End Sub

Private Sub cmdWest_Click()
    TopX = TopX - 1
    CellX = CellX - 1
    If TopX = -1 Then
        TopX = 0
        CellX = CellX + 1
    End If
    PaintPicViewport    'new sub that calls all Paintx functions that are in the *.bas files
End Sub

Private Sub Form_Paint()
    PaintPicViewport    'new sub that calls all Paintx functions that are in the *.bas files
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        cmdWest_Click
    ElseIf KeyCode = vbKeyDown Then
        cmdSouth_Click
    ElseIf KeyCode = vbKeyUp Then
        cmdNorth_Click
    ElseIf KeyCode = vbKeyRight Then
        cmdEast_Click
    End If
End Sub
Private Sub Form_Load()
'Dim TempX, TempY
Dim Count As Integer
    frmMapEdit.ScaleMode = vbPixels
'    picViewport.ScaleMode = vbPixels
    'picCanvas.ScaleMode = 3
'    For Count = 0 To 4 Step 1
'        picTerrain(Count).ScaleMode = vbPixels
'    Next Count
    CursorXSize = 1 'initialize the variables(more later)
    CursorYSize = 1
    ScreenX = 0
    ScreenY = 0
    TopX = 10
    TopY = 10
    TerrainType = 1
    SelectX = -1
    SelectY = -1
    bToolTips = True    'give them tooltips, and let them eat cake(unsupported until we get
    With lvwTerrain.ListItems                                   'one drawn)
    For Count = 1 To imlTerrain.ListImages.Count
        .Add Count, , imlTerrain.ListImages(Count).Key, imlTerrain.ListImages(Count).Index ', frmMapEdit.imlTerrainSm.ListImages(Count).Index
    Next Count  'boy it took me a long time to figure out how to load the image in from the
    'ImageList(unfortunately, Bob will probably teach it next class)
    End With
    
    frmMapEdit.Height = Screen.Height - 1440 'make the form almost fill the screen.
    frmMapEdit.Width = Screen.Width - 1440
    'various functions...usually self-explanatory, so I'll spare you the explanation
    AdLabel
'    PaintThumbNail
End Sub
Private Sub AdLabel()
        Randomize Timer 'generate a hilarious comment for the Ad Label
        'think up new comments and send them the Nathan!!
Dim Comment As Integer
    Comment = CInt(Rnd * 8) + 1 'nine 'ads'
    With lblAd
    .FontSize = 36
    Select Case Comment
        Case 1
            .Caption = "Your Ad Here"
        Case 2
            .Caption = "Buy Shoes -- " & vbCrLf & "Nike™ Shoes"
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
            .Caption = "Eat mor chikin!®"
        Case 9
            .FontSize = 24
            .Caption = "Emus taste good, " & vbCrLf & "like poultry should."
        Case Else
            MsgBox "Select Case Error in AdLabel!"
            .Caption = "Your Ad Here"
    End Select
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'NOTE: In a desparate effort to get Mapedit to close completely when you click the X button, I am commenting out all code except for
'Close Fileno and Close ObjFileno!!

    Close Fileno
    Close ObjFileno
    'this is the old code--but it doesn't close completely for some reason.
'If Fileno <> 0 And ObjFileno <> 0 And frmOpen.Tag = "" Then     'if we've got an open file and aren't
''still trying to get the user to open or create a file from frmOpen
'    Dim Answer As Integer
'    Answer = MsgBox("Save Changes?", vbYesNoCancel + vbQuestion, "Save Map")
'    If Answer = vbYes Then
'        SaveMap Fileno, ScreenX, ScreenY
'        SaveThings ObjFileno, ScreenX, ScreenY
'        Close Fileno
'        Close ObjFileno
'        End
'    ElseIf Answer = vbNo Then
'        Close Fileno
'        Close ObjFileno
'        End
'    ElseIf Answer = vbCancel Then
'        Cancel = True
'    Else    'this is unneeded but I put it in anyway.
'        MsgBox "Really big MsgBox Error!"
'        Cancel = True
'        Exit Sub
'    End If
'End If

End Sub

Private Sub Form_Resize()
    'here I'm going to resize the listview to cater to people with high-res screens...
    With lvwTerrain
    Dim RIGHT As Integer, Bottom As Integer
        RIGHT = .Left + .Width
        Bottom = .Top + .Height
        RIGHT = (frmMapEdit.ScaleWidth - RIGHT) + .Width
        If RIGHT > 17 Then .Width = RIGHT
        Bottom = (frmMapEdit.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
'        frmMapEdit.Width = 100
'        Right = 98
'        .Width = 28
        '(100 - 98) + 28
        '2 + 28
        '30
    End With
End Sub

Private Sub hsbJump_Change()
'here put the jump code from cmdJump

Dim lngJumpX As Long
    If hsbJump.Value > ScreenX And hsbJump.Value < ScreenX + 20 Then  'we just need to move the screen a little
        TopX = hsbJump.Value - ScreenX
        PaintPicViewport
    Else    'need a real jump.
        lngJumpX = (hsbJump.Value \ 10) * 10   'get the screen coordinates
        
        If lngJumpX > 9 And lngJumpX < (MapXSize - 20) Then
            'MsgBox "Transmitting jump data at 16.6 M/Sec."  'more Warped Star Truck feedback
            SaveMap Fileno, ScreenX, ScreenY
            ScreenX = lngJumpX - 10
            TopX = hsbJump.Value - ScreenX
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
            PaintPicViewport
            'MsgBox "We have arrived safely at (" & JumpX & ", " & JumpY & "), Caption!"
        Else
            MsgBox "Error, Caption! Invalid Jump Co-ordinates!"
        End If
    End If
    'lblStatus.Caption = "X: " & hsbJump.Value  'this taken care of in _Scroll.
End Sub
Private Sub hsbJump_Scroll()
    lblStatus.Caption = "X: " & hsbJump.Value
End Sub

Private Sub lblAd_Click()
    AdLabel
End Sub

Private Sub lvwTerrain_ItemClick(ByVal Item As ComctlLib.ListItem)
    TerrainType = Item.Index
End Sub

Private Sub mnuAdd_Click()
    wSelect = Selected  'tell cmdAdd that we've selected something
    SelectX = CellX 'and set the position to the current cursor position
    SelectY = CellY
    cmdAdd_Click
End Sub

Private Sub mnuCleanThingsArray_Click()
'this function 'cleans' the thing array--it takes all Things whose X is -1(the invisible flag) and Y is something else and
'resets the values of the whole Thing. That way the Restore no longer works and we've got rid of some nasty screen-too-full problems.
'Don't use this too much because it spoils people's fun :).
    CleanThingArray 'just call the PackRat function.
End Sub

Private Sub mnuCleanThingsFile_Click()
    MsgBox "Because it would hang your computer almost as long as File|New does!", vbOKOnly, "This feature not implemented in this version."
End Sub

Private Sub mnuCursorSize_Click()
Dim Temp As String
Dim OldXSize As Integer
    Temp = InputBox("Enter cursor X size: (Do not go over ten)", "Cursor Size", "1")
    If Temp = "" Then   'they canceled...um come on; we don't need to set the cursor to a default; we need to exit sub with no changes!!
        'therefore I'm changing this code. from setting CursorXSize to 1 to Exit Sub
        Exit Sub    'don't change the cursor size at all!!
    Else
        OldXSize = CursorXSize  'this so if they cancel in the Y Inputbox we *completely* reset the cursor size settings.
        CursorXSize = CInt(Temp)
    End If
    Temp = InputBox("Enter cursor Y size: (Do not go over ten)", "Cursor Size", "1")
    If Temp = "" Then
        CursorXSize = OldXSize
        Exit Sub    'same here.
    Else
        CursorYSize = CInt(Temp)
    End If
    'now make sure that if somebody(i.e. RJ) tries to crash the program they can't over/undersize the cursor.
    If CursorXSize < 1 Then CursorXSize = 1
    If CursorXSize > 10 Then CursorXSize = 10
    If CursorYSize < 1 Then CursorYSize = 1
    If CursorYSize > 10 Then CursorYSize = 10
End Sub

Private Sub mnuExit_Click()
Dim Answer As Integer
    If Fileno <> 0 And ObjFileno <> 0 Then
        Answer = MsgBox("Save Changes?", vbYesNoCancel + vbQuestion, "Save Map")
        If Answer = vbYes Then
            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
            Close Fileno
            Close ObjFileno
        ElseIf Answer = vbNo Then
            Close Fileno
            Close ObjFileno
        ElseIf Answer = vbCancel Then
            Exit Sub    'oops, they changed their mind.
        End If
    End If
    End
End Sub

'Private Sub mnuIconSize_Click()
'    If mnuIconSize.Checked = False Then
'        lvwTerrain.View = lvwSmallIcon
'        mnuIconSize.Checked = True
'    Else
'        lvwTerrain.View = lvwIcon
'        mnuIconSize.Checked = False
'    End If
'End Sub

Private Sub mnuMove_Click()
    wSelect = Selected
    SelectX = CellX
    SelectY = CellY
    'call cmdMove the first time to tell it that we're ready to move
    cmdMove_Click
End Sub

Public Sub mnuNew_Click()
Dim Count As Long
Dim Dummy As Integer
Dim TempThing As Thing
Dim FirstTime As Boolean
'Dim Opener As String 'this is now global(I think)
    On Error GoTo ErrHandler
    If frmOpen.Tag <> "" Then
        frmOpen.Tag = ""
        FirstTime = True   'the user didn't actually click the New menu(but the Open form)
    End If
    
    CMDialog1.Flags = cdlOFNOverwritePrompt
    CMDialog1.InitDir = App.Path
    CMDialog1.Filter = "Map Files (*.map)|*.map|All Files (*.*)|*.*"
    CMDialog1.filename = "Untitled.map"
    CMDialog1.ShowSave
    
    Opener = CMDialog1.filename 'OK we've got a good file name, and the user doesn't want to
    'cancel
    
    If Not FirstTime Then   'check if this is the user requesting a new map
        Dim TempX, TempY
        Do  'make SURE that the user can't enter a bad number or cancel
            TempX = InputBox("Enter X size for map(not less than 30)", , "1000")
        Loop Until TempX <> ""
        Do
            TempY = InputBox("Enter Y size for map(not less than 30)", , "1000")
        Loop Until TempY <> ""
        MapXSize = TempX
        MapYSize = TempY
        If MapXSize < 30 Then MapXSize = 30
        If MapYSize < 30 Then MapYSize = 30
        Close Fileno    'make sure we close the current file(maybe we should save the map too
        'but it is saved a lot anyway...
        Close ObjFileno
    End If
    
    Fileno = FreeFile   'continue opening the file
    Open Opener For Random As #Fileno Len = Len(Dummy) 'dummy is an integer since I can't
                                                        'remember what VB's sizeof looks like
    ObjFileno = FreeFile
    ObjOpener = (Left(Opener, Len(Opener) - 3)) & "thi"
    Open ObjOpener For Random As #ObjFileno Len = Len(TempThing)
    'now fill the 'Things' file with empty 'things'. Note -1 means unitialized to the program(later: I made -1 a constant named NONE)
    TempThing.Desc = NONE
    TempThing.Movement = STILL
    TempThing.Type = NONE
    TempThing.x = NONE
    TempThing.Y = NONE
    For Count = 1 To (((CLng(MapXSize) \ 10) * (CLng(MapYSize) \ 10)) * 10) + 1 'note int division
        Put #ObjFileno, Count, TempThing
    Next Count
        'Put map size at beginning  (sadly, I have figured out how to do it the way I wanted
        'to originally, but it is working this way, so who cares?)
    Put #Fileno, , MapXSize
    Put #Fileno, , MapYSize
'fill the map with -1(nothing)
'bug fix!!:you must convert MapX,Ysize to Longs before multiplying them if the combined
'total is bigger than 32767. That's because VB reserves a temporary variable that you don't
'know about when it multiplies stuff. If both variables are ints, it(stupidly) assumes that
'the answer will not be out of range and makes the temporary var an int also.
    For Count = 3 To (CLng(MapYSize) * CLng(MapXSize)) + 3 Step 1 'start at third position(offset from MapX,YSize)
            Put #Fileno, Count, -1
    Next Count
    'init variables
    ScreenX = 0
    ScreenY = 0
    TopX = 10
    TopY = 10

    'we're done
    LoadMap Fileno, ScreenX, ScreenY
    LoadThings ObjFileno, ScreenX, ScreenY  'new!!
    If FirstTime = True Then
        frmMapEdit.Show
    End If
    
    frmMapEdit.Caption = Opener + " - Map Editor"
    PaintPicViewport
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    If Err.Number = cdlCancel Then
        If FirstTime = True Then    'a tag set when the user pressed a command button back
            'on frmOpen (this for correct timing)
            frmOpen.Tag = "FirstTime"
            frmOpen.Show
            Unload frmMapEdit
            Exit Sub
        Else
            Exit Sub
        End If
    Else
        Err.Raise Err.Number
    End If

End Sub
'checked to here
Public Sub mnuOpen_Click()
Dim Count As Long
Dim Temp As Long
Dim Dummy As Integer, DummyThing As Thing
Dim FirstTime As Boolean
'Dim Opener As String   'this is now global(I think)
    On Error GoTo ErrorHandler
    If frmOpen.Tag <> "" Then  'we're doing startup; the user must open a file or return
    'to frmOpen
        frmOpen.Tag = ""
        FirstTime = True
    End If
    
    CMDialog1.InitDir = App.Path   'so we don't start up in some
    'weird directory. Maybe eventually I will allow this to be
    'customized by the user.    (Note: take a consensus on whether this should be changed to
    'C:\My Documents
    CMDialog1.filename = ""
    CMDialog1.Filter = "Map editor files (*.map)|*.map|All files (*.*)|*.*"
    CMDialog1.FilterIndex = 1
    CMDialog1.ShowOpen
    If FirstTime = False Then
        Close Fileno
        Close ObjFileno
    End If
    Fileno = FreeFile
    'init variables
    ScreenX = 0
    ScreenY = 0
    TopX = 10
    TopY = 10

    Opener = CMDialog1.filename 'txtOpen.Text
    Open Opener For Random As #Fileno Len = Len(Dummy)
    
    DummyThing.Desc = -1
    DummyThing.Movement = 0
    DummyThing.Type = -1
    DummyThing.x = -1
    DummyThing.Y = -1
    ObjFileno = FreeFile
    ObjOpener = (Left$(Opener, Len(Opener) - 3) & "thi")
    Open ObjOpener For Random As #ObjFileno Len = Len(DummyThing)
    
    Get #Fileno, , MapXSize
    Get #Fileno, , MapYSize

    LoadMap Fileno, ScreenX, ScreenY
    LoadThings ObjFileno, ScreenX, ScreenY
        'size the scroll bars
    hsbJump.Max = MapXSize
    hsbJump.LargeChange = MapXSize \ 30 'divide the map into thirty pieces
    vsbJump.Max = MapYSize
    vsbJump.LargeChange = MapYSize \ 30

    frmMapEdit.Caption = Opener + " - Map Editor"
    frmMapEdit.Show
    Exit Sub
ErrorHandler:
    'User pressed the Cancel button
    If Err.Number = cdlCancel Then  'cancel code(probably a constant somewhere, but
        'VBHelp tells me to put in this number.)
        If FirstTime = True Then 'if you enter a value, then Cancel you'll still get the bug.
            frmOpen.Tag = "FirstTime" 'Tough!
            frmOpen.Show
            Unload frmMapEdit
            Exit Sub
        Else
            Exit Sub
        End If
    Else
        Err.Raise Err.Number
    End If
End Sub

Private Sub mnuProperties_Click()
    wSelect = Selected
    SelectX = CellX
    SelectY = CellY
    cmdProperties_Click
End Sub

Private Sub mnuRefreshThumbnail_Click()
    PaintThumbNail
End Sub

Private Sub mnuRemove_Click()
    wSelect = Selected
    SelectX = CellX
    SelectY = CellY
    cmdRemove_Click
End Sub

Private Sub mnuRequireClick_Click()
'flip the required boolean and the menu's check
    bRequireClick = Not bRequireClick
    mnuRequireClick.Checked = Not mnuRequireClick.Checked
End Sub

Private Sub mnuReset_Click()
    CursorXSize = 1
    CursorYSize = 1
    'that's all, folks!!
End Sub

Private Sub mnuSave_Click()
    If Fileno <> 0 Then
        If vbYes = MsgBox("This item is pointless since the map is constantly auto-saved. Save anyway?" _
        , vbYesNo, "Map Editor") Then
            SaveMap Fileno, ScreenX, ScreenY
            SaveThings ObjFileno, ScreenX, ScreenY
        End If
    End If
End Sub


Private Sub mnuScrollFaster_Click()
    With tmrScroll
    If .Interval > 100 Then .Interval = .Interval - 100
    End With
End Sub

Private Sub mnuScrollSlower_Click()
    With tmrScroll
        .Interval = .Interval + 100
    End With
End Sub

Private Sub mnuShowSolid_Click()
    bShowSolid = Not (bShowSolid)   'inverse the values inside the menu and the flag
    mnuShowSolid.Checked = Not (mnuShowSolid.Checked)
End Sub

Private Sub mnuToolTips_Click()
    bToolTips = Not (bToolTips) 'inverse the values inside the menu and the flag
    mnuToolTips.Checked = Not (mnuToolTips.Checked)
    picViewport.ToolTipText = ""    'make sure that we flush out the ToolTipText variable
    'because otherwise even when you turn ToolTips off, you get the last one you were
    'pointing at.
End Sub

Private Sub mnuUnselect_Click()
    wSelect = UNSELECTED    'reset selection variables
    SelectX = -1
    SelectY = -1
End Sub

Private Sub mnuWhatsThis_Click()
Dim Count As Integer
    IsThing CellX, CellY, Count 'this to get the Count value.
    'now tell the user the picture name
    MsgBox imlThings.ListImages(Things(Count).Type).Key, , "Description"
End Sub

Private Sub picThumbnail_Click()
    PaintThumbNail
End Sub

Private Sub picViewport_Click()
Dim CountX As Integer, CountY As Integer
    If wSelect = SELECTING Then
        SelectX = CellX 'set the selected variables so that the Object cmds can acces them.
        SelectY = CellY
        wSelect = Selected  'set the flag to tell PaintPic to draw SelHilight on the SelectedX,Y
        cmdAdd.Enabled = True   'enable the right buttons that were previously disabled
        cmdMove.Enabled = True
        cmdProperties.Enabled = True
        cmdRemove.Enabled = True
        Exit Sub
    ElseIf wSelect = Selected And CellX = SelectX And CellY = SelectY Then
        wSelect = UNSELECTED
        SelectX = -1
        SelectY = -1
        cmdAdd.Enabled = False   'disable the right buttons that were previously enabled
        cmdMove.Enabled = False
        cmdProperties.Enabled = False
        cmdRemove.Enabled = False

        Exit Sub
    ElseIf wSelect = MOVING Then
        SelectX = CellX
        SelectY = CellY
        cmdMove_Click
        Exit Sub
    End If
    If CursorXSize < 2 And CursorYSize < 2 Then 'this could probably be done more
    'efficiently using the same code for each.
        Map(((CellY - ScreenY) * 30) + (CellX - ScreenX)) = TerrainType
    Else
        If ((CellX - 1) + CursorXSize <= 10) And ((CellY - 1) + CursorYSize <= 10) Then
            For CountY = 0 To CursorYSize - 1 Step 1
                For CountX = 0 To CursorXSize - 1 Step 1
'                    Mid(Map(CellY + CountY), ((CellX + 1) + CountX), 1) = Temp
                    Map((((CellY - ScreenY) + CountY) * 30) + ((CellX - ScreenX) + CountX)) = TerrainType
                Next CountX
            Next CountY
        End If
    End If

End Sub


Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbKeyLButton Then
        bBlocking = True 'turn on drag-to-draw only if clicking
    'with the LEFT mouse button.
    End If
    If Button = vbKeyRButton Then   'they're (probably) clicking with the right mouse button.
        'something here to show a popup menu.
    Dim Dummy As Integer
        'set the appropriate menus to Visible
        
        'see if they get What's this help, moving, properties, removing
        '(i.e. if they clicked on an object)
        If IsThing(CellX, CellY, Dummy) Then
            mnuWhatsThis.Visible = True
            mnuMove.Visible = True
            mnuProperties.Visible = True
            mnuRemove.Visible = True
        Else    'nope, all they get is add
            mnuWhatsThis.Visible = False
            mnuMove.Visible = False
            mnuProperties.Visible = False
            mnuRemove.Visible = False
        End If
        
    '    mnuAdd.Visible = True  'always Visible
        
                'now see if we should show mnuUnselect
        If wSelect = Selected Then mnuUnselect.Visible = True
        If wSelect = MOVING Then
            mnuUnselect.Caption = "&Cancel Move"
        Else
            mnuUnselect.Caption = "&Unselect All"
        End If

        PopupMenu mnuContext
    End If
End Sub

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static PointX As Integer 'two values that keep track of the cellx,y relative to the viewports
    Static PointY As Integer   'top,left; not the map's top, left.
    Static OldPointX As Integer, OldPointY As Integer
    CellX = x \ 32
    CellY = Y \ 32
    OldPointX = PointX
    OldPointY = PointY
    PointX = CellX  'two values that keep track of the cellx,y relative to the viewports
    PointY = CellY   'top,left; not the map's top, left.
    
    CellX = CellX + TopX + ScreenX  'but here we add that to make the map's top, left.
    CellY = CellY + TopY + ScreenY

    If bBlocking = True And ((PointX - 1) + CursorXSize < 10) And ((PointY - 1) + CursorYSize < 10) Then
    'don't know why I used this name...stole it from Prog Win95
    'i think. Anyway, the user is dragging the mouse across picViewport, so we'll let him
    'draw.
        Dim CountX As Long, CountY As Long  'for speed(on 32-bit systems, anyway) just joking. actually this is required due to the possible size of the map.
        'NEW: these might not HAVE to be Long, but just to be safe...besides, they *are* faster. I tested them once.
        If CursorXSize < 2 And CursorYSize < 2 Then 'just set 1 tile at a time.
        
            Map(((CellY - ScreenY) * 30) + (CellX - ScreenX)) = TerrainType
        Else
            
            For CountY = 0 To CursorYSize - 1 Step 1    'do it this way only if they're drawing
                For CountX = 0 To CursorXSize - 1 Step 1 'drawing large blocks
                    Map((((CellY - ScreenY) + CountY) * 30) + ((CellX - ScreenX) + CountX)) = TerrainType
                Next CountX
            Next CountY
        End If
'        If X < 16 Then
'            iScrolling = WEST
'            tmrScroll.Enabled = True
'        ElseIf X > 304 Then
'            iScrolling = EAST
'            tmrScroll.Enabled = True
'        ElseIf Y < 16 Then
'            iScrolling = NORTH
'            tmrScroll.Enabled = True
'        ElseIf Y > 304 Then
'            iScrolling = SOUTH
'            tmrScroll.Enabled = True
'        Else
'            iScrolling = STOPPED
'            tmrScroll.Enabled = False
'        End If
    
    End If



    Dim TerrainVal As Integer
    TerrainVal = Map((((CellY - ScreenY)) * MAP_ARRAYX) + ((CellX - ScreenX)))
    If TerrainVal = -1 And bToolTips Then 'this blank checker for drawing the map; the game itself uses a faster PaintViewPort function and will bug up if you
    'leave any true blanks in; you *have* to draw with the blank tile; it's actually tile number 17, not -1 like true blank.
        picViewport.ToolTipText = "Blank"
    ElseIf bToolTips Then
        picViewport.ToolTipText = imlTerrain.ListImages(Map((((CellY - ScreenY)) * MAP_ARRAYX) + ((CellX - ScreenX)))).Key
    ElseIf Not bToolTips Then
        picViewport.ToolTipText = ""
    End If
    If OldPointX <> PointX Or OldPointY <> PointY Then  'only paint if we HAVE to.
        PaintPicViewport   'wa ha ha! test what happens if we leave this out! Later: oops...that didn't work like I thought it would...
    End If
End Sub

Private Sub picViewport_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbKeyLButton Then bBlocking = False   'turn off drag-to-draw with the Left
    'mouse button.
    'maybe put a call to PaintThumbnail here(depending on how fast it is)
    PaintThumbNail
    iScrolling = Stopped
    tmrScroll.Enabled = False
End Sub

Private Sub tmrScroll_Timer()
    If iScrolling = West Then
        cmdWest_Click
    ElseIf iScrolling = East Then
        cmdEast_Click
    ElseIf iScrolling = South Then
        cmdSouth_Click
    ElseIf iScrolling = North Then
        cmdNorth_Click
    End If
End Sub

Private Sub PaintThumbNail()
'This sub paints the little 90x90 pixel picture box with psychedelic colors supposed to represent the different tiles.
'This is so the user gets a sort of a 'radar'. This is because Rachel was so insistent that she get a map...

'to determine the colors, I use this formula: B = (Index Mod 32) * 8                    I could have done the red first, but I didn't! :)
'                                                                G = ((Index \ 32) Mod 32) * 8
'                                                               R = ((Index \ 32 * 32) Mod 32) * 8
'this formula produces exactly 32767 or 65536 colors(I can't remember which) Either way, each tile will have its own unique color.
    'first loop through all the tiles in the map array.
Dim XCount As Long
Dim YCount As Long
Dim intIndex As Integer
Dim R As Byte, G As Byte, B As Byte 'can you guess what these do??
    For YCount = 0 To MAP_ARRAYY - 1 Step 1
        For XCount = 0 To MAP_ARRAYX - 1 Step 1
            intIndex = Map((YCount * MAP_ARRAYX) + XCount)
            'get b,g,r here
            If intIndex < 0 Then
                R = 255
                G = 255
                B = 255
            Else
                B = (intIndex Mod 32) * 8
                G = ((intIndex \ 32) Mod 32) * 8
                R = ((intIndex \ 1024) Mod 32) * 8
            End If
            
            'now paint a line .... BF that is 3 by 3
            picThumbnail.Line (XCount * 3, YCount * 3)-((XCount * 3) + 2, (YCount * 3) + 2), RGB(R, G, B), BF
            'that's all, folks!
        Next XCount
    Next YCount
    'now paint the objects:
    For intIndex = 0 To 89 Step 1   'note the re-used variable names :)
         XCount = Things(intIndex).x - ScreenX
         YCount = Things(intIndex).Y - ScreenY
         picThumbnail.PSet ((XCount * 3) + 1, (YCount * 3) + 1), RGB(255, 0, 0)
    Next intIndex
End Sub

Private Sub vsbJump_Change()
'here put the jump code from cmdJump

Dim lngJumpY As Long
    If vsbJump.Value > ScreenY And vsbJump.Value < ScreenY + 20 Then  'we just need to move the screen a little
        TopY = vsbJump.Value - ScreenY
        PaintPicViewport
    Else    'need a real jump.
        lngJumpY = (vsbJump.Value \ 10) * 10   'get the screen coordinates
        
        If lngJumpY > 9 And lngJumpY < (MapYSize - 20) Then
            'MsgBox "Transmitting jump data at 16.6 M/Sec."  'more Warped Star Truck feedback
            SaveMap Fileno, ScreenX, ScreenY
            ScreenY = lngJumpY - 10
            TopY = vsbJump.Value - ScreenY
            LoadMap Fileno, ScreenX, ScreenY
            LoadThings ObjFileno, ScreenX, ScreenY
            PaintPicViewport
            'MsgBox "We have arrived safely at (" & JumpX & ", " & JumpY & "), Caption!"
        Else
            MsgBox "Error, Caption! Invalid Jump Co-ordinates!"
        End If
    End If
End Sub

Private Sub vsbJump_Scroll()
    lblStatus.Caption = "Y: " & vsbJump.Value
End Sub
