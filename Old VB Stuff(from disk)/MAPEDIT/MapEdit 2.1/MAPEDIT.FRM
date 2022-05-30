VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmMapEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor for Chrysalis"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   1335
   ClientWidth     =   10890
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
   Icon            =   "MAPEDIT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   726
   Begin ComctlLib.ListView lvwTerrain 
      Height          =   4800
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   8467
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327680
      Icons           =   "imlTerrain"
      SmallIcons      =   "imlTerrainSm"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MouseIcon       =   "MAPEDIT.frx":0442
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   6840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton cmdEast 
      Appearance      =   0  'Flat
      Caption         =   ">"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton cmdSouth 
      Appearance      =   0  'Flat
      Caption         =   "\/"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton cmdNorth 
      Appearance      =   0  'Flat
      Caption         =   "/\"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton cmdWest 
      Appearance      =   0  'Flat
      Caption         =   "<"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox picViewport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   360
      MousePointer    =   4  'Icon
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin ComctlLib.ImageList imlTerrainSm 
      Left            =   7440
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   23
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":045E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0570
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0682
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":09B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":0F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":1024
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":112E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":1240
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":1352
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":1894
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":1DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":2328
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":2DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":311E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":3470
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":37C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlTerrain 
      Left            =   6120
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   23
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":3B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":3E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":4148
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":4462
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":477C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":4A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":4DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":50CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":53E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":56FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":5A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":5D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":5FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":62DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":65F8
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":6E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":767C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":7ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":8720
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":8F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":9BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":A816
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MAPEDIT.frx":AAA8
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   705
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuIconSize 
         Caption         =   "Small &Icons"
      End
      Begin VB.Menu mnuCursorSize 
         Caption         =   "Cursor &Size"
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
'doesn't work out.
'Anyway, it will be nice to have to laugh at.

'version 2: many years later.
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
'at startup and shutdown of the program. I anticipate that this will increase perform-
'ance somewhat. Note: Completed! But it doesn't have any noticable impact on my system.
    '6.I changed the colors on the controls to system colors.
    '7.(07/21)Added just about all of the error checking. Note: the program still
'crashes when you move the mouse outside the picture box while still drawing if
'the cursor is outside the range of the 30*30 Map array(i.e. near the top, left,
'bottom, or right; especially when using a large cursor size).
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
Dim TerrainType As Integer
Dim ScreenX As Long
Dim ScreenY As Long
Dim TopX As Integer
Dim TopY As Integer
Dim CellY As Integer
Dim CellX As Integer
Dim bBlocking As Boolean    'so that you can drag the mouse to paint
'Dim MapType As String * 3 'won't need this anymore with more >256(up to 32K) tiles
'Dim Filename As String * 8
Dim CursorXSize As Integer   'I had forgotten about this cool function
Dim CursorYSize As Integer
Dim Fileno As Integer
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Const SRCCOPY = &HCC0020

Private Sub cmdEast_Click()
'Dim Dummy As Integer ' As Long
    TopX = TopX + 1
    CellX = CellX + 1
    If TopX = 21 Then
        TopX = 20
        CellX = CellX - 1
    End If
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    'DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

'    If TopX > (MapXSize - 9) Then
'        TopX = TopX - 1
'        CellX = CellX - 1
'    End If
  
'    Dummy = BitBlt(picViewport.hDC, 0, 0, 320, 320, picCanvas.hDC, TopX * 32, TopY * 32, SRCCOPY)
End Sub


Private Sub cmdNorth_Click()
''Dim Dummy As Integer ' As Long
    TopY = TopY - 1
    CellY = CellY - 1
    If TopY = -1 Then
        TopY = 0
        CellY = CellY + 1
    End If
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    'DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

'    If TopY < 0 Then
'        TopY = TopY + 1
'        CellY = CellY + 1
'    End If
'Dummy = BitBlt(picViewport.hDC, 0, 0, 320, 320, picCanvas.hDC, TopX * 32, TopY * 32, SRCCOPY)

End Sub

Private Sub cmdSouth_Click()
''Dim Dummy As Integer 'As Long
    TopY = TopY + 1
    CellY = CellY + 1
    If TopY = 21 Then
        TopY = 20
        CellY = CellY - 1
    End If
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    'DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

'    If TopY > (MapYSize - 9) Then
'        TopY = TopY - 1
'        CellY = CellY - 1
'    End If
'Dummy = BitBlt(picViewport.hDC, 0, 0, 320, 320, picCanvas.hDC, TopX * 32, TopY * 32, SRCCOPY)

End Sub

Private Sub cmdWest_Click()
'Dim Dummy As Integer 'As Long
    TopX = TopX - 1
    CellX = CellX - 1
    If TopX = -1 Then
        TopX = 0
        CellX = CellX + 1
    End If
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    'DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

'        TopX = TopX + 1
'        CellX = CellX + 1
'    End If
'    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY
'Dummy = BitBlt(picViewport.hDC, 0, 0, 320, 320, picCanvas.hDC, TopX * 32, TopY * 32, SRCCOPY)

End Sub

Private Sub Form_Paint()
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    'DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32,
    'CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

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
'checked to here
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
    With lvwTerrain 'I'm not sure how much time I'm saving by using the With keyword
    'Bob claims it is a lot, but I can't tell the difference(I need to buy a slower machine)(Ha, ha)
    For Count = 1 To imlTerrain.ListImages.Count
        .ListItems.Add Count, , , imlTerrain.ListImages(Count).Index, imlTerrain.ListImages(Count).Index
    Next Count  'boy it took me a long time to figure out how to load the image in from the
    'ImageList(unfortunately, Bob will probably teach it next class)
    End With
'    mnuNew_Click
'    PaintMap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
End Sub



Private Sub Form_Unload(Cancel As Integer)
If Fileno <> 0 And frmOpen.Tag = "" Then    'if we've got an open file and aren't
'still trying to get the user to open or create a file from frmOpen
    If MsgBox("Save Changes?", vbYesNo + vbQuestion, "Save Map") = vbYes Then
        SaveMap Fileno, ScreenX, ScreenY
    End If
    Close Fileno
    End
End If

End Sub

Private Sub lvwTerrain_ItemClick(ByVal Item As ComctlLib.ListItem)
    TerrainType = Item.Index
End Sub

Private Sub mnuClear_Click()
Dim Count As Long
    If vbNo = MsgBox("This could take a long time -- as much time a creating a new, large map does. Do you still want to Clear?", vbYesNo + vbQuestion, "Clear") Then
        Exit Sub
    End If
    
    'ReDim Map(MapYSize) As String '* MapYSize
    For Count = 0 To MapYSize * MapXSize Step 1
        Put #Fileno, , -1
    Next Count
    'picCanvas.Clear
    TopX = 0
    TopY = 0
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

End Sub

Private Sub mnuCursorSize_Click()
Dim Temp As String 'so it's slower. who cares?
    Temp = InputBox("Enter cursor X size: (Do not go over ten)", "Cursor Size", "1")
    If Temp = "" Then CursorXSize = 1
    CursorXSize = CInt(Temp)
    If CursorXSize < 1 Then CursorXSize = 1
    If CursorXSize > 10 Then CursorXSize = 10
    Temp = InputBox("Enter cursor Y size: (Do not go over ten)", "Cursor Size", "1")
    If Temp = "" Then CursorYSize = 1
    CursorYSize = CInt(Temp)
    If CursorYSize < 1 Then CursorYSize = 1
    If CursorYSize > 10 Then CursorYSize = 10
End Sub

Private Sub mnuExit_Click()
    If Fileno <> 0 Then
    If MsgBox("Save Changes?", vbYesNo + vbQuestion, "Save Map") = vbYes Then
        SaveMap Fileno, ScreenX, ScreenY
    End If
    Close Fileno
    End If
    End
End Sub

Private Sub mnuIconSize_Click()
    If mnuIconSize.Checked = False Then
        lvwTerrain.View = lvwSmallIcon
        mnuIconSize.Checked = True
    Else
        lvwTerrain.View = lvwIcon
        mnuIconSize.Checked = False
    End If
End Sub

Public Sub mnuNew_Click()
Dim Count As Long
Dim Dummy As Integer
Dim FirstTime As Boolean
'Dim Opener As String 'this is now global(I think)
    On Error GoTo ErrHandler
    If frmOpen.Tag <> "" Then
        frmOpen.Tag = ""
        FirstTime = True   'the user didn't actually click the New menu(but the Open form)
    End If
    
    CMDialog1.Filter = "Map Files (*.map)|*.map|All Files (*.*)|*.*"
    CMDialog1.filename = "Untitled.map"
    CMDialog1.ShowSave
'    If CMDialog1.filename = "Untitled.map" Then 'check for Cancel
'        If FirstTime = True Then
'            frmOpen.Show
'            Unload frmMapEdit
'            Exit Sub
'        Else
'            Exit Sub
'        End If
'        MsgBox "You must enter a value!", vbOKOnly, "Error"
'        mnuNew_Click
'    End If
    Opener = CMDialog1.filename 'OK we've got a good file name, and the user doesn't want to
    'cancel
    
    'a little error checking(not much yet)
    If Dir(Opener) = Opener Then    'see if the user wants to overwrite the file(doesn't seem
                            'to be working yet) but definitely not critical because I think
                            'that the WINDOWS way is not to ask the user.
        Dim Answer As Long
        Answer = MsgBox("Do you want to overwrite the current file?", vbQuestion + vbYesNo _
        , "Map Edit")
        If Answer = vbYes Then Kill Opener
    End If
    
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
    End If
    
    Fileno = FreeFile   'continue opening the file
    Open Opener For Random As #Fileno Len = Len(Dummy) 'dummy is an integer since I can't
                                                        'remember what VB's sizeof looks like
        'Put map size at beginning  (sadly, I have figured out how to do it the way I wanted
        'to originally, but it is working this way, so who cares?)
    Put #Fileno, , MapXSize
    Put #Fileno, , MapYSize
'fill the map with -1(nothing)
    For Count = 3 To MapYSize * MapXSize Step 1 'start at third position(offset from MapX,YSize)
            Put #Fileno, Count, -1
    Next Count
    'init variables
    ScreenX = 0
    ScreenY = 0
    TopX = 10
    TopY = 10

    'we're done
    LoadMap Fileno, ScreenX, ScreenY
    If FirstTime = True Then
        frmMapEdit.Show
    End If
'Ready to show the user now!
'    picTerrain(0).Picture = LoadPicture(App.Path & "\forest1.ico")
'    picTerrain(1).Picture = LoadPicture(App.Path & "\water1.ico")
'    picTerrain(2).Picture = LoadPicture(App.Path & "\grass1.ico")
'    picTerrain(3).Picture = LoadPicture(App.Path & "\gravel1.ico")
'    picTerrain(4).Picture = LoadPicture(App.Path & "\mnt1.ico")
    frmMapEdit.Caption = Opener + " - Map Editor"
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

ErrHandler:
    'User pressed the Cancel button
    If Err.Number = 32755 Then  'cancel code(probably a constant somewhere, but
        'VBHelp tells me to put in this number.)
        If FirstTime = True Then    'a tag set when the user pressed a command button back
            'on frmOpen (this for correct timing)
            frmOpen.Tag = "FirstTime"
            frmOpen.Show
            Unload frmMapEdit
            Exit Sub
        Else
            Exit Sub
        End If
    End If

End Sub
'checked to here
Public Sub mnuOpen_Click()
Dim Count As Long
Dim Temp As Long
Dim Dummy As Integer
Dim FirstTime As Boolean
'Dim Opener As String   'this is now global(I think)
    On Error GoTo ErrorHandler
    If frmOpen.Tag <> "" Then  'we're doing startup; the user must open a file or return
    'to frmOpen
        frmOpen.Tag = ""
        FirstTime = True
    End If
    CMDialog1.filename = "" 'this to help me detect Cancel
    'more easily
    CMDialog1.Filter = "Map editor files (*.map)|*.map|All files (*.*)|*.*"
    CMDialog1.FilterIndex = 1
    CMDialog1.ShowOpen
'    If CMDialog1.filename = "Untitled.map" Then 'oops, they canceled(although if you enter a
'    End If
    If FirstTime = False Then
        Close Fileno
    End If
    Fileno = FreeFile
    'init variables
    ScreenX = 0
    ScreenY = 0
    TopX = 10
    TopY = 10

    Opener = CMDialog1.filename 'txtOpen.Text
    Open Opener For Random As #Fileno Len = Len(Dummy)
    
    Get #Fileno, , MapXSize
    Get #Fileno, , MapYSize

    LoadMap Fileno, ScreenX, ScreenY
    frmMapEdit.Show
    'Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
    CursorXSize * 32, CursorYSize * 32


    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY
    frmMapEdit.Caption = Opener + " - Map Editor"
    'Man:We don't need the Picture boxes anymore, and we are loading the pictures in at
    'Form_Load. If you want to see the load routine go there.
    'Crono:OK. Thanks.
    'Man:Bye
    'Nadia:Crono! I think we should look into this.
'        picTerrain(0).Picture = LoadPicture(App.Path & "\forest1.ico")
'        picTerrain(1).Picture = LoadPicture(App.Path & "\water1.ico")
'        picTerrain(2).Picture = LoadPicture(App.Path & "\grass1.ico")
'        picTerrain(3).Picture = LoadPicture(App.Path & "\gravel1.ico")
'        picTerrain(4).Picture = LoadPicture(App.Path & "\mnt1.ico")
    'Now paint the map
ErrorHandler:
    'User pressed the Cancel button
    If Err.Number = 32755 Then  'cancel code(probably a constant somewhere, but
        'VBHelp tells me to put in this number.)
        If FirstTime = True Then 'value, then Cancel you'll still get the bug...tough!
            frmOpen.Tag = "FirstTime"
            frmOpen.Show
            Unload frmMapEdit
            Exit Sub
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub mnuSave_Click()
    If Fileno <> 0 Then
        If vbYes = MsgBox("This item is pointless since the map is constantly auto-saved. Save anyway?" _
        , vbYesNo, "Map Editor") Then
            SaveMap Fileno, ScreenX, Screen
        End If
    End If
End Sub

Private Sub picTerrain_Click(Index As Integer)
    TerrainType = Index
End Sub

Private Sub picViewport_Click()
Dim CountX, CountY
        If CursorXSize < 2 And CursorYSize < 2 Then
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

Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bBlocking = True 'turn on drag-to-draw
End Sub

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CellX = X \ 32
    CellY = Y \ 32
    CellX = CellX + TopX + ScreenX
    CellY = CellY + TopY + ScreenY
    
    If bBlocking = True And ((CellX - TopX - ScreenX - 1) + CursorXSize < 10) And ((CellY - TopY - ScreenY - 1) + CursorYSize < 10) Then
    'don't know why I used this name...stole it from Prog Win95
    'i think. Anyway, the user is dragging the mouse across picViewport, so we'll let him
    'draw.
        Dim CountX As Long, CountY As Long  'for speed(on 32-bit systems, that is) just joking. actually this is required due to the possible size of the map.
        If CursorXSize < 2 And CursorYSize < 2 Then 'just set 1 tile at a time.
        
            Map(((CellY - ScreenY) * 30) + (CellX - ScreenX)) = TerrainType
        Else
            
            For CountY = 0 To CursorYSize - 1 Step 1    'do it this way only if they're drawing
                For CountX = 0 To CursorXSize - 1 Step 1 'drawing large blocks
                    Map((((CellY - ScreenY) + CountY) * 30) + ((CellX - ScreenX) + CountX)) = TerrainType
                Next CountX
            Next CountY
         End If
    End If
    Paintmap picViewport, imlTerrain, Fileno, TopX, TopY, TopX, TopY, ScreenX, ScreenY, CellX, CellY
    If bBlocking Then
        DrawSelHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
        CursorXSize * 32, CursorYSize * 32
    Else
        DrawHighLight picViewport, (CellX - TopX - ScreenX) * 32, (CellY - TopY - ScreenY) * 32, _
        CursorXSize * 32, CursorYSize * 32
    End If

    'tell the user his position
    lblPosition.Caption = "X: " & CellX & " Y: " & CellY & " TopX = " & TopX & " TopY = " & TopY

End Sub

Private Sub picViewport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bBlocking = False   'turn off drag-to-draw
End Sub
