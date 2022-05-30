VERSION 5.00
Begin VB.Form frmZackAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About The Knight in Shiny Armor"
   ClientHeight    =   4485
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5205
   ControlBox      =   0   'False
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
   ScaleHeight     =   4485
   ScaleWidth      =   5205
   Begin VB.PictureBox picZackMan 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   -360
      MousePointer    =   5  'Size
      Picture         =   "ZACKABOT.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   $"ZACKABOT.frx":0272
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "frmZackAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LeftEdge
Dim TopEdge

Private Sub cmdOk_Click()
    frmZackAbout.Hide
End Sub

Private Sub Form_Activate()
    Do
        Do Until picZackMan.Left > frmZackAbout.ScaleWidth - picZackMan.Width
            picZackMan.Picture = frmZackman.picZackrmid.Picture
            LeftEdge = LeftEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackrmov.Picture
            LeftEdge = LeftEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackrmid.Picture
            LeftEdge = LeftEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackrstl.Picture
            LeftEdge = LeftEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            DoEvents
        Loop
        picZackMan.Left = frmZackAbout.ScaleWidth - picZackMan.Width
        Do Until picZackMan.Top < 0
            picZackMan.Picture = frmZackman.picZackumid.Picture
            TopEdge = TopEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackumov.Picture
            TopEdge = TopEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackumid.Picture
            TopEdge = TopEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackustl.Picture
            TopEdge = TopEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            DoEvents
        Loop
        picZackMan.Top = 0
        Do Until picZackMan.Left < 0
            picZackMan.Picture = frmZackman.picZacklmid.Picture
            LeftEdge = LeftEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZacklmov.Picture
            LeftEdge = LeftEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZacklmid.Picture
            LeftEdge = LeftEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZacklstl.Picture
            LeftEdge = LeftEdge - 30
            picZackMan.Move LeftEdge, TopEdge
            DoEvents
        Loop
        picZackMan.Left = 0
        Do Until picZackMan.Top > frmZackAbout.ScaleWidth - picZackMan.Width
            picZackMan.Picture = frmZackman.picZackdmid.Picture
            TopEdge = TopEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackdmov.Picture
            TopEdge = TopEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackdmid.Picture
            TopEdge = TopEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            picZackMan.Picture = frmZackman.picZackdstl.Picture
            TopEdge = TopEdge + 30
            picZackMan.Move LeftEdge, TopEdge
            DoEvents
        Loop
        picZackMan.Top = frmZackAbout.ScaleHeight - picZackMan.Height
    Loop
End Sub

Private Sub Form_Load()
    LeftEdge = picZackMan.Left
    TopEdge = picZackMan.Top
End Sub

