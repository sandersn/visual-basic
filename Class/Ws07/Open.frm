VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Picture for Database"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2520
      Pattern         =   "*.bmp;*.gif;*.wmf;*.tif;*.jpg;*.pcx;*.ico"
      TabIndex        =   6
      Top             =   2835
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   900
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return and Set"
      Default         =   -1  'True
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    frmMusicHdr.Show
End Sub

Private Sub cmdClear_Click()
    imgPreview.Picture = LoadPicture("")
End Sub

Private Sub cmdReturn_Click()
    frmMusicHdr.imgCover.Picture = frmOpen.imgPreview.Picture
    frmMusicHdr.cmdChange_Click '.Value = True
    frmMusicHdr.cmdSave_Click '.Value = True
    Unload Me
    frmMusicHdr.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Path
End Sub

Private Sub File1_Click()
Dim FullPath As String
    FullPath = File1.Path & "\" & File1.filename
    imgPreview.Picture = LoadPicture(FullPath)
End Sub

Private Sub Form_Load()
    frmMusicHdr.Hide
End Sub
