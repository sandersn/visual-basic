VERSION 5.00
Begin VB.Form frmMusicHdr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Selection"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmMusicHdr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtYearReleased 
      DataField       =   "YearReleased"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtLabel 
      DataField       =   "RecordLabel"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   8
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox txtArtist 
      DataField       =   "Artist"
      DataSource      =   "datMusicHdr"
      Height          =   255
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   6
      Top             =   840
      Width           =   5055
   End
   Begin VB.TextBox txtCollectionID 
      DataField       =   "CollectionID"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Data datMusicHdr 
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\Visual Basic\Class\Derek\Ch2\CDLibe.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Music_Hdr"
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblYearReleased 
      AutoSize        =   -1  'True
      Caption         =   "Year Released:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Recording Label:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Volume Title:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lblArtist 
      AutoSize        =   -1  'True
      Caption         =   "Artist / Group:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   990
   End
   Begin VB.Label CollectionID 
      AutoSize        =   -1  'True
      Caption         =   "Colection ID:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "frmMusicHdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'notes: 1.need to find some way to add the recordnumber when you press add.
'2. Need to use Option Explicit!
'3. After doing 1, set Locked to true for txtCollectionID
Private Sub cmdAdd_Click()
datMusicHdr.Recordset.AddNew
txtYearReleased.Text = 1996
txtYearReleased.SelStart = 0
txtYearReleased.SelLength = Len(txtYearReleased.Text)
cmdSave.Enabled = True
cmdAdd.Enabled = False
cmdChange.Enabled = False
cmdDelete.Enabled = False
txtArtist.SetFocus
End Sub

Private Sub cmdChange_Click()
datMusicHdr.Recordset.Edit
cmdSave.Enabled = True
cmdAdd.Enabled = False
cmdChange.Enabled = False
cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
datMusicHdr.Recordset.Delete
datMusicHdr.Refresh
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdSave_Click()
datMusicHdr.Recordset.Update
cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdChange.Enabled = True
cmdDelete.Enabled = True
End Sub


