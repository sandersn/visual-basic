VERSION 5.00
Begin VB.Form frmMusicHdr 
   Caption         =   "Music Selection"
   ClientHeight    =   3300
   ClientLeft      =   2430
   ClientTop       =   2040
   ClientWidth     =   4740
   Icon            =   "MusicHdr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtYearReleased 
      DataField       =   "YearReleased"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtLabel 
      DataField       =   "RecordLabel"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   7
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtArtist 
      DataField       =   "Artist"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtCollectionID 
      DataField       =   "CollectionID"
      DataSource      =   "datMusicHdr"
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Data datMusicHdr 
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\Visual Basic\Class\Ws02\CDLib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Music_Hdr"
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label lblYearReleased 
      AutoSize        =   -1  'True
      Caption         =   "Year Released:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Recording Label:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   345
   End
   Begin VB.Label lblArtist 
      AutoSize        =   -1  'True
      Caption         =   "Artist/Group:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lblCollectionID 
      AutoSize        =   -1  'True
      Caption         =   "Collection ID:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmMusicHdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    datMusicHdr.Recordset.AddNew 'Add the record
    txtYearReleased.Text = Year(Now)  'I can't get the Year working yet
    'Note: just found the function(Now) in passing while looking
    'through VB Books Online. Who says you don't see useful stuff
    'while looking for something else??
    txtYearReleased.SelStart = 0 'Provide some defaults for the user
    txtYearReleased.SelLength = Len(txtYearReleased.Text)
    cmdSave.Enabled = True  'Only allow save after adding
    cmdAdd.Enabled = False 'Can't double add
    cmdChange.Enabled = False 'Can't edit record if you're already
    'adding one!!
    cmdDelete.Enabled = False 'Can't delete an incomplete add
    txtCollectionID.SetFocus  'put Cursor in the first field for entering
    'information
    'Only unlock the text boxes if the user
    txtCollectionID.Locked = False 'presses Add or Edit
    txtArtist.Locked = False
    txtTitle.Locked = False
    txtLabel.Locked = False
    txtYearReleased.Locked = False
End Sub

Private Sub cmdChange_Click()
    datMusicHdr.Recordset.Edit
    cmdSave.Enabled = True  'CAN save after done editing
    cmdChange.Enabled = False 'can't change while already changing
    cmdAdd.Enabled = False 'can't add in the middle of changing
    
    txtArtist.SetFocus  'This time I'll move down artist because the CollectionID
    'is *supposed* to already have been set.(Boy is that verb tense confusing)
    Rem does rem still work??
    Rem Boy!! I can't believe REM still works!!
    'Only unlock the text boxes if the user
    txtCollectionID.Locked = False 'presses Add or Edit
    txtArtist.Locked = False
    txtTitle.Locked = False
    txtLabel.Locked = False
    txtYearReleased.Locked = False

End Sub

Private Sub cmdDelete_Click()
    datMusicHdr.Recordset.Delete ' Delete and then close and open the DB again
    datMusicHdr.Refresh
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSave_Click()
    datMusicHdr.Recordset.Update
    cmdSave.Enabled = False 'can't save AGAIN
    cmdDelete.Enabled = True 'but you can delete
    cmdAdd.Enabled = True 'you can also add...
    cmdChange.Enabled = True 'or change the current record
    
    cmdAdd.SetFocus 'If you don't do this, the focus jumps to Delete...Fun!!
    
    'Disable the text boxes unless the user presses
    txtCollectionID.Locked = True 'Add or Edit
    txtArtist.Locked = True
    txtTitle.Locked = True
    txtLabel.Locked = True
    txtYearReleased.Locked = True

End Sub
