VERSION 5.00
Begin VB.Form frmMusicHdr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Selection"
   ClientHeight    =   4695
   ClientLeft      =   2415
   ClientTop       =   2025
   ClientWidth     =   4740
   Icon            =   "MusicHdr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
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
   Begin VB.Label lblHint 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   4575
      WordWrap        =   -1  'True
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
Dim SavePressed As Boolean
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
'    'Only unlock the text boxes if the user
'    txtCollectionID.Locked = False 'presses Add or Edit
'    txtArtist.Locked = False
'    txtTitle.Locked = False
'    txtLabel.Locked = False
'    txtYearReleased.Locked = False
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled = True Then
        lblHint.Caption = "Begin adding a new database record."
    End If
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
    'Only unlock the text boxes if the user presses Add or Edit
'    txtCollectionID.Locked = True   'note that we don't unlock the collectionID unless we're
'    txtArtist.Locked = False    'adding a new field(even then I would like to automatically
'    txtTitle.Locked = False     'put it in there and STILL not ley the user edit it.
'    txtLabel.Locked = False
'    txtYearReleased.Locked = False

End Sub

Private Sub cmdChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdChange.Enabled = True Then
        lblHint.Caption = "Edit the current record. If you do this, then edit the record, then click save, you will not get the annoying error message asking you to save your data."
    End If
End Sub

Private Sub cmdDelete_Click()
    datMusicHdr.Recordset.Delete ' Delete and then close and open the DB again
    datMusicHdr.Refresh
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdDelete.Enabled = True Then
        lblHint.Caption = "Delete a record. Be careful-- this cannot be undone!"
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "Don't look to me for help on this!"
End Sub

Private Sub cmdSave_Click()
    SavePressed = True
    datMusicHdr.Recordset.Update
    cmdSave.Enabled = False 'can't save AGAIN
    cmdDelete.Enabled = True 'but you can delete
    cmdAdd.Enabled = True 'you can also add...
    cmdChange.Enabled = True 'or change the current record
    
    cmdAdd.SetFocus 'If you don't do this, the focus jumps to Delete...Fun!!
'    'Disable the text boxes unless the user presses
'    txtCollectionID.Locked = True 'Add or Edit
'    txtArtist.Locked = True
'    txtTitle.Locked = True
'    txtLabel.Locked = True
'    txtYearReleased.Locked = True      'disable my hack(it was fun while it lasted)

End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSave.Enabled = True Then
        lblHint.Caption = "Save the current record. This only works if you have clicked on Add or Edit. If you simply started changing fields, you must click the left or right buttons on the database control."
    End If
End Sub

Private Sub datMusicHdr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "Click the far left and right buttons to move to the beginning and end of the database, respectively. Click the inner left and right buttons to move one record forward or backward."
End Sub

Private Sub datMusicHdr_Validate(Action As Integer, Save As Integer)
    If SavePressed = True Then  'the user has already pressed save.
        Save = SavePressed
    Else
        If Save = True Then 'but if he hasn't...
        Dim Ans As Integer
            Ans = MsgBox("Data changed. Want to save?", _
            vbYesNo + vbExclamation, "Data not saved")  'we need to find out if he meant to.
            If Ans = vbNo Then
                Save = False
            End If
        End If
    End If
    SavePressed = False 'always reset save indicator
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "Ready"
End Sub

Private Sub lblArtist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The name of the artist or group who performed this piece of 'music'"
End Sub

Private Sub lblCollectionID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The number of the current record inside the database."

End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The company who funded the production of this piece of 'music'."
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The name of this piece of 'music'."

End Sub

Private Sub lblYearReleased_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The year that this piece of 'music' was available to the general public. For some reason, Bob said not to let you enter a year earlier than 1900. Why, I wonder? What if you want to enter some early music recorded on wax cylinders in the 1890's? What if you want to make up FAKE songs?"

End Sub

Private Sub txtArtist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The name of the artist or group who performed this piece of 'music'"

End Sub

Private Sub txtCollectionID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The number of the current record inside the database."

End Sub

Private Sub txtLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The company who funded the production of this piece of 'music'."
End Sub

Private Sub txtTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The name of this piece of 'music'."

End Sub

Private Sub txtYearReleased_LostFocus()
    If txtYearReleased.Text < 1900 Then
        Beep '(vbQuestion)
        MsgBox "Invalid Year", , "Big, Bad, Ugly DataBase Error!!!!"
        txtYearReleased.SelStart = 0
        txtYearReleased.SelLength = Len(txtYearReleased.Text)
        txtYearReleased.SetFocus
        
    End If
End Sub

Private Sub txtYearReleased_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHint.Caption = "The year that this piece of 'music' was available to the general public. For some reason, Bob said not to let you enter a year earlier than 1900. Why, I wonder? What if you want to enter some early music recorded on wax cylinders in the 1890's? What if you want to make up FAKE songs?"
End Sub
