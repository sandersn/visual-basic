VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMusicHdr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Selection"
   ClientHeight    =   5400
   ClientLeft      =   2415
   ClientTop       =   2310
   ClientWidth     =   7560
   Icon            =   "MusicHdr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7560
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      Max             =   16
   End
   Begin ComctlLib.StatusBar sbrHint 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   5145
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "MusicHdr.frx":0442
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2933
      Top             =   0
   End
   Begin TabDlg.SSTab tabCDLib 
      Height          =   4245
      Left            =   53
      TabIndex        =   5
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7488
      _Version        =   327680
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "MusicHdr.frx":045E
      Tab(0).ControlCount=   18
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblYearReleased"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblArtist"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCollectionID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRating"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgCover"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "updYearReleased"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtYearReleased"
      Tab(0).Control(8).Enabled=   -1  'True
      Tab(0).Control(9)=   "txtLabel"
      Tab(0).Control(9).Enabled=   -1  'True
      Tab(0).Control(10)=   "txtTitle"
      Tab(0).Control(10).Enabled=   -1  'True
      Tab(0).Control(11)=   "txtArtist"
      Tab(0).Control(11).Enabled=   -1  'True
      Tab(0).Control(12)=   "txtCollectionID"
      Tab(0).Control(12).Enabled=   -1  'True
      Tab(0).Control(13)=   "datMusicHdr"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "datCategory"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dbcboCategory"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame2"
      Tab(0).Control(17).Enabled=   0   'False
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "MusicHdr.frx":047A
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      TabCaption(2)   =   "Categories"
      TabPicture(2)   =   "MusicHdr.frx":0496
      Tab(2).ControlCount=   3
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "datCatChange"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtCategoryID"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtCategoryDesc"
      Tab(2).Control(2).Enabled=   0   'False
      Begin VB.TextBox txtCategoryDesc 
         DataField       =   "CategoryDesc"
         DataSource      =   "datCatChange"
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtCategoryID 
         DataField       =   "CategoryID"
         DataSource      =   "datCatChange"
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Data datCatChange 
         Connect         =   "Access"
         DatabaseName    =   "C:\My Documents\Visual Basic\Class\Ws06\CDLib.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Category"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rating"
         Height          =   1935
         Left            =   -69840
         TabIndex        =   19
         Top             =   2160
         Width           =   1575
         Begin VB.OptionButton optRating 
            Caption         =   "Mind Melting"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton optRating 
            Caption         =   "Awful"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton optRating 
            Caption         =   "Fair"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optRating 
            Caption         =   "Good"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optRating 
            Caption         =   "Excellent"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Media"
         Height          =   1215
         Left            =   -69840
         TabIndex        =   18
         Top             =   960
         Width           =   1575
         Begin VB.CheckBox chkCassette 
            Caption         =   "Cassette"
            DataField       =   "CassetteMedia"
            DataSource      =   "datMusicHdr"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox chkLP 
            Caption         =   "LP"
            DataField       =   "LPMedia"
            DataSource      =   "datMusicHdr"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkCD 
            Caption         =   "CD"
            DataField       =   "CDMedia"
            DataSource      =   "datMusicHdr"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSDBCtls.DBCombo dbcboCategory 
         Bindings        =   "MusicHdr.frx":04B2
         DataField       =   "CategoryID"
         DataSource      =   "datMusicHdr"
         Height          =   315
         Left            =   -70440
         TabIndex        =   17
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   327680
         ListField       =   "CategoryDesc"
         BoundColumn     =   "CategoryID"
         Text            =   ""
      End
      Begin VB.Data datCategory 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\My Documents\Visual Basic\Class\Ws04\CDLib.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -70680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Category"
         Top             =   3600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data datMusicHdr 
         Connect         =   "Access"
         DatabaseName    =   "C:\My Documents\Visual Basic\Class\Ws02\CDLib.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   465
         Left            =   -74880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Music_Hdr"
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtCollectionID 
         DataField       =   "CollectionID"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   11
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtArtist 
         DataField       =   "Artist"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   -73440
         MaxLength       =   40
         TabIndex        =   10
         Top             =   900
         Width           =   2535
      End
      Begin VB.TextBox txtTitle 
         DataField       =   "Title"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   -73440
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1260
         Width           =   3255
      End
      Begin VB.TextBox txtLabel 
         DataField       =   "RecordLabel"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   -73440
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1620
         Width           =   2775
      End
      Begin VB.TextBox txtYearReleased 
         DataField       =   "YearReleased"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   -73250
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2040
         Width           =   660
      End
      Begin ComCtl2.UpDown updYearReleased 
         Height          =   285
         Left            =   -73440
         TabIndex        =   7
         Top             =   2040
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   327680
         Value           =   1900
         Alignment       =   0
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtYearReleased"
         BuddyDispid     =   196630
         OrigLeft        =   2280
         OrigTop         =   1980
         OrigRight       =   2475
         OrigBottom      =   2265
         Max             =   3000
         Min             =   1900
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Image imgCover 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "MediaImage"
         DataSource      =   "datMusicHdr"
         Height          =   2175
         Left            =   -72480
         Stretch         =   -1  'True
         Top             =   1995
         Width           =   2175
      End
      Begin VB.Label lblRating 
         BackColor       =   &H0000FFFF&
         Caption         =   "lblRating/Not Visible"
         DataField       =   "Rating"
         DataSource      =   "datMusicHdr"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblCollectionID 
         AutoSize        =   -1  'True
         Caption         =   "Collection ID:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   16
         Top             =   540
         Width           =   945
      End
      Begin VB.Label lblArtist 
         AutoSize        =   -1  'True
         Caption         =   "Artist/Group:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   15
         Top             =   900
         Width           =   900
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   -73920
         TabIndex        =   14
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Recording Label:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   13
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label lblYearReleased 
         AutoSize        =   -1  'True
         Caption         =   "Year Released:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   12
         Top             =   1980
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5153
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4193
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3233
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   2393
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   1433
      TabIndex        =   0
      ToolTipText     =   "Add a record"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Clock"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Row"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuChange 
         Caption         =   "&Change Row"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Row"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
         Shortcut        =   ^D
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuChooseColors 
         Caption         =   "&Colors"
         Begin VB.Menu mnuColor 
            Caption         =   "&Red"
            Index           =   0
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Green"
            Index           =   1
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Blue"
            Index           =   2
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Default"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu mnuFonts 
         Caption         =   "&Fonts..."
      End
   End
End
Attribute VB_Name = "frmMusicHdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'just added controls for cat. tab.
'have added code for cmdAdd only.
Dim SavePressed As Boolean
Dim CatSavePressed As Boolean
Private Sub cmdAdd_Click()
Select Case tabCDLib.Tab
    Case 0  'general
        txtCollectionID.Locked = False
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
        mnuSave.Enabled = cmdSave.Enabled
        mnuAdd.Enabled = cmdAdd.Enabled 'this is so I can easily (Cut & paste) keep the
        mnuChange.Enabled = cmdChange.Enabled   'menus and command buttons in sync.
        mnuDelete.Enabled = cmdDelete.Enabled
        
        txtCollectionID.SetFocus  'put Cursor in the first field for entering
        'information
    Case 1  'Notes
        'Worker:Sorry, kid. This one's under construction. We're supposed to be done
        'soon... But I think we're waiting for an official to inspect it!
           'Crono:OK. I'll try again later.
           'Nadia:Crono! Maybe we should try to find that official and tell him to come out
           'here so we can cross the bridge!
    Case 2  'Categories
            'Gaurd:You there! Halt! Did you know that this is a dangerous unfinished area?
            'Crono:...
            'Gaurd:Ha! I knew it! You kids get on out of here before you get hurt.
            txtCategoryID.Locked = False
            datCatChange.Recordset.AddNew
            txtCategoryID.SelStart = 0
            txtCategoryID.SelLength = Len(txtCategoryID.Text)
            cmdSave.Enabled = True  'Only allow save after adding
            cmdAdd.Enabled = False 'Can't double add
            cmdChange.Enabled = False 'Can't edit record if you're already
            'adding one!!
            cmdDelete.Enabled = False 'Can't delete an incomplete add
            mnuSave.Enabled = cmdSave.Enabled
            mnuAdd.Enabled = cmdAdd.Enabled 'this is so I can easily (Cut & paste) keep the
            mnuChange.Enabled = cmdChange.Enabled   'menus and command buttons in sync.
            mnuDelete.Enabled = cmdDelete.Enabled
            
            txtCategoryID.SetFocus  'put Cursor in the first field for entering

End Select

End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdAdd.Enabled = True Then
        sbrHint.SimpleText = "Begin adding a new database record."
    End If
End Sub

Public Sub cmdChange_Click()
Select Case tabCDLib.Tab
    Case 0
        datMusicHdr.Recordset.Edit
        cmdSave.Enabled = True  'CAN save after done editing
        cmdChange.Enabled = False 'can't change while already changing
        cmdAdd.Enabled = False 'can't add in the middle of changing
        'is *supposed* to already have been set.(Boy is that verb tense confusing)
        mnuSave.Enabled = cmdSave.Enabled
        mnuAdd.Enabled = cmdAdd.Enabled 'this is so I can easily (Cut & paste) keep the
        mnuChange.Enabled = cmdChange.Enabled   'menus and command buttons in sync.
        mnuDelete.Enabled = cmdDelete.Enabled

        If frmMusicHdr.Visible = True Then 'there is an easier way to do this, I'm sure
        'but it works, so who cares
        '(incidentally, the problem is that this sub is called from two places:frmMusicHdr
        'and frmOpen(in code). If it's called from frmOpen, the program crashes at this
        'method because frmOpen is the active form and is, in fact, modal.
            txtArtist.SetFocus  'This time I'll move down artist because the CollectionID
        End If
        Rem does rem still work??
        Rem Boy!! I can't believe REM still works!!
    Case 1  'Notes
        'Worker:Sorry, kid. This one's under construction. We're supposed to be done
        'soon... But I think we're waiting for an official to inspect it!
           'Crono:OK. I'll try again later.
           'Nadia:Crono! Maybe we should try to find that official and tell him to come out
           'here so we can cross the bridge!
    Case 2  'Categories
            'Gaurd:You there! Halt! Did you know that this is a dangerous unfinished area?
            'Crono:...
            'Gaurd:Ha! I knew it! You kids get on out of here before you get hurt.
        datCatChange.Recordset.Edit
        cmdSave.Enabled = True  'CAN save after done editing
        cmdChange.Enabled = False 'can't change while already changing
        cmdAdd.Enabled = False 'can't add in the middle of changing
        'is *supposed* to already have been set.(Boy is that verb tense confusing)
        mnuSave.Enabled = cmdSave.Enabled
        mnuAdd.Enabled = cmdAdd.Enabled 'this is so I can easily (Cut & paste) keep the
        mnuChange.Enabled = cmdChange.Enabled   'menus and command buttons in sync.
        mnuDelete.Enabled = cmdDelete.Enabled

End Select

End Sub

Private Sub cmdChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdChange.Enabled = True Then
        sbrHint.SimpleText = "Edit the current record. If you do this, then edit the record, then click save, you will not get the annoying error message asking you to save your data."
    End If
End Sub

Private Sub cmdDelete_Click()
Select Case tabCDLib.Tab
    Case 0
        datMusicHdr.Recordset.Delete ' Delete and then close and open the DB again
        datMusicHdr.Refresh
    Case 1  'Notes
        'Worker:Sorry, kid. This one's under construction. We're supposed to be done
        'soon... But I think we're waiting for an official to inspect it!
           'Crono:OK. I'll try again later.
           'Nadia:Crono! Maybe we should try to find that official and tell him to come out
           'here so we can cross the bridge!
    Case 2  'Categories
            'Gaurd:You there! Halt! Did you know that this is a dangerous unfinished area?
            'Crono:...
            'Gaurd:Ha! I knew it! You kids get on out of here before you get hurt.
        datCatChange.Recordset.Delete ' Delete and then close and open the DB again
        datMusicHdr.Refresh

End Select

End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdDelete.Enabled = True Then
        sbrHint.SimpleText = "Delete a record. Be careful-- this cannot be undone!"
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "Don't look to me for help on this!"
End Sub

Public Sub cmdSave_Click()
Select Case tabCDLib.Tab
    Case 0
        If txtCollectionID.Locked = False Then _
            txtCollectionID.Locked = True   'make sure that the user only gets to enter the
            'number for the record ONCE.
        SavePressed = True
        datMusicHdr.Recordset.Update
        cmdSave.Enabled = False 'can't save AGAIN
        cmdDelete.Enabled = True 'but you can delete
        cmdAdd.Enabled = True 'you can also add...
        cmdChange.Enabled = True 'or change the current record
        mnuSave.Enabled = cmdSave.Enabled
        mnuAdd.Enabled = cmdAdd.Enabled 'this is so I can easily (Cut & paste) keep the
        mnuChange.Enabled = cmdChange.Enabled   'menus and command buttons in sync.
        mnuDelete.Enabled = cmdDelete.Enabled

        If frmMusicHdr.Visible = True Then  'hide the form when we load a picture.
        'Then show it again when we're done. But, if we don't check to see if it's
        'invisible when we call cmdSave_Click, we won't know whether or not it's been
        'called from the button actually being clicked, or just called from frmOpen.
        '(We'll get a bug if we try to SetFocus while on frmOpen)
            cmdAdd.SetFocus 'If you don't do this, the focus jumps to Delete...Fun!!
        End If
    Case 1  'Notes
        'Worker:Sorry, kid. This one's under construction. We're supposed to be done
        'soon... But I think we're waiting for an official to inspect it!
           'Crono:OK. I'll try again later.
           'Nadia:Crono! Maybe we should try to find that official and tell him to come out
           'here so we can cross the bridge!
    Case 2  'Categories
            'Gaurd:You there! Halt! Did you know that this is a dangerous unfinished area?
            'Crono:...
            'Gaurd:Ha! I knew it! You kids get on out of here before you get hurt.
        If txtCategoryID.Locked = False Then _
            txtCategoryID.Locked = True   'make sure that the user only gets to enter the
            'number for the record ONCE.
        CatSavePressed = True
        datCatChange.Recordset.Update
        'let's try this to update the DBCombo...
        datMusicHdr.Refresh
        datCategory.Refresh
        cmdSave.Enabled = False 'can't save AGAIN
        cmdDelete.Enabled = True 'but you can delete
        cmdAdd.Enabled = True 'you can also add...
        cmdChange.Enabled = True 'or change the current record
        mnuSave.Enabled = cmdSave.Enabled
        mnuAdd.Enabled = cmdAdd.Enabled 'this is so I can easily (Cut & paste) keep the
        mnuChange.Enabled = cmdChange.Enabled   'menus and command buttons in sync.
        mnuDelete.Enabled = cmdDelete.Enabled

End Select

End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSave.Enabled = True Then
        sbrHint.SimpleText = "Save the current record. This only works if you have clicked on Add or Edit. If you simply started changing fields, you must click the left or right buttons on the database control."
    End If
End Sub
Private Sub datCatChange_Validate(Action As Integer, Save As Integer)
    If CatSavePressed = True Then  'the user has already pressed save.
        Save = CatSavePressed   'it would be easier to read here if Bob had said Save = True
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
    CatSavePressed = False 'always reset save indicator

End Sub

Private Sub datMusicHdr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "Click the far left and right buttons to move to the beginning and end of the database, respectively. Click the inner left and right buttons to move one record forward or backward."
End Sub

Private Sub datMusicHdr_Reposition()
    If IsNull(datMusicHdr.Recordset("Rating")) Then 'set it to the default
        optRating(0).Value = True
    Else
        optRating(datMusicHdr.Recordset("Rating")).Value = True
    End If
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

Private Sub Form_Load()
Dim I As Integer
    For I = 1 To 6
        sbrHint.Panels.Add
    Next
    With sbrHint.Panels
        .Item(1).Style = sbrTime
        .Item(1).Width = 1000
        .Item(2).Style = sbrDate
        .Item(2).Width = 1000
        .Item(3).Style = sbrIns
        .Item(3).Width = 1000
        .Item(4).Style = sbrCaps
        .Item(4).Width = 1000
        .Item(5).Style = sbrScrl
        .Item(5).Width = 1000
        .Item(6).Style = sbrNum
        .Item(6).Width = 1000
    End With
'    cmdChange.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "Ready"
End Sub

Private Sub imgCover_DblClick()
    Load frmOpen
    frmOpen.imgPreview.Picture = Me.imgCover.Picture
    frmOpen.Show vbModal
End Sub

Private Sub lblArtist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The name of the artist or group who performed this piece of 'music'"
End Sub

Private Sub lblCollectionID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The number of the current record inside the database."

End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The company who funded the production of this piece of 'music'."
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The name of this piece of 'music'."

End Sub

Private Sub lblYearReleased_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The year that this piece of 'music' was available to the general public. For some reason, Bob said not to let you enter a year earlier than 1900. Why, I wonder? What if you want to enter some early music recorded on wax cylinders in the 1890's? What if you want to make up FAKE songs?"

End Sub

Private Sub mnuAdd_Click()
    cmdAdd_Click
End Sub

Private Sub mnuChange_Click()
    cmdChange_Click
End Sub

Private Sub mnuColor_Click(Index As Integer)
Static OldIndex As Integer
    mnuColor(OldIndex).Enabled = True
    mnuColor(OldIndex).Checked = False
    mnuColor(Index).Enabled = False
    mnuColor(Index).Checked = True
    OldIndex = Index
    Select Case Index
        Case 0:
            frmMusicHdr.BackColor = vbRed
        Case 1:
            frmMusicHdr.BackColor = vbGreen
        Case 2:
            frmMusicHdr.BackColor = vbBlue
        Case 3:
            frmMusicHdr.BackColor = vbButtonFace    'system color constant
    End Select
End Sub

Private Sub mnuCopy_Click()
Dim Temp As Control
    Set Temp = Screen.ActiveControl 'why do I have to use the set command? I think
    'beacuse Temp is only a POINTER to a control. Therefore all I'm doing is using special
    'syntax to say: give temp the address of screen.activecontrol, not a copy of the actual
    'control...
    If TypeOf Temp Is TextBox Or TypeOf Temp Is DBCombo Then    'the only valid types
        Clipboard.SetText Temp.SelText  'of controls on this form that you can type text on.
        'note that these menus do NOT support Cut & Pasting of images.
    End If
End Sub

Private Sub mnuCut_Click()
Dim Temp As Control
    Set Temp = Screen.ActiveControl 'why do I have to use the set command? I think
    'beacuse Temp is only a POINTER to a control. Therefore all I'm doing is using special
    'syntax to say: give temp the address of screen.activecontrol, not a copy of the actual
    'control...
    If TypeOf Temp Is TextBox Or TypeOf Temp Is DBCombo Then    'the only valid types
        Clipboard.SetText Temp.SelText  'of controls on this form that you can type text on.
        Temp.SelText = ""   'note that these menus do NOT support Cut & Pasting of images.
    End If
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuFonts_Click()
Dim ThisControl As Control
    CommonDialog1.Flags = cdlCFScreenFonts
    CommonDialog1.ShowFont
    If CommonDialog1.FontName = "" Then
        MsgBox "Font name must be entered!", vbOKOnly, "No font name."
        Exit Sub
    ElseIf CommonDialog1.FontSize > 16 Then
        MsgBox "Font size must be 16 or less.", vbOKOnly + vbExclamation, "Font too big"
        Exit Sub
    End If
    For Each ThisControl In Me.Controls
    With ThisControl
        If .Tag <> "SKIP" Then
            If TypeOf ThisControl Is TextBox Then
                .FontName = CommonDialog1.FontName
                .FontSize = CommonDialog1.FontSize
                .FontBold = CommonDialog1.FontBold
                .FontItalic = CommonDialog1.FontItalic
                .Height = 0
            ElseIf TypeOf ThisControl Is Label Or TypeOf ThisControl Is CheckBox Or TypeOf ThisControl Is OptionButton Then
                .FontName = CommonDialog1.FontName
                .FontSize = CommonDialog1.FontSize
                .FontBold = CommonDialog1.FontBold
                .FontItalic = CommonDialog1.FontItalic
            End If
        End If
    End With
    Next ThisControl
End Sub

Private Sub mnuPaste_Click()
Dim Temp As Control
    Set Temp = Screen.ActiveControl 'why do I have to use the set command? I think
    'beacuse Temp is only a POINTER to a control. Therefore all I'm doing is using special
    'syntax to say: give temp the address of screen.activecontrol, not a copy of the actual
    'control...
    If TypeOf Temp Is TextBox Or TypeOf Temp Is DBCombo Then    'the only valid types
            Temp.SelText = Clipboard.GetText(1) 'of controls on this form that you can type
            'text on. Note that these menus do NOT support Cut & Pasting of images.
    End If

End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub optRating_Click(Index As Integer)
    lblRating.Caption = Index
End Sub

Private Sub sbrHint_Click()
    With sbrHint
        If .Style = sbrNormal Then
            .Style = sbrSimple
        Else
            .Style = sbrNormal
        End If
    End With
End Sub

Private Sub Timer1_Timer()
Static Count As Integer
Static GoingUp As Boolean
    'Count = 0
    If Count >= 100 Then
        GoingUp = False
    ElseIf Count <= 0 Then
        GoingUp = True
    End If
    If GoingUp Then
        Count = Count + 20
    Else
        Count = Count - 20
    End If
    lblTime.ForeColor = RGB((Count * 2) + 50, 0, 0)
End Sub

Private Sub Timer2_Timer()
    lblTime.Caption = Format(Now, "hh:mm:ss AM/PM - d mmmm yyyy")
End Sub

Private Sub txtArtist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The name of the artist or group who performed this piece of 'music'"

End Sub

Private Sub txtCollectionID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The number of the current record inside the database."

End Sub

Private Sub txtLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The company who funded the production of this piece of 'music'."
End Sub

Private Sub txtTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbrHint.SimpleText = "The name of this piece of 'music'."

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
    sbrHint.SimpleText = "The year that this piece of 'music' was available to the general public. For some reason, Bob said not to let you enter a year earlier than 1900. Why, I wonder? What if you want to enter some early music recorded on wax cylinders in the 1890's? What if you want to make up FAKE songs?"
End Sub
Private Sub updYearReleased_GotFocus()
    updYearReleased.Value = txtYearReleased.Text
End Sub
