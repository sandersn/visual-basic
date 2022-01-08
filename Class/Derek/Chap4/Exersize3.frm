VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmMusicHdr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Selection"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "Exersize3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabCDLib 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Exersize3.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblYearReleased"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblArtist"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CollectionID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRating"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgCover"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtYearReleased"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtLabel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTitle"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtArtist"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCollectionID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "datMusicHdr"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "datCategory"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Timer1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "UpDown1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dbcdoCategory"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "Exersize3.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Categories"
      TabPicture(2)   =   "Exersize3.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "datCatChange"
      Tab(2).ControlCount=   1
      Begin MSDBCtls.DBCombo dbcdoCategory 
         Height          =   315
         Left            =   6360
         TabIndex        =   28
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   327681
         Text            =   "DBCombo1"
      End
      Begin VB.Data datCatChange 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\My Documents\Visual Basic\Class\Derek\Chap4\CDLibe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -72720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Category"
         Top             =   3120
         Width           =   2175
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1964
         TabIndex        =   27
         Top             =   2400
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1900
         Alignment       =   0
         BuddyControl    =   "txtYearReleased"
         BuddyDispid     =   196623
         OrigLeft        =   1920
         OrigTop         =   2400
         OrigRight       =   2160
         OrigBottom      =   2655
         Max             =   3000
         Min             =   1900
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   8400
         Top             =   1080
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rating"
         Height          =   1455
         Left            =   3600
         TabIndex        =   18
         Top             =   3240
         Width           =   2415
         Begin VB.OptionButton optRating 
            Caption         =   "Fair"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optRating 
            Caption         =   "Good"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optRating 
            Caption         =   "Excellent"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Media"
         Height          =   1455
         Left            =   360
         TabIndex        =   17
         Top             =   3240
         Width           =   2415
         Begin VB.CheckBox chkCassette 
            Caption         =   "Cassette"
            DataField       =   "CassetteMedia"
            DataSource      =   "datMusicHdr"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkLP 
            Caption         =   "LP"
            DataField       =   "LPMedia"
            DataSource      =   "datMusicHdr"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkCD 
            Caption         =   "CD"
            DataField       =   "CDMedia"
            DataSource      =   "datMusicHdr"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Data datCategory 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\My Documents\Visual Basic\Class\Derek\Chap4\CDLibe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Category"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data datMusicHdr 
         Connect         =   "Access"
         DatabaseName    =   "C:\My Documents\Visual Basic\Class\Derek\Chap4\CDLibe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Music_Hdr"
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox txtCollectionID 
         DataField       =   "CollectionID"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtArtist 
         DataField       =   "Artist"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   10
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtTitle 
         DataField       =   "Title"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtLabel 
         DataField       =   "RecordLabel"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtYearReleased 
         DataField       =   "YearReleased"
         DataSource      =   "datMusicHdr"
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   7
         Top             =   2400
         Width           =   855
      End
      Begin VB.Image imgCover 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "MediaImage"
         DataSource      =   "datMusicHdr"
         Height          =   1335
         Left            =   6360
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblRating 
         BackColor       =   &H0000FFFF&
         Caption         =   "lblRating/not visible"
         DataField       =   "Rating"
         DataSource      =   "datMusicHdr"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label CollectionID 
         AutoSize        =   -1  'True
         Caption         =   "Colection ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblArtist 
         AutoSize        =   -1  'True
         Caption         =   "Artist / Group:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Volume Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Recording Label:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblYearReleased 
         AutoSize        =   -1  'True
         Caption         =   "Year Released:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Clock"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6480
      TabIndex        =   26
      Top             =   0
      Width           =   885
   End
   Begin VB.Label lblHint 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   7935
   End
End
Attribute VB_Name = "frmMusicHdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SavePressed As Boolean
Dim somevar As Variant
Private Sub cmdAdd_Click()
Select Case tabCDLib.Tab
      Case 0
        datMusicHdr.Recordset.AddNew
        txtYearReleased.Text = 1996
        txtYearReleased.SelStart = 0
        txtYearReleased.SelLength = Len(txtYearReleased.Text)
        cmdSave.Enabled = True
        cmdAdd.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
        txtArtist.SetFocus
        
    Case 1
    Case 2
End Select
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint.Caption = "This adds to the list of listings"
End Sub


Private Sub cmdChange_Click()
Select Case tabCDLib.Tab
    Case 0
        datMusicHdr.Recordset.Edit
        cmdSave.Enabled = True
        cmdAdd.Enabled = False
        cmdChange.Enabled = False
        cmdDelete.Enabled = False
    Case 1
    Case 2
End Select
End Sub

Private Sub cmdChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint.Caption = "This allows for editing to your list presant"
End Sub


Private Sub cmdDelete_Click()
Select Case tabCDLib.Tab
    Case 0
        datMusicHdr.Recordset.Delete
        datMusicHdr.Refresh
    Case 1
    Case 2
End Select
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "This deletes current record"
End Sub


Private Sub cmdExit_Click()
Select Case tabCDLib.Tab
    Case 0
        End
    Case 1
        End
    Case 2
        End
End Select
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "You need help if you do not understand what this botton does"
End Sub


Private Sub cmdSave_Click()
Select Case tabCDLib.Tab
    Case 0
        datMusicHdr.Recordset.Update
        cmdSave.Enabled = False
        cmdAdd.Enabled = True
        cmdChange.Enabled = True
        cmdDelete.Enabled = True
        SavePressed = True
    Case 1
    Case 2
End Select
End Sub


Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "This saves changes made"
End Sub


Private Sub datMusicHdr_Reposition()
If IsNull(datMusicHdr.Recordset("Rating")) Then
        optRating(0).Value = True

Else
        optRating(datMusicHdr.Recordset("Rating")).Value = True
End If

End Sub

Private Sub datMusicHdr_Validate(Action As Integer, Save As Integer)
If SavePressed = True Then
    Save = SavePressed
Else
  If Save = True Then
    Dim Ans As Integer
    Ans = MsgBox("Data changed. Want to save?", _
                    vbYesNo + vbExclamation, "Data not saved")
    If Ans = vbNo Then
                        Save = False
    End If
  End If
End If
SavePressed = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "This is the Form"
End Sub


Private Sub lblHint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint.Caption = "This is the hint box"
End Sub


Private Sub optRating_Click(Index As Integer)
         lblRating.Caption = Index
End Sub

Private Sub Timer1_Timer()
somevar = Now

'somevar = Format$(Now, "hh:mm:ss AMPM - mmm, d yyyy")
lblTime.Caption = somevar
End Sub

Private Sub txtYearReleased_LostFocus()
If txtYearReleased.Text < 1900 Then
    MsgBox ("Invalid Year")
   ' txtYearReleased.SelText
   ' txtYearReleased.SelLength
    txtYearReleased.SetFocus
End If
End Sub
