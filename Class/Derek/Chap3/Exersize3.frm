VERSION 5.00
Begin VB.Form frmMusicHdr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Selection"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "Exersize3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
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
      DatabaseName    =   "C:\My Documents\Visual Basic\Class\Derek\Chap3\CDLibe.mdb"
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
      TabIndex        =   15
      Top             =   4920
      Width           =   7935
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
Option Explicit
Dim SavePressed As Boolean

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

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint.Caption = "This adds to the list of listings"
End Sub


Private Sub cmdChange_Click()
datMusicHdr.Recordset.Edit
cmdSave.Enabled = True
cmdAdd.Enabled = False
cmdChange.Enabled = False
cmdDelete.Enabled = False
End Sub

Private Sub cmdChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint.Caption = "This allows for editing to your list presant"
End Sub


Private Sub cmdDelete_Click()
datMusicHdr.Recordset.Delete
datMusicHdr.Refresh
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "This deletes current record"
End Sub


Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "You need help if you do not understand what this botton does"
End Sub


Private Sub cmdSave_Click()
SavePressed = True
datMusicHdr.Recordset.Update
cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdChange.Enabled = True
cmdDelete.Enabled = True
End Sub


Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "This saves changes made"
End Sub


Private Sub datMusicHdr_Validate(Action As Integer, Save As Integer)
If SavePressed = True Then
    Save = SavePressed
Else
  If Save = True Then
    Dim Ans As Integer
    Ans = MsgBox("Data changed. Want to save?", vbYesNo + vbExclamation, "Data not saved")
    If Ans = vbNo Then
        Save = False
    End If
  End If
End If
SavePressed = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint = "Ready"
End Sub


Private Sub lblHint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHint.Caption = "This is the hint box"
End Sub


Private Sub txtYearReleased_LostFocus()
If txtYearReleased.Text < 2000 Then
    MsgBox ("Invalid Year")
   txtYearReleased.SelStart = 0
   txtYearReleased.SelLength = 4
    txtYearReleased.SetFocus
End If
End Sub


