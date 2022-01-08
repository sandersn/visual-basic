VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Library Browser"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "Browser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9763
      _Version        =   327680
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      MouseIcon       =   "Browser.frx":044A
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Browser.frx":0466
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Browser.frx":0560
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Browser.frx":065A
            Key             =   "Open"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBCDLib As Database
Dim rsMusicHdr As Recordset
Dim NodeX As Node

Private Sub Form_Load()
Dim i As Integer
Dim y As Integer
Dim SavedNodeIndex

    Set DBCDLib = DBEngine.OpenDatabase("C:\My Documents\Visual Basic\Class\Ws07\CDLib.mdb")
    Set NodeX = TreeView1.Nodes.Add()   'create the first node
    NodeX.Text = "Music Library"
    NodeX.Image = "Open"
    NodeX.Expanded = True
    
    For i = 0 To 25
        Set NodeX = TreeView1.Nodes.Add(1, tvwChild)
        NodeX.Text = Chr$(65 + i)
        NodeX.Key = Chr$(65 + i)
        NodeX.Image = "Closed"
        SavedNodeIndex = NodeX.Index
        Set rsMusicHdr = DBCDLib.OpenRecordset _
        ("select * from Music_Hdr Where Artist like '" _
        & Chr$(65 + i) & "*'")
        
        Do Until rsMusicHdr.EOF
            Set NodeX = TreeView1.Nodes.Add(SavedNodeIndex, tvwChild)
            NodeX.Text = rsMusicHdr!Artist
            NodeX.Key = Chr$(65 + 1) & Str$(y) ' unique ID
            NodeX.Tag = rsMusicHdr!CollectionId 'music item ID
            NodeX.Image = "Leaf"    'image from imagelist1
            'move to next record in rstitles.
            rsMusicHdr.MoveNext
            y = y + 1
        Loop
    Next
End Sub


Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
    Node.Image = "Closed"
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
    Node.Image = "Open"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    If Node.Children = 0 Then
        frmMusicHdr!datMusicHdr.RecordSource _
         = "select * from Music_Hdr Where CollectionID = " _
        & Node.Tag
        frmMusicHdr!datMusicHdr.Refresh
        frmMusicHdr.Show vbModal
    End If
End Sub
