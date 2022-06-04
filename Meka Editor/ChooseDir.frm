VERSION 5.00
Begin VB.Form frmChooseDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Meka Directory"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "ChooseDir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose the directory which contains Meka"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmChooseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Me.Tag = Dir1.Path
    Me.Hide
    frmBlitters.Show
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dir1.Path = Dir1.List(Dir1.ListIndex)
    End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    If Command$ <> "" And Me.Tag = "" Then
        'we'll just assume they typed in a pure command withOUT the trailing \
        Me.Tag = Trim$(Command$)
        'snip the trailing \, if present
        If Right$(Me.Tag, 1) = "\" Then Me.Tag = Left$(Me.Tag, Len(Me.Tag) - 1)
        Me.Hide
        frmBlitters.Show
    End If
End Sub
