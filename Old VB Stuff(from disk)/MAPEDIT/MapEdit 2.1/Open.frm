VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome to the Map Editor!"
   ClientHeight    =   2445
   ClientLeft      =   2355
   ClientTop       =   2010
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Existing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtMapYSize 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox txtMapXSize 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   960
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H008080FF&
      Caption         =   $"Open.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label lblMapYSize 
      AutoSize        =   -1  'True
      Caption         =   "&Height(Map Y):"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label lblMapXSize 
      AutoSize        =   -1  'True
      Caption         =   "&Width(Map X):"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdOK_Click()
    If txtMapXSize.Text = "" Or txtMapYSize.Text = "" Then
        Exit Sub
    End If
    MapXSize = txtMapXSize.Text
    MapYSize = txtMapYSize.Text
    If txtMapXSize.Text < 30 Then MapXSize = 30
    If txtMapYSize.Text < 30 Then MapYSize = 30
    frmOpen.Tag = "FirstTime"
    Load frmMapEdit
    frmMapEdit.mnuNew_Click
    'frmMapEdit.Show
    Unload frmOpen
End Sub

Private Sub cmdOpen_Click()
    Load frmMapEdit 'load but do not show(so that PaintMap is not called yet.
    'frmMapEdit is actually shown within New or Open_Click
    frmOpen.Tag = "FirstTime"   'set the flag so that New and Open know that we're
    'calling from frmOpen
    frmMapEdit.mnuOpen_Click
    If frmOpen.Tag = "" Then    'this for timing: so that if the user didn't press
    'Cancel from within mnuOpen, we'll Unload frmOpen. Otherwise we'll show frmOpen
    'again from within mnuOpen when they press Cancel. We also have to reset the
    'calling from frmOpen indicator.
        Unload frmOpen
        Exit Sub
    End If
    frmOpen.Tag = "FirstTime"
End Sub

Private Sub Form_Load()
    frmMapEdit.CMDialog1.CancelError = True 'make sure that
    'I can tell EVERY time the user presses Cancel
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        End
    End If
End Sub

