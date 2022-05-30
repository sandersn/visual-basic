VERSION 5.00
Begin VB.Form frmHyper 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "The Hyperator"
   ClientHeight    =   4020
   ClientLeft      =   2430
   ClientTop       =   1485
   ClientWidth     =   2430
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "HYPER.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   2430
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompute 
      Appearance      =   0  'Flat
      Caption         =   "&Compute"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtOunces 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.VScrollBar vsbOunces 
      Height          =   3735
      LargeChange     =   5
      Left            =   120
      Max             =   250
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblOunces 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ounces of soda drunk per day:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmHyper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Mr. Spellmann and\or anybody else reading this!
'This is not my final version of this program, but if I
'don't finish my other(newer) version of this program, this
'will do in a pinch. Originally this program was a way to
'test if I could use input from a text box and a scroll
'bar at the same time. It worked so well I am using this
'as back up extra credit.

Private Sub cmdCompute_Click()
Dim Msg As String
    'Make a message telling the amount of soda drunk.
    Msg = "You drank " & txtOunces.Text & " ounces of soda."
    'Print the message.
    MsgBox Msg
    'Make sure the next time you use the scroll bar it
    'isn't in a totally different place from where you
    'used the text box to specify an amount.
    vsbOunces.Value = txtOunces.Text
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub txtOunces_Change()
    If txtOunces.Text = "" Then
        txtOunces.Text = "0"
        txtOunces.SelStart = 0
        txtOunces.SelLength = 3
    End If
    If Val(txtOunces.Text) > 250 Then
        txtOunces.Text = 250
        txtOunces.SelStart = 0
        txtOunces.SelLength = 3
    End If
    vsbOunces.Value = txtOunces.Text
End Sub

Private Sub txtOunces_GotFocus()
    txtOunces.SelStart = 0
    txtOunces.SelLength = 3
End Sub

Private Sub vsbOunces_Change()
     'Make sure the text box shows the value of the scroll
     'bar.
     txtOunces.Text = vsbOunces.Value
End Sub

Private Sub vsbOunces_Scroll()
    'Call the vsbOunces_Change procedure.
    vsbOunces_Change
End Sub

