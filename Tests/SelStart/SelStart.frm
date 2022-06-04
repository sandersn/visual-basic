VERSION 5.00
Begin VB.Form frmSelStart 
   Caption         =   "Set Textbox .SelStart"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "2"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to Set .SelStart"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "SelStart.frx":0000
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmSelStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'comments
'textbox cursor positions start at 0, no bounds checking, it just stops the program if it's below 0,
'or puts it at the end for calues too large.
'there are *2* spaces reserved for each vbCrLf, not 1.
Private Sub Command1_Click()
'awful names, who cares? it's just a quick test.
    Text1.SelStart = Text2.Text
    Text1.SetFocus
End Sub
