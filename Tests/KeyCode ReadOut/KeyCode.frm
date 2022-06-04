VERSION 5.00
Begin VB.Form frmKeyCode 
   Caption         =   "KeyCode Display"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Click text box to reset focus"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Key Code:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmKeyCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub List1_GotFocus()
    Text1.SetFocus
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    List1.AddItem KeyCode
    Text1.Text = "" 'reset
End Sub
