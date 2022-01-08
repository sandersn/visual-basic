VERSION 5.00
Begin VB.Form frmStupid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Stupid Program"
   ClientHeight    =   1380
   ClientLeft      =   3375
   ClientTop       =   3720
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   3008
      TabIndex        =   3
      Top             =   720
      Width           =   1425
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   435
      Left            =   1568
      TabIndex        =   2
      Top             =   720
      Width           =   1425
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   435
      Left            =   128
      TabIndex        =   1
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label lblStupid 
      AutoSize        =   -1  'True
      Caption         =   "Are you stupid??"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmStupid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
    MsgBox "Yes you are!"
End Sub

Private Sub cmdYes_Click()
    MsgBox "No you're not!"
End Sub
