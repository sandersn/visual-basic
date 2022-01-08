VERSION 5.00
Begin VB.Form frmStupid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Are you stupid?"
   ClientHeight    =   1380
   ClientLeft      =   5025
   ClientTop       =   4200
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   1980
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmStupid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
    MsgBox ("Yes you are!")
End Sub

Private Sub cmdYes_Click()
    MsgBox ("No you're not!")
End Sub
