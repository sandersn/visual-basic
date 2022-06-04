VERSION 5.00
Begin VB.Form frmGetWindowsDir 
   Caption         =   "Get Windows Directory"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdGetWindowsDir 
      Caption         =   "&Windows Directory"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmGetWindowsDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGetWindowsDir_Click()
Dim strWinPath As String
    strWinPath = GetWinDirectory
    MsgBox strWinPath
End Sub
