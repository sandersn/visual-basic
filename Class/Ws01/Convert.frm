VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKilos 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtMiles 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Kilometers:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Miles:"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MILESTOKILOS = 1.6

Private Sub cmdConvert_Click()
    txtKilos.Text = txtMiles.Text * MILESTOKILOS
End Sub

Private Sub cmdOK_Click()
'a comment
    frmConvert = Nothing
End Sub
