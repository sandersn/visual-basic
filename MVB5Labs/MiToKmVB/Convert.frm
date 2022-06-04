VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Convert Miles to Kilometers"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKilos 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtMiles 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   360
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label lblKilos 
      AutoSize        =   -1  'True
      Caption         =   "Kilometers:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvert_Click()
    txtKilos.Text = txtMiles.Text * 1.6
End Sub
