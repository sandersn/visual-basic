VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Form2"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form2"
   ScaleHeight     =   1875
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image imgAbout 
      Height          =   330
      Left            =   240
      Picture         =   "About.frx":0000
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblAbout 
      Caption         =   "Conversion Application"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub
