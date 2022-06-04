VERSION 5.00
Begin VB.Form frmTestPoint 
   Caption         =   "Test Point Function"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picView 
      AutoSize        =   -1  'True
      Height          =   3405
      Left            =   240
      Picture         =   "TestPoint.frx":0000
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   5
      Top             =   1200
      Width           =   3900
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Text            =   "Y"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "X"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdBF 
      Caption         =   "BF PicBox"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Color by .Point"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblReadout 
      Caption         =   "Ready"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTestPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBF_Click()
Static intLastColor As Integer
    intLastColor = intLastColor + 1
    If intLastColor > 15 Then intLastColor = 0
    picView.Line (0, 0)-(picView.Width, picView.Height), QBColor(intLastColor), BF
End Sub

Private Sub cmdFind_Click()
    lblReadout.Caption = picView.Point(txtX.Text, txtY.Text)
End Sub

Private Sub picView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblReadout.Caption = picView.Point(X, Y)
End Sub
