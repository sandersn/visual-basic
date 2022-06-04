VERSION 5.00
Begin VB.Form frmAnimate 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStore 
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   120
      Picture         =   "Animate.frx":0000
      ScaleHeight     =   340
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   720
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_Click()

End Sub

Private Sub Form_Click()
Static bGoing As Boolean
Dim intFrame As Integer
Dim sngLoopTime As Single
    bGoing = Not bGoing
    If bGoing Then
        Do
            'show the next frame
            picShow.PaintPicture picStore.Picture, 0, 0, 20, 20, 0, intFrame * 20, 20, 20
            intFrame = intFrame + 1
            If intFrame = 17 Then intFrame = 0
            sngLoopTime = Timer
            Do
                DoEvents
            Loop While sngLoopTime > Timer - 0.1
        Loop While bGoing = True
    End If
End Sub

