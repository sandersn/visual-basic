VERSION 5.00
Begin VB.Form frmPoints 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graph Points and Lines and stuff and things®"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAnisotropic 
      Caption         =   "Draw A&nisotropic"
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "&Animate Graph"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Graph"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtPoints 
      Height          =   285
      Left            =   7680
      TabIndex        =   1
      Text            =   "6"
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      ScaleHeight     =   200
      ScaleLeft       =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   -100
      ScaleWidth      =   200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   7680
      TabIndex        =   5
      Top             =   1800
      Width           =   1850
      Begin VB.CheckBox chkAnimateBothAxes 
         Caption         =   "Animate second a&xis"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtPointsY 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "5"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblSecondAxis 
         AutoSize        =   -1  'True
         Caption         =   "Points on &second axis:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Label lblAd 
      Caption         =   "Your Ad Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3375
      Left            =   7560
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblNumPoints 
      AutoSize        =   -1  'True
      Caption         =   "Number of Points:"
      Height          =   195
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bAnisotropic As Boolean
Private bAnimateBothAxes As Boolean
Private Sub DrawGraph(intPoints As Integer, Optional intPointsY As Integer)
Dim i As Integer, j As Integer
'real function ^_^
Picture1.Cls
If intPointsY = 0 Then intPointsY = intPoints
For i = 1 To intPoints Step 1
    For j = 1 To intPointsY Step 1
        'inside section
        Picture1.Line ((100 / intPoints) * i, 0)-(0, (100 / intPointsY) * j)
        Picture1.Line (-(100 / intPoints) * i, 0)-(0, -(100 / intPointsY) * j)
        Picture1.Line (-(100 / intPoints) * i, 0)-(0, (100 / intPointsY) * j)
        Picture1.Line ((100 / intPoints) * i, 0)-(0, -(100 / intPointsY) * j)
        'outside section
        Picture1.Line (100 - ((100 / intPoints) * i), 100)-(100, 100 - ((100 / intPointsY) * j))
        Picture1.Line (100 - ((100 / intPoints) * i), -100)-(100, -100 + ((100 / intPointsY) * j))
        Picture1.Line (-100 + ((100 / intPoints) * i), -100)-(-100, -100 + ((100 / intPointsY) * j))
        Picture1.Line (-100 + ((100 / intPoints) * i), 100)-(-100, 100 - ((100 / intPointsY) * j))
    Next j
Next i
Picture1.Line (-100, 0)-(100, 0)
Picture1.Line (0, -100)-(0, 100)

End Sub
Private Sub SLEP(sngSec As Single)
Dim sngStart As Single
    sngStart = Timer
    Do
    Loop While Timer - sngStart < sngSec
End Sub

Private Sub SLP2(sngMSec As Single)
Dim sngStart As Single
    sngStart = Timer
    Do
    Loop While (Timer - sngStart) * 1000 < sngMSec
End Sub


Private Sub chkAnimateBothAxes_Click()
    bAnimateBothAxes = Not bAnimateBothAxes
End Sub

Private Sub chkAnimateBothAxes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAd.Visible = True
End Sub

Private Sub chkAnisotropic_Click()
    bAnisotropic = Not bAnisotropic
    Frame1.Enabled = bAnisotropic
    txtPointsY.Enabled = bAnisotropic
    chkAnimateBothAxes.Enabled = bAnisotropic
    lblSecondAxis.Enabled = bAnisotropic
End Sub

Private Sub cmdAnimate_Click()
Dim i As Integer, j As Integer
    If IsNumeric(txtPoints.Text) Then
        If bAnisotropic Then
            If IsNumeric(txtPointsY.Text) Then
                If bAnimateBothAxes Then
                    For i = 1 To txtPoints.Text
                        Picture1.PSet (-70, 0)
                        Picture1.Print "Take a Shower"
                        For j = 1 To txtPointsY.Text
                            DrawGraph i, j
                            SLEP (0.1)
                        Next j
                    Next i
                Else
                    For i = 1 To txtPoints.Text
                        Picture1.PSet (-70, 0)
                        Picture1.Print "Take a Shower"
                        DrawGraph i, txtPointsY.Text
                        SLEP (0.1)
                    Next i
                End If
            End If
        Else
            For i = 1 To txtPoints.Text
                Picture1.PSet (-70, 0)
                Picture1.Print "Take a Shower"
                DrawGraph i
                SLEP (0.1)
            Next i
        End If
    End If
    txtPoints.SetFocus
End Sub

Private Sub cmdDraw_Click()
If IsNumeric(txtPoints.Text) Then
    If bAnisotropic Then
        If IsNumeric(txtPointsY.Text) Then
            'subliminal message
            Picture1.PSet (-70, 0)
            Picture1.Print "Take a Shower"
            SLP2 (5) 'so you can see it just barely
            DrawGraph txtPoints.Text, txtPointsY.Text
        End If
    Else
        'subliminal message
        Picture1.PSet (-70, 0)
        Picture1.Print "Take a Shower"
        SLP2 (5) 'so you can see it just barely
        DrawGraph txtPoints.Text
    End If
End If
txtPoints.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAd.Visible = False
End Sub

Private Sub lblAd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAd.Visible = False
End Sub

Private Sub Picture1_Click()
    cmdDraw_Click
End Sub

Private Sub txtPoints_GotFocus()
    txtPoints.SelStart = 0
    txtPoints.SelLength = Len(txtPoints.Text)
End Sub

Private Sub txtPointsY_GotFocus()
    txtPointsY.SelStart = 0
    txtPointsY.SelLength = Len(txtPointsY.Text)
End Sub
