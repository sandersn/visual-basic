VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMmxFxDemo 
   Caption         =   "Eagle Engine Simulator"
   ClientHeight    =   6375
   ClientLeft      =   2055
   ClientTop       =   870
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   Begin VB.Frame fraOrientation 
      Caption         =   "&Orientation"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
      Begin VB.OptionButton optOrientation 
         Caption         =   "Left"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Right"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Bottom"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Top"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Pictures (*.bmp)|*.bmp|Icons (*.ico)|*.ico"
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load New"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "MmxFxDemo.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox picBackBuffer 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   6000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmMmxFxDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Enum Direction
'    Top = 0
'    Bottom = 1
'    Rght = 2
'    Lft = 3
'End Enum
'Dim Orientation As Direction

Private Sub cmdLoad_Click()
    On Error GoTo Hoho   'quit if they didn't select a filename
    CommonDialog1.ShowOpen

    picImage.Picture = LoadPicture(CommonDialog1.filename)
    picImage.AutoSize = True
    picBackBuffer.Width = picImage.Width * 2
    picBackBuffer.Height = picImage.Height * 2
Hoho:
    Exit Sub
End Sub

Private Sub cmdStart_Click()
Dim X As Long, Y As Long, i As Integer
Dim lngColor(1 To 4) As Long
Dim possColor(1 To 4) As Long
Dim tmpColor As Long
'let's go in an timerless loop
'now that we're doing it more complicatedly, we have to check to see which direction to come from
'Select Case Orientation
'    Case Direction.Top
With picImage
    cmdStart.Caption = "Wait--Working"
    cmdStart.Enabled = False
    cmdLoad.Enabled = False
    For Y = 0 To .Height Step 1
        For X = 0 To .Width Step 1
            For i = 1 To 4 Step 1
                lngColor(1) = .Point(X, Y) 'get the color of our point from the Image
                lngColor(2) = .Point(X - 1, Y)
                lngColor(3) = .Point(X, Y - 1)
                lngColor(4) = .Point(X - 1, Y - 1)
                possColor(1) = lngColor(1)
                If lngColor(2) <> possColor(1) Then
                    possColor(2) = lngColor(2)
                'else it's the same, so leave alone at the top
                'now we draw a line on the buffer from our current positon to the bottom.
                picBackBuffer.PSet (X, Y), mostColor
                DoEvents    'make sure we don't hang
            Next i
        Next X
        'when complete we'll BitBlt the contents of picBackBuffer to picView--but not yet. I wnat to see how ugly it looks!
        'Later: it's not ugly at all--we may be able to forgo the BitBlt completely!
    Next Y
    cmdStart.Caption = "Start"
    cmdStart.Enabled = True
    cmdLoad.Enabled = True
End With
'    Case Direction.Bottom
'With picImage
'
'    cmdStart.Caption = "Wait--Working"
'    cmdStart.Enabled = False
'    cmdLoad.Enabled = False
'    For Y = .Height To 0 Step -1
'        For X = 0 To .Width Step 1
'            lngColor = .Point(X, Y)
'            'now we draw a line on the buffer from our current positon to the top.
'            picBackBuffer.Line (X, Y)-(X, 0), lngColor
'            DoEvents
'        Next X
'    Next Y
'    cmdStart.Caption = "Start"
'    cmdStart.Enabled = True
'    cmdLoad.Enabled = True
'End With
'    Case Direction.Lft
'With picImage
'    cmdStart.Caption = "Wait--Working"
'    cmdStart.Enabled = False
'    cmdLoad.Enabled = False
'    For X = 0 To .Width Step 1
'        For Y = 0 To .Height Step 1
'            lngColor = .Point(X, Y)
'            'now we draw a line on the buffer from our current positon to the top.
'            picBackBuffer.Line (X, Y)-(.Width, Y), lngColor
'            DoEvents
'        Next Y
'    Next X
'    cmdStart.Caption = "Start"
'    cmdStart.Enabled = True
'    cmdLoad.Enabled = True
'End With
'    Case Direction.Rght
'With picImage
'
'    cmdStart.Caption = "Wait--Working"
'    cmdStart.Enabled = False
'    cmdLoad.Enabled = False
'    For X = .Width To 0 Step -1
'        For Y = 0 To .Height Step 1
'            lngColor = .Point(X, Y)
'            'now we draw a line on the buffer from our current positon to the top.
'            picBackBuffer.Line (X, Y)-(0, Y), lngColor
'            DoEvents
'        Next Y
'    Next X
'    cmdStart.Caption = "Start"
'    cmdStart.Enabled = True
'    cmdLoad.Enabled = True
'End With
'End Select
End Sub

Private Sub optOrientation_Click(Index As Integer)
    Orientation = Index
End Sub
