VERSION 5.00
Begin VB.Form frmArraySpeed 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   8640
   ClientTop       =   3375
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   2475
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txt2D 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txt1D 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmArraySpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Command1_Click()
'this test demos the speed of 2D array versus figuring the position(in 2D) from a 1D array myself.
Dim i As Integer, j As Integer
Dim oneD(10000) As Integer
Dim twoD(100, 100) As Integer
Dim TimeElapsed As Double
    TimeElapsed = Timer
    'now show the values in each one(very fast) in a label, then display the total elapsed time.
    For i = 0 To 99
        For j = 1 To 100
            Label1 = oneD((i * 100) + j)
        Next j
    Next i
    txt1D.Text = (Timer - TimeElapsed)
    TimeElapsed = Timer
    For i = 1 To 100
        For j = 1 To 100
            Label1 = twoD(i, j)
        Next j
    Next i
    txt2D.Text = (Timer - TimeElapsed)
End Sub
