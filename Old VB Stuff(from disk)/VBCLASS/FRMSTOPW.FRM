VERSION 2.00
Begin Form frmStopWatch 
   Caption         =   "Stop Watch"
   ClientHeight    =   2790
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5595
   Height          =   3195
   Icon            =   FRMSTOPW.FRX:0000
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   5595
   Top             =   1140
   Width           =   5715
   Begin CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   975
   End
   Begin CommandButton cmdStop 
      Caption         =   "S&top Timer"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin CommandButton cmdStart 
      Caption         =   "&Start Timer"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin TextBox txtElapsed 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1455
   End
   Begin TextBox txtStop 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin TextBox txtStart 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   1455
   End
   Begin Label lblElapsed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Elapsed Time:"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin Label lblStop 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Stop Time:"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   930
   End
   Begin Label lblStart 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Start time:"
      Height          =   195
      Left            =   1050
      TabIndex        =   2
      Top             =   480
      Width           =   885
   End
End
Option Explicit

Sub cmdExit_Click ()
    'I decided to make this program be nice to you when you
    'leave. (Except it gives you a mildly reproachful Beep)
    Beep
    MsgBox "Goodbye!"
    End
End Sub

Sub cmdStart_Click ()
    'This part gets the start time into the txtStart text box.
     txtStart.Text = Format(Now, "ttttt")
     'This is the part that clears the bottom two text boxs.
     txtStop.Text = " "
     txtElapsed.Text = " "
End Sub

Sub cmdStop_Click ()
    'Here is where I got the txtStop time.
    txtStop.Text = Format(Now, "ttttt")
    'This is the part where I really has trouble. I tried
    'The Val and Str functions before Dad suggested trying
    'the Format function(which is hinted at in the LR in the
    'Str[$] reference)Part of the solution is in the help for
    'Format.
    txtElapsed.Text = Format(txtStop.Text, "ss") - Format(txtStart.Text, "ss") & " second(s)."
End Sub

