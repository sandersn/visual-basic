VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   600
   End
   Begin MCI.MMControl MMControl3 
      Height          =   1125
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   1984
      _Version        =   327680
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "Sequencer"
      FileName        =   "C:\My Documents\Multimedia\WINDOWS\DESKTOP\FREEDO~1\HeroWolf.mid"
   End
   Begin MCI.MMControl MMControl2 
      Height          =   1125
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   1984
      _Version        =   327680
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "Sequencer"
      FileName        =   "C:\My Documents\Multimedia\WINDOWS\DESKTOP\FREEDO~1\Intro.mid"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   1800
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   3175
      _Version        =   327680
      Orientation     =   1
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "AVIVideo"
      FileName        =   "C:\My Documents\Multimedia\WINDOWS\DESKTOP\FREEDO~1\Windstor.avi"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    End
End Sub

Private Sub Form_Load()
    Timer1.Interval = 6000
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.Command = "Open"
'        MMControl3.Notify = False
'    MMControl3.Wait = True
'    MMControl3.Shareable = False
'    MMControl3.Command = "Open"
        MMControl2.Notify = False
    MMControl2.Wait = True
    MMControl2.Shareable = False
    MMControl2.Command = "Open"


End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
    MMControl2.Command = "Close"
'    MMControl3.Command = "Close"
End Sub

Private Sub Timer1_Timer()
    MMControl1.Command = "Play"
    MMControl2.Command = "Play"
    Timer1.Interval = 0
End Sub
