VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmTimer 
   Caption         =   "Long Job"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   450
      _Version        =   327680
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "Timer.frx":0000
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Start"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTimerID As Long

Private Sub cmdStop_Click()
    If lngTimerID = 0 Then
        StartTimer
    Else
        EndTimer
    End If
End Sub
Public Sub UpdateProgressBar()
Dim IncValue As Integer
    IncValue = ProgressBar1.Value + 3
    If IncValue >= 100 Then
        ProgressBar1.Value = 100
        EndTimer
    Else
        ProgressBar1.Value = IncValue
    End If
    
End Sub
Private Sub StartTimer()
    lngTimerID = SetTimer(0, 0, 200, AddressOf TimerProc)
    ProgressBar1.Value = 0
    cmdStop.Caption = "&Stop"
End Sub
Private Sub EndTimer()
    KillTimer 0, lngTimerID
'    If lngTimerID <> 0 Then
'        lngTimerID = KillTimer(frmTimer.hwnd, 0)
'    End If
    lngTimerID = 0
    cmdStop.Caption = "&Start"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndTimer
End Sub
