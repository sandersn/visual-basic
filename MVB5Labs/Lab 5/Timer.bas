Attribute VB_Name = "Timer"
Option Explicit

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


Public Sub TimerProc(ByVal hwnd As Long, ByVal msg As Long, ByVal idEvent As Long, ByVal curTime As Long)
    UpdateProgressBar
End Sub
