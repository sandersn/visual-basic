VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "Timer Holder"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Callback As Async

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Callback.DoWork
End Sub
