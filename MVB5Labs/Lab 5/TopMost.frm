VERSION 5.00
Begin VB.Form frmTopMost 
   Caption         =   "TopMost Window"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   1095
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Window's Position"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optNotTopMost 
         Caption         =   "Non-Topmost window"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optTopMost 
         Caption         =   "Topmost window"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTopMost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2



Private Sub cmdExit_Click()
    End
End Sub

Private Sub optNotTopMost_Click()
Dim lResult As Long
Dim flags As Long
    flags = SWP_NOSIZE
    lResult = SetWindowPos(frmTopMost.hwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, flags)
End Sub

Private Sub optTopMost_Click()
Dim lResult As Long
Dim flags As Long
    flags = SWP_NOSIZE
    lResult = SetWindowPos(frmTopMost.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)

End Sub
