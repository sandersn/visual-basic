VERSION 5.00
Begin VB.Form frmTaskBar 
   Caption         =   "TaskBar Listing"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4


Private Sub Form_Load()
Dim szMy As String * 60
Dim x As Integer
Dim CurrWnd As Long
    CurrWnd = GetWindow(frmTaskBar.hwnd, GW_HWNDFIRST)
    Do Until CurrWnd = 0
        If IsWindowVisible(CurrWnd) <> False And _
          GetWindow(CurrWnd, GW_OWNER) = 0 Then
            x = GetWindowText(CurrWnd, szMy, 30)
            List1.AddItem szMy
        End If
        CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
        x = DoEvents
    Loop
End Sub
