VERSION 5.00
Object = "{7CDAE33A-0321-11D3-ADB9-646109C10000}#1.0#0"; "SmartScrollBar.ocx"
Object = "{7CDAE34A-0321-11D3-ADB9-646109C10000}#1.0#0"; "VSmartScrollBar.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VSmartScrollBar.VSScroll vsbTest 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3413
   End
   Begin VB.CommandButton cmdClickMe 
      Caption         =   "&Click Me"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "YoHoHo"
      Top             =   840
      Width           =   1215
   End
   Begin SmartScrollBar.HSScroll hsbTest 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
   End
   Begin VB.Label lblScroll 
      AutoSize        =   -1  'True
      Caption         =   "I'm a &Label"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label lblChange 
      AutoSize        =   -1  'True
      Caption         =   "&I am Here"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClickMe_Click()
    vsbTest.Value = 13
End Sub

Private Sub hsbTest_Change()
    lblChange = hsbTest.Value
End Sub

Private Sub hsbTest_GotFocus()
    cmdClickMe.SetFocus
End Sub

Private Sub hsbTest_Scroll()
    lblChange = hsbTest.Value
End Sub

Private Sub vsbTest_Change()
    lblChange = vsbTest.Value
End Sub

Private Sub vsbTest_Scroll()
    lblScroll = vsbTest.Value
End Sub
