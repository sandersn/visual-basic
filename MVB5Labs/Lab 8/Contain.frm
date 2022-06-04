VERSION 5.00
Object = "*\AStock.vbp"
Begin VB.Form frmContain 
   Caption         =   "Stock Ticker Container"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&Activate"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin Stock.StockControl StockControl1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmContain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActivate_Click()
    'toggle active property
    StockControl1.Active = Not StockControl1.Active
End Sub

Private Sub cmdRefresh_Click()
    'force a refresh of the stock price
    StockControl1.Refresh
End Sub

Private Sub StockControl1_TickerKeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
