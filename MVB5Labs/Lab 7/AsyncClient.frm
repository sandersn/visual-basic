VERSION 5.00
Begin VB.Form frmAsyncClient 
   Caption         =   "Asynchronous Client"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStartAsync 
      Caption         =   "&Start Async Method"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmAsyncClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents mAsync As AsyncSrv.Async
Attribute mAsync.VB_VarHelpID = -1

Private Sub cmdStartAsync_Click()
    txtStatus.Text = "Processing..."
    mAsync.AsyncMethod
End Sub

Private Sub Form_Load()
    Set mAsync = New AsyncSrv.Async
End Sub

Private Sub mAsync_Complete(Result As String)
    txtStatus.Text = Result
End Sub
