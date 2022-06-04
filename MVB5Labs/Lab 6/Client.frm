VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdManage 
      Caption         =   "&Manage"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "&Approve"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processing..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents cc As Lab.CreditCard
Attribute cc.VB_VarHelpID = -1

Private Sub cmdApprove_Click()
    cc.CardNumber = 1234
    cc.ExpireDate = "1/1/01"
    cc.PurchaseAmount = 500
    MsgBox cc.Approve
End Sub

Private Sub cmdManage_Click()
    frmManage.Show vbModal
End Sub

Private Sub Form_Load()
    Set cc = New Lab.CreditCard
End Sub
