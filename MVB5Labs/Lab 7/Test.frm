VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   1980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Test CC Class"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waiting for Entry..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cc As Lab.CreditCard

Private Sub cc_Status(ByVal strName As String)
    lblStatus.Caption = strName
End Sub

Private Sub Command1_Click()
On Error GoTo PurchaseErr:
    Set cc = New Lab.CreditCard
    cc.ExpireDate = "1/1/12"
    cc.PurchaseAmount = 100
    cc.CardNumber = 1234
    MsgBox cc.Approve
    lblStatus = "Waiting for entry..."
    Exit Sub
PurchaseErr:
    If Err.Number = vbObjectError + 1000 Then MsgBox "Purchase must be larger than zero!" Else Exit Sub
End Sub

