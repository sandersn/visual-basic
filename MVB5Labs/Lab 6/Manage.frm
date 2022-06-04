VERSION 5.00
Begin VB.Form frmManage 
   Caption         =   "Manage Credit Limit"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewLimit 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "New Limit:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Current Limit:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   915
   End
   Begin VB.Label lblCurLimit 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mng As IManage

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsNumeric(txtNewLimit) Then
        mng.CreditLimit = txtNewLimit.Text
        Unload Me
        Exit Sub
    Else
        txtNewLimit.SetFocus
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set mng = frmClient.cc
    If mng Is Nothing Then
        MsgBox "The IManage interface is not supported on this object."
        Unload Me
        Exit Sub
    End If
    lblCurLimit = mng.CreditLimit
    txtNewLimit = ""
End Sub


Private Sub txtNewLimit_GotFocus()
    txtNewLimit.SelStart = 0
    txtNewLimit.SelLength = Len(txtNewLimit)
End Sub
