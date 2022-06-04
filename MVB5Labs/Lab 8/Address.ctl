VERSION 5.00
Begin VB.UserControl Address 
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   ScaleHeight     =   1620
   ScaleWidth      =   2430
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company Name:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
End
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Property Get CompanyName()
Attribute CompanyName.VB_MemberFlags = "3c"
    CompanyName = txtCompanyName.Text
End Property
Public Property Let CompanyName(strName As String)
    If CanPropertyChange("CompanyName") Then
        txtCompanyName.Text = strName
    End If
End Property
Public Property Get Address()
Attribute Address.VB_MemberFlags = "1c"
    Address = txtAddress.Text
End Property
Public Property Let Address(ByVal strAddress As String)
    If CanPropertyChange("Address") Then
        txtAddress.Text = strAddress
    End If
End Property
Private Sub txtAddress_Change()
    PropertyChanged "Address"
End Sub
Private Sub txtCompanyName_Change()
    PropertyChanged "CompanyName"
End Sub
