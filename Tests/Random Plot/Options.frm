VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Plot Options"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHiContrast 
      Caption         =   "Hi-Contrast"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtDistance 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtPoints 
      Height          =   285
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame fraType 
      Caption         =   "&Type"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optType 
         Caption         =   "Minimum &Space"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Random"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Minimum Distance:"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Number of points:"
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intType As Integer

Private Sub cmdOK_Click()
'check to make sure that they gave us a good number.
    If Not IsNumeric(txtPoints.Text) Then
        txtPoints.Text = ""
        txtPoints.SetFocus
        Exit Sub
    End If
    If intType = 1 And Not IsNumeric(txtDistance.Text) Then
        txtPoints.Text = ""
        txtPoints.SetFocus
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub optType_Click(Index As Integer)
    intType = Index
End Sub
Public Sub ReturnOptions(ByRef lngPoints As Long, ByRef intPlotStyle As Integer, ByRef lngDistance, ByRef intHighContrast As Integer)
    lngPoints = txtPoints.Text
    intPlotStyle = intType
    If intType = 1 Then
        lngDistance = txtDistance.Text
    Else
        lngDistance = -1
    End If
    intHighContrast = chkHiContrast.Value
End Sub
