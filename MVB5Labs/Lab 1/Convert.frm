VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtCelsius 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtFahrenheit 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "&Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label lblFahrenheit 
      AutoSize        =   -1  'True
      Caption         =   "Fahrenheit:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
   Begin VB.Label lblCelcius 
      AutoSize        =   -1  'True
      Caption         =   "&Celsius:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub


Private Sub txtCelsius_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    txtFahrenheit.Text = (txtCelsius.Text * 9 / 5) + 32
    Exit Sub
ErrHandle:
    If Err.Number = 12 Then
        txtFahrenheit.Text = "Can't Convert"
    End If
End Sub


Private Sub txtFahrenheit_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    txtCelsius.Text = (txtFahrenheit.Text - 32) * (5 / 9)
    Exit Sub
ErrHandle:
    If Err.Number = 12 Then txtCelsius.Text = "Can't Convert"
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    txtUserName.Text = UCase(txtUserName.Text)
    txtUserName.SelStart = Len(txtUserName.Text)
End Sub
