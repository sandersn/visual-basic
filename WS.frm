VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Westfield Chat Server"
   ClientHeight    =   1710
   ClientLeft      =   1935
   ClientTop       =   1755
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock WChat 
      Left            =   3720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.TextBox txtText 
      Height          =   990
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   720
      Width           =   4710
   End
   Begin VB.TextBox txtport 
      Height          =   285
      Left            =   1380
      TabIndex        =   3
      Text            =   "6667"
      Top             =   390
      Width           =   540
   End
   Begin VB.TextBox txtadd 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   75
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Server Address:"
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1185
   End
End
Attribute VB_Name = "frmWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
txtadd.Text = WChat.LocalIP
WChat.LocalPort = "6667"
WChat.Listen
'FrmChatB.Visible = True
End Sub

Private Sub txtText_Change()
WChat.SendData txtText.Text
End Sub

Private Sub wchat_ConnectionRequest(ByVal requestID As Long)
If WChat.State <> sckClosed Then WChat.Close
    WChat.Accept requestID
End Sub

Private Sub WChat_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    WChat.GetData strData, vbString
    txtText.Text = strData 'txtText.Text +

End Sub
