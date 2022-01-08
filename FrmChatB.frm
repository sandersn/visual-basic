VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FrmChatB 
   Caption         =   "Westfield Chat"
   ClientHeight    =   2670
   ClientLeft      =   1290
   ClientTop       =   2460
   ClientWidth     =   7320
   Icon            =   "FrmChatB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7320
   Begin RichTextLib.RichTextBox txtrec 
      Height          =   1920
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   3387
      _Version        =   327680
      TextRTF         =   $"FrmChatB.frx":0442
   End
   Begin VB.TextBox txtha 
      Height          =   285
      Left            =   5160
      TabIndex        =   8
      Top             =   270
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4980
      Top             =   3060
   End
   Begin VB.CommandButton Cmddc 
      Caption         =   "Dis-Connect"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3840
      TabIndex        =   7
      Top             =   2250
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock wchat 
      Left            =   8595
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtnick 
      Height          =   285
      Left            =   5070
      TabIndex        =   6
      Top             =   1155
      Width           =   2190
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5100
      TabIndex        =   3
      Top             =   1470
      Width           =   1020
   End
   Begin VB.TextBox txtHP 
      Height          =   285
      Left            =   6675
      TabIndex        =   1
      Text            =   "6667"
      Top             =   555
      Width           =   540
   End
   Begin VB.TextBox txtCB 
      Enabled         =   0   'False
      Height          =   300
      Left            =   45
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1935
      Width           =   4920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nickname"
      Height          =   225
      Left            =   5265
      TabIndex        =   5
      Top             =   885
      Width           =   1845
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Server Address"
      Height          =   210
      Left            =   5235
      TabIndex        =   4
      Top             =   45
      Width           =   1860
   End
   Begin VB.Label Label4 
      Caption         =   "Port:"
      Height          =   180
      Left            =   5190
      TabIndex        =   2
      Top             =   630
      Width           =   1200
   End
End
Attribute VB_Name = "FrmChatB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdconnect_Click()
On Error GoTo errorhandle
wchat.Close
Cmddc.Enabled = True
FrmChatB.Width = 5100
FrmChatB.Height = 3090
wchat.RemoteHost = txtha.Text
wchat.RemotePort = txtHP.Text
wchat.Connect
cmdconnect.Enabled = False
txtCB.Enabled = True
txtCB.SetFocus
wchat.SendData txtnick.Text + " Connected"
txtha.Enabled = False
txtHP.Enabled = False
txtnick.Enabled = False

'~~~~~~~~~~~~~~~~~~~~~
errorhandle:
    Select Case Err.Number
    Case 10049
        MsgBox "I CAN NOT CONNECT! Either the server is not up right now, or you do not have correct address in the address box."
        Unload Me
        Load Me
    End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Cmddc_Click()
On Error Resume Next
txtCB.Text = "***" + txtnick.Text + " Disconnected***"
wchat.SendData txtCB.Text
txtCB.Text = ""
Timer1.Enabled = True
Cmddc.Enabled = False
txtha.Enabled = True
txtHP.Enabled = True
txtnick.Enabled = True
FrmChatB.Width = 7440
FrmChatB.Height = 2625
End Sub

Private Sub Form_Load()
FrmChatB.Visible = True
FrmChatB.KeyPreview = True
FrmChatB.Height = 2625
FrmChatB.Width = 7440
End Sub

Private Sub Timer1_Timer()
wchat.Close
Cmddc.Enabled = False
'cmdconnect.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub txtCB_KeyPress(KeyAscii As Integer)
On Error GoTo errorhandle
If KeyAscii = 13 Then
KeyAscii = 0
'~~~~~Sends the reload command to the server
If txtCB.Text = "/server reload" Then
wchat.SendData txtCB.Text
txtCB.Text = ""
txtHP.Text = txtHP.Text + 1
Cmddc.Value = True
cmdconnect.Value = True
Exit Sub
Stop
End If
'~~~~~Sends the Close command to the server
If txtCB.Text = "/server close" Then
wchat.SendData txtCB.Text
txtCB.Text = ""
Exit Sub
End If
'~~~~~Sends the text to the server
wchat.SendData ("<" + txtnick.Text + "> " + txtCB.Text)
txtCB.Text = ""
End If
errorhandle:
    Select Case Err.Number
        Case 40006
        MsgBox "I CAN NOT CONNECT! Either the server is not up right now, or you do not have correct address in the address box."
        cmdconnect.Enabled = True
        txtCB.Enabled = False
    End Select
    
End Sub



Private Sub txtHa_Change()
cmdconnect.Enabled = True
If txtnick.Text = "" Then: cmdconnect.Enabled = False
If txtha.Text = "" Then: cmdconnect.Enabled = False
If txtHP.Text = "" Then: cmdconnect.Enabled = False

End Sub

Private Sub txtHP_Change()
cmdconnect.Enabled = True
If txtnick.Text = "" Then: cmdconnect.Enabled = False
If txtha.Text = "" Then: cmdconnect.Enabled = False
If txtHP.Text = "" Then: cmdconnect.Enabled = False
End Sub

Private Sub txtnick_Change()
cmdconnect.Enabled = True
If txtnick.Text = "" Then: cmdconnect.Enabled = False
If txtha.Text = "" Then: cmdconnect.Enabled = False
If txtHP.Text = "" Then: cmdconnect.Enabled = False
End Sub

Private Sub txtrec_Change()
txtrec.SelStart = Len(txtrec) + 1
End Sub

Private Sub wchat_ConnectionRequest(ByVal requestID As Long)
If wchat.State <> sckClosed Then wchat.Close
    wchat.Accept requestID
End Sub

Private Sub WChat_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    wchat.GetData strData, vbString
    txtrec.SelText = vbCrLf & strData
End Sub
Private Sub mnuopt_Click()
Frmopt.Visible = True
End Sub

