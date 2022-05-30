VERSION 5.00
Begin VB.Form frmTalkBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   210
   ClientTop       =   855
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblSpeech 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Speech"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5430
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTalkBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
If ClickToEnd = True Then 'Close the window if there are no buttons, and the form is clicked
    frmTalkBox.Hide 'note:do not unload, just hide.
End If
End Sub

Private Sub Form_GotFocus()
Static bInHere As Boolean
    If ClickToEnd = False And bInHere = False Then    'If there are buttons, but we haven't gotten a value yet,
        If Choice = NONE Then
            bInHere = True
            frmButtons.Show vbModal ' show the Buttons form in modal mode
        End If
        'hide Talkbox after showing the buttons. Or if we've already gotten a value, go ahead and hide us.
        frmTalkBox.Hide
    End If
    bInHere = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_Click  'close the window when the user presses a key.
End Sub

Private Sub lblSpeech_Click()
'If there are no buttons, close the window if clicked; This was put in because the user may click the label, rather than the form
    Form_Click
End Sub
