VERSION 5.00
Begin VB.Form frmButtons 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (2)"
      Height          =   372
      Index           =   2
      Left            =   204
      TabIndex        =   1
      Top             =   601
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (3)"
      Height          =   372
      Index           =   3
      Left            =   2124
      TabIndex        =   10
      Top             =   601
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (16)"
      Height          =   372
      Index           =   16
      Left            =   204
      TabIndex        =   8
      Top             =   3961
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (14)"
      Height          =   372
      Index           =   14
      Left            =   204
      TabIndex        =   7
      Top             =   3481
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (12)"
      Height          =   372
      Index           =   12
      Left            =   204
      TabIndex        =   6
      Top             =   3001
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (10)"
      Height          =   372
      Index           =   10
      Left            =   204
      TabIndex        =   5
      Top             =   2521
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (0)"
      Height          =   372
      Index           =   0
      Left            =   204
      TabIndex        =   0
      Top             =   121
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (4)"
      Height          =   372
      Index           =   4
      Left            =   204
      TabIndex        =   2
      Top             =   1081
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (6)"
      Height          =   372
      Index           =   6
      Left            =   204
      TabIndex        =   3
      Top             =   1561
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (8)"
      Height          =   372
      Index           =   8
      Left            =   204
      TabIndex        =   4
      Top             =   2041
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (17)"
      Height          =   372
      Index           =   17
      Left            =   2124
      TabIndex        =   17
      Top             =   3961
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (15)"
      Height          =   372
      Index           =   15
      Left            =   2124
      TabIndex        =   16
      Top             =   3481
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (13)"
      Height          =   372
      Index           =   13
      Left            =   2124
      TabIndex        =   15
      Top             =   3001
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (11)"
      Height          =   372
      Index           =   11
      Left            =   2124
      TabIndex        =   14
      Top             =   2521
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (9)"
      Height          =   372
      Index           =   9
      Left            =   2124
      TabIndex        =   13
      Top             =   2160
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (1)"
      Height          =   372
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   121
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (7)"
      Height          =   372
      Index           =   7
      Left            =   2160
      TabIndex        =   12
      Top             =   1561
      Width           =   1812
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Choice (5)"
      Height          =   372
      Index           =   5
      Left            =   2124
      TabIndex        =   11
      Top             =   1081
      Width           =   1812
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChoice_Click(Index As Integer)
    Choice = Index      'Tell the TalkBox sub which button was pressed
    frmButtons.Hide   'Unload the TalkBox forms from memory
'    frmTalkBox.Hide    'this statement now inside frmTalkbox_Activate
End Sub

Private Sub Form_Load()
    'Adjust to the right of the TalkBox form
    frmButtons.Left = (frmTalkBox.Left + frmTalkBox.Width) + 1
    frmButtons.Top = frmTalkBox.Top
End Sub
