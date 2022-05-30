VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmThingEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thing Editor"
   ClientHeight    =   5280
   ClientLeft      =   330
   ClientTop       =   600
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtWeight 
      Height          =   285
      Left            =   2220
      TabIndex        =   10
      Text            =   "1"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkImmovable 
      Caption         =   "&Immovable"
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   4000
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Top             =   4620
      Width           =   2535
   End
   Begin VB.ListBox lstMoveStyle 
      Height          =   3570
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin ComctlLib.ListView lvwType 
      Height          =   3615
      Left            =   2220
      TabIndex        =   3
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6376
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   4395
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4395
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "&Weight: Must be between 1 and 255. 0 is immovable."
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   3900
      Width           =   1920
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "&Description Line Number:"
      Height          =   195
      Left            =   3720
      TabIndex        =   8
      Top             =   4400
      Width           =   1785
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Note: X and Y are chosen back in the MapEditor using the Move button."
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   5130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Movement Style:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Type:"
      Height          =   195
      Left            =   2220
      TabIndex        =   2
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "frmThingEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkImmovable_Click()
    If chkImmovable.Value = vbUnchecked Then  'uncheck it!
        With txtWeight
            .Text = "1"
            .Enabled = True
'            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Else    'check it!
        txtWeight.Enabled = False
        txtWeight.Text = "0"
    End If
End Sub

Private Sub cmdCancel_Click()
    frmThingEdit.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtDesc.Text <> "" Then
        If IsNumeric(CInt(txtDesc.Text)) Then
            frmThingEdit.Tag = "OK"
            Me.Hide
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Count As Integer
    lvwType.Icons = frmMapEdit.imlThings
With frmMapEdit.imlThings
For Count = 1 To .ListImages.Count
    lvwType.ListItems.Add Count, , .ListImages(Count).Key, .ListImages(Count).Index
Next
End With
With lstMoveStyle   'ummm...I just now realized this is hard coded--bad, bad! but this is how it stays for now =[
                                'I also just realized that I've been leaving out 'escape' out of the mapeditor even though it's implemented...oops
    .AddItem "Still"
    .AddItem "Random"
    .AddItem "Follow"
    .AddItem "Scripted Path"
    .AddItem "Escape"
    .AddItem "Ship"
End With
End Sub

Private Sub lvwType_ItemClick(ByVal Item As ComctlLib.ListItem)
    If Item.Index < PERSON Then
        lblDesc.Caption = "Description Line Number:"
        chkImmovable.Enabled = True
        chkImmovable_Click
    ElseIf Item.Index >= PERSON And Item.Index < MONSTER Then
        lblDesc.Caption = "Script Number:"
        chkImmovable.Enabled = False
        txtWeight.Enabled = False
        txtWeight.Text = 0
    Else
        lblDesc.Caption = "Unused for Monsters:"
        chkImmovable.Enabled = False
        txtWeight.Enabled = False
        txtWeight.Text = 0
    End If
End Sub
