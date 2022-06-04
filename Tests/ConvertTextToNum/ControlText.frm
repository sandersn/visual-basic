VERSION 5.00
Begin VB.Form frmControlText 
   Caption         =   "Type text or numbers. * is wildcard."
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Alphabet order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
      Begin VB.OptionButton optAlphaOrder 
         Caption         =   "abc, ABC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optAlphaOrder 
         Caption         =   "ABC, abc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label lblMode 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click for About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblText 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4545
   End
   Begin VB.Label lblOutput 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4545
   End
End
Attribute VB_Name = "frmControlText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data(1 To 100) As Integer
Dim TextTyped(1 To 100) As String
Dim CursorPos As Integer
Dim curNum As Variant
Dim bAlphaOrder As Boolean

Private Sub lblMode_Click()
Dim msg As String
    msg = "In thingy32 this will be the switch between Relative and Table Search. Please test this program to see if the entry works. Currently there is no search code behind this, so don't expect search results!" & vbCrLf
    msg = msg & "I couldn't get anything besides a label to work for this because the other controls don't give me the fine control over key handling."
    msg = msg & vbCrLf & "One note: do you like the labels with borders better than the ones with just brackets [] like in thingy32?"
    MsgBox msg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim offset As Integer
    If KeyAscii > 64 And KeyAscii < 91 Then 'UCase
        If curNum <> "" Then
            Form_KeyPress 32    'commit number before we process this letter.
        End If
        offset = IIf(bAlphaOrder, 26, 0)
        Data(CursorPos) = KeyAscii - 64 + offset
        TextTyped(CursorPos) = Chr$(KeyAscii)
        CursorPos = CursorPos + 1
        curNum = ""
    ElseIf KeyAscii > 96 And KeyAscii < 123 Then    'LCase
        If curNum <> "" Then
            Form_KeyPress 32    'commit number before we process this letter.
        End If
        offset = IIf(bAlphaOrder, 0, 26)
        Data(CursorPos) = KeyAscii - 96 + offset
        TextTyped(CursorPos) = Chr$(KeyAscii)
        CursorPos = CursorPos + 1
        curNum = ""
    ElseIf KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = Asc("-") Then
        curNum = curNum & Chr$(KeyAscii)    'cat the number just typed
    ElseIf KeyAscii = vbKeyBack Then
        If curNum <> "" Then    'if there is a pending number...
            curNum = Left$(curNum, Len(curNum) - 1) 'trunc curNum by one    'not working.
        Else
            If CursorPos > 1 Then CursorPos = CursorPos - 1
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If curNum <> "" Then
            If curNum < 32767 Then
            'make sure we have a number, and that it's not an overflow.
                Data(CursorPos) = curNum
                TextTyped(CursorPos) = curNum
                CursorPos = CursorPos + 1
                curNum = ""
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If curNum <> "" Then
            Form_KeyPress 32    'commit number before we process this letter.
        End If
        If CursorPos > 1 Then   'don't let them type wildcards at the beginning of a search
            Data(CursorPos) = 32767
            TextTyped(CursorPos) = "*"
            CursorPos = CursorPos + 1
            curNum = ""
        End If
    End If
    PaintData
End Sub
Private Sub PaintData()
Dim i As Integer
    'here paint the int array and the string(string part should be easy)
    With lblOutput
    .Caption = ""
    lblText.Caption = ""
    For i = 1 To CursorPos - 1 Step 1
        If Data(i) = 32767 Then
            .Caption = .Caption & "* "
        Else    'normal
            .Caption = .Caption & Data(i) & " "
        End If
        lblText.Caption = lblText.Caption & TextTyped(i) & " "
    Next i
    If curNum <> "" Then    'paint the number being entered as well
        lblText.Caption = lblText.Caption & curNum
    End If
    End With
End Sub

Private Sub Form_Load()
    CursorPos = 1
End Sub

Private Sub optAlphaOrder_GotFocus(Index As Integer)
    'flip order
    If Index = 0 Then bAlphaOrder = False Else bAlphaOrder = True
    'clear data
    Erase Data
    Erase TextTyped
    curNum = ""
    CursorPos = 1
    lblText.Caption = ""
    lblOutput.Caption = ""
End Sub

Private Sub optAlphaOrder_KeyPress(Index As Integer, KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub
