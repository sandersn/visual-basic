VERSION 5.00
Begin VB.Form frmHufTree 
   Caption         =   "Test tree Structure implemented via VB Classes"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmHufTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'everything seems to be working except either my traverse-to-store code is buggy and not sorting correctly
'or my traverse-to-display code is buggy, because it *appears* that some matches are not stopping where they ought.
Private Const SIZE = 99999
Private max As Integer

Private Sub Form_Load()
Randomize Timer
Dim i As Long
Dim target(0 To SIZE) As String
Dim bContinue As Boolean
Dim begin As Single
    begin = Timer
    For i = 0 To SIZE
        target(i) = String$(3, Chr$(CInt(Rnd() * 120) + 33))
    Next i
    txtOutput.Text = txtOutput.Text & "Gen time: " & (Timer - begin) & vbCrLf
Dim top As New HufNode 'fill this somehow in real program
    top.Strin = "ccc"
    top.Freq = 1
Dim current As HufNode
    begin = Timer
    For i = 0 To SIZE Step 1
        Set current = top
    'fill algorithm... if it's equal, inc int data.
    'else if less, traverse to left until hit bottom of tree
    'else if more, traverse to right until hit bottom
        Do
            bContinue = False
            If (current.Strin = target(i)) Then
                current.Freq = current.Freq + 1
            ElseIf current.Strin > target(i) Then
                If current.IsLeft Then
                    Set current = current.Left
                    bContinue = True
                Else    'add
                    Set current.Left = New HufNode
                    current.Left.Strin = target(i)
                    current.Left.Freq = 1
                    current.IsLeft = True
                End If
            Else    'current.strin > target
                If current.IsRight Then
                    Set current = current.Right
                    bContinue = True
                Else
                    Set current.Right = New HufNode
                    current.Right.Strin = target(i)
                    current.Right.Freq = 1
                    current.IsRight = True
                End If
            End If
        Loop While (bContinue)
    Next i
    txtOutput.Text = txtOutput.Text & "Process time: " & (Timer - begin) & vbCrLf
    'traverse pre-order algorithm:
    begin = Timer
    max = 0
    PreOrderSearch top
    txtOutput.Text = "Search time: " & (Timer - begin) & vbCrLf & txtOutput.Text
    txtOutput.Text = txtOutput.Text & vbCrLf & vbCrLf & "Max: " & max
End Sub

Private Sub PreOrder(node As HufNode)
    txtOutput.Text = txtOutput.Text & node.Strin & " " & node.Freq & vbCrLf
    If (node.IsLeft) Then
        PreOrder node.Left
    End If
    If (node.IsRight) Then
        PreOrder node.Right
    End If
End Sub
Private Sub PreOrderSearch(node As HufNode)
    If max < node.Freq Then
        max = node.Freq
    End If
    txtOutput.Text = txtOutput.Text & node.Strin & " " & node.Freq & vbCrLf
    If (node.IsLeft) Then
        PreOrderSearch node.Left
    End If
    If (node.IsRight) Then
        PreOrderSearch node.Right
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        txtOutput.Height = Me.ScaleHeight - 12
        txtOutput.Width = Me.ScaleWidth - 12
    End If
End Sub
