VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMidiPlayer 
   Caption         =   "Midi List Player"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   9000
   Icon            =   "MidiPlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Randomize &List"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox chkDisplayMode 
      Caption         =   "M&ode"
      DownPicture     =   "MidiPlayer.frx":0442
      Height          =   855
      Left            =   240
      Picture         =   "MidiPlayer.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Percent Complete"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop &Forever"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Shift &Down"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Shift &Up"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Re&move All"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   5100
      IntegralHeight  =   0   'False
      Left            =   6240
      TabIndex        =   16
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove <<"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add >>"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   5160
      Left            =   3480
      Pattern         =   "*.mid;*.wav"
      TabIndex        =   14
      Top             =   720
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      _Version        =   327681
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Dir&ectory:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F&iles Selected:"
      Height          =   195
      Left            =   6240
      TabIndex        =   15
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File&name:"
      Height          =   195
      Left            =   3480
      TabIndex        =   13
      Top             =   480
      Width           =   675
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5235
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Ready"
      Height          =   195
      Left            =   5760
      TabIndex        =   18
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmMidiPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intPos As Integer
Dim strFilename As String
Dim bDisplayMode As Boolean
Private Function GetRand(intTop As Integer, Optional intBottom As Integer = 0)
    GetRand = Int(Rnd * (intTop - intBottom + 1)) + intBottom
End Function
Private Sub chkDisplayMode_Click()

    bDisplayMode = Not bDisplayMode
    If bDisplayMode Then chkDisplayMode.ToolTipText = "Time Left" Else chkDisplayMode.ToolTipText = "Percent Complete"
    'now we check to see if we're playing NOW; if we are then we have to update the current settings.
    If cmdPlay.Enabled = False Then 'yup we need to update the label NOW
         If bDisplayMode Then
            lblStatus.Caption = "left of " & Fix(MMControl1.Length \ 600) & ":" & ((MMControl1.Length Mod 600) / 10) & " in " & strFilename
        Else
            lblStatus.Caption = "% of " & Fix(MMControl1.Length \ 600) & ":" & ((MMControl1.Length Mod 600) / 10) & " in " & strFilename
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    'check for !selected
    If File1.ListIndex = -1 Then
'        MsgBox "You must select a file before pressing Play"   'don't tell them anything--assume they're not THAT stupid.
        Exit Sub
    End If
    
    If Right$(File1.Path, 1) = "\" Then
       List1.AddItem File1.Path & File1.filename
    Else
       List1.AddItem File1.Path & "\" & File1.filename 'root dir case
    End If
    List1.ListIndex = List1.ListCount - 1
    File1.SetFocus
End Sub

Private Sub cmdDown_Click()
Dim strTemp As String
Dim lngIndex As Long
With List1
    If .ListIndex <> -1 And .ListIndex < .ListCount - 1 Then
        lngIndex = .ListIndex
        strTemp = .List(.ListIndex)
        .RemoveItem .ListIndex
        .AddItem strTemp, lngIndex + 1
        .ListIndex = lngIndex + 1
    End If
End With

End Sub

Private Sub cmdPause_Click()
Static intStart As Integer
With MMControl1
    If .Mode = mciModePlay Then
        intStart = .Position
        .Command = "Pause"
        cmdStop.Enabled = False
    ElseIf .Mode = mciModePause Then
        .From = intStart
        .Notify = True
        .Wait = False
        .Command = "Play"
        cmdStop.Enabled = True
    Else
        'ignore!!
    End If
End With
End Sub

Private Sub cmdPlay_Click()
Dim intStart As Integer, intTemp As Integer
    If List1.ListCount = 0 Then Exit Sub    'no songs selected!
    cmdPause.Enabled = True
    cmdPlay.Enabled = False
    cmdStop.Enabled = True
'    cmdRemove.Enabled = False
'    cmdRemoveAll.Enabled = False
'    cmdUp.Enabled = False
'    cmdDown.Enabled = False
    cmdPause.SetFocus   'this may need to go after the .listindex command...
    intPos = 0
    List1.ListIndex = 0
    MMControl1.Shareable = False
    MMControl1.filename = List1.List(intPos)
    Select Case LCase$(Right$(MMControl1.filename, 3))
        Case "wav"
            MMControl1.DeviceType = "WaveAudio"
        Case "mid"
            MMControl1.DeviceType = "Sequencer"
    End Select
    'first set the appropriate properties and Open the Mci.
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Command = "Open"
    'put do loop here that gets filename
    Do
        intStart = intTemp
        intTemp = InStr(intTemp + 1, MMControl1.filename, "\")
    Loop Until intTemp = 0
    strFilename = Right$(MMControl1.filename, Len(MMControl1.filename) - intStart)

    'set the fixed part of the status
    If bDisplayMode Then
        lblStatus.Caption = "left of " & Fix(MMControl1.Length \ 600) & ":" & ((MMControl1.Length Mod 600) / 10) & " in " & strFilename
    Else
        lblStatus.Caption = "% of " & Fix(MMControl1.Length \ 600) & ":" & ((MMControl1.Length Mod 600) / 10) & " in " & strFilename
    End If
    'then Play the Mci -- after setting the properties AGAIN
    MMControl1.Notify = True
    MMControl1.Wait = False
    MMControl1.Command = "Play"
End Sub

Private Sub cmdRandom_Click()
Dim strTemp As String
Dim i As Integer
Dim lngRand As Long
'first off we'll do a check to see how many songs are in the list.
With List1
    If .ListCount <= 1 Then    'quit as fast as possible
        Exit Sub
    ElseIf .ListCount = 2 Then 'just switch 'em
        strTemp = .List(1)
        .RemoveItem 1
        .AddItem strTemp, 0
    ElseIf .ListCount = 3 Then  'flip the top and bottom
        'move bottom to top
        strTemp = .List(2)
        .RemoveItem 2
        .AddItem strTemp, 0
        'now top to bottom
        strTemp = .List(1)
        .RemoveItem 1
        .AddItem strTemp
    Else    'we have to do a full random shuffle
        For i = 1 To List1.ListCount Step 1
            lngRand = GetRand(.ListCount - 1)
            strTemp = .List(lngRand)
            .RemoveItem lngRand
            .AddItem strTemp, GetRand(.ListCount - 1)
        Next i
    End If
End With
End Sub

Private Sub cmdRemove_Click()
Dim intTemp As Integer
    If List1.ListIndex > -1 Then    'make sure we have a selection
        intTemp = List1.ListIndex
        List1.RemoveItem List1.ListIndex
        If intTemp > List1.ListCount - 1 Then
            List1.ListIndex = intTemp - 1
        Else
            List1.ListIndex = intTemp
        End If
    End If
End Sub

Private Sub cmdRemoveAll_Click()
    List1.Clear
End Sub

Private Sub cmdStop_Click()
    If MMControl1.Mode <> mciModePlay Then Exit Sub
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    cmdPause.Enabled = False
    cmdPlay.Enabled = True
'    cmdAdd.Enabled = True
'    cmdRemove.Enabled = True
'    cmdRemoveAll.Enabled = True
'    cmdUp.Enabled = True
'    cmdDown.Enabled = True
    cmdStop.Enabled = False
    File1.SetFocus
    lblStatus.Caption = "Ready"
    lblPercent.Caption = ""
    Me.Caption = "Midi List Player"
End Sub


Private Sub cmdUp_Click()
Dim strTemp As String
Dim lngIndex As Long
With List1
    If .ListIndex > 0 Then
        lngIndex = .ListIndex
        strTemp = .List(.ListIndex)
        .RemoveItem .ListIndex
        .AddItem strTemp, lngIndex - 1
        .ListIndex = lngIndex - 1
    End If
End With
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub


Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dir1.Path = Dir1.List(Dir1.ListIndex)
    End If
End Sub

Private Sub Drive1_Change()
On Error GoTo Mistake
'Static strDrv As String
'strDrv = Drive1.Drive
Dir1.Path = Drive1.Drive
Mistake:
    If Err.Number = 68 Then Drive1.Drive = "c:": Exit Sub
    'this is a hack because I can't figure out how to reset the drive to the 'old' drive letter, so you'd better have a working c:\ drive for this :)
    'see above commented lines o' code for my efforts.
End Sub


Private Sub File1_DblClick()
    cmdAdd_Click
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdAdd_Click
        If File1.ListIndex < File1.ListCount - 1 Then
            File1.ListIndex = File1.ListIndex + 1
        Else
            'warning, if you press enter and hold it down, funny things happen when U reach the end, namely U get the last 2 songs alterated until the listbox maxes
            'out.
            File1.ListIndex = File1.ListIndex - 1
        End If
    End If
End Sub

Private Sub Form_Load()
    bDisplayMode = False    'false==Percent
    Randomize Timer 'this for the random function.
End Sub

Private Sub Form_Resize()
'put code here to resize all the list boxes--height only for the file and dir and both height and width for the true list

    With List1  'listbox first--no good reason.
Dim Rgt As Integer, Bottom As Integer
        Rgt = .Left + .Width
        Bottom = .Top + .Height
        Rgt = (frmMidiPlayer.ScaleWidth - Rgt) + .Width
        If Rgt > 17 Then .Width = Rgt
        Bottom = (frmMidiPlayer.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With
    
    With Dir1
        Bottom = .Top + .Height
        Bottom = (frmMidiPlayer.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With
    With File1
        Bottom = .Top + .Height
        Bottom = (frmMidiPlayer.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With
'unused because we don't put the buttons below anything anymore.
'    With cmdUp  'the most used cmd in this code
'        .Top = Me.ScaleHeight - .Height - 128
'        cmdDown.Top = Me.ScaleHeight - .Height - 128
'        chkDisplayMode.Top = Me.ScaleHeight - .Height - 128
'        cmdAdd.Top = Me.ScaleHeight - (.Height * 2) - 256
'        cmdRemove.Top = Me.ScaleHeight - (.Height * 2) - 256
'        cmdRemoveAll.Top = Me.ScaleHeight - (.Height * 2) - 256
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub

Private Sub List1_DblClick()
    'If MMControl1.Mode = mciModeNotOpen Then   'we allow them to remove any time they want now
    cmdRemove_Click
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then 'And MMControl1.Mode = mciModeNotOpen Then
        cmdRemove_Click
    End If
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
Dim bRewind As Boolean
Dim intStart As Integer
Dim intTemp As Integer

    
    If NotifyCode = mciAborted Then
        If MMControl1.Mode = mciModePause Then Exit Sub 'make sure we just haven't paused...
        'first close the open file...
        MMControl1.Notify = False
        MMControl1.Wait = True
        MMControl1.Command = "Close"
        lblStatus.Caption = "Ready"
        lblPercent.Caption = ""
        Me.Caption = "Midi List Player"
        cmdPause.Enabled = False
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
'        cmdRemove.Enabled = True
'        cmdRemoveAll.Enabled = True
'        cmdUp.Enabled = True
'        cmdDown.Enabled = True
        Exit Sub
    ElseIf intPos >= List1.ListCount - 1 Then   'the bigger than is for when stupid users delete
    'a lot of songs after chunking a bunch in.
        If chkLoop.Value = vbChecked Then
            intPos = -1 'this because we increment it by 1 a couple lines down.
            If List1.List(0) = List1.List(List1.ListCount - 1) Then bRewind = True
        Else
            'first close the open file...
            MMControl1.Notify = False
            MMControl1.Wait = True
            MMControl1.Command = "Close"
            lblStatus.Caption = "Ready"
            lblPercent.Caption = ""
            Me.Caption = "Midi List Player"
            cmdPause.Enabled = False
            cmdPlay.Enabled = True
            cmdStop.Enabled = False
'            cmdAdd.Enabled = True
'            cmdRemove.Enabled = True
'            cmdRemoveAll.Enabled = True
'            cmdUp.Enabled = True
'            cmdDown.Enabled = True
            Exit Sub 'then quit!
        End If
    End If
    
    intPos = intPos + 1
    List1.ListIndex = intPos
    
    If bRewind = False And intPos <> 0 Then 'we need to make sure and check to see if the next filename is identical...
        If List1.List(List1.ListIndex) = List1.List(List1.ListIndex - 1) And List1.ListIndex <> List1.ListCount - 1 Then bRewind = True
    End If
    
    If bRewind = False Then
        'first close the open file...
        MMControl1.Notify = False
        MMControl1.Wait = True
        MMControl1.Command = "Close"
        
        MMControl1.Shareable = False
        MMControl1.filename = List1.List(intPos)
        Select Case LCase$(Right$(MMControl1.filename, 3))
            Case "wav"
                MMControl1.DeviceType = "WaveAudio"
            Case "mid"
                MMControl1.DeviceType = "Sequencer"
        End Select
        'first set the appropriate properties and Open the Mci.
        MMControl1.Notify = False
        MMControl1.Wait = True
        MMControl1.Command = "Open"
        'then Play the Mci -- after setting the properties AGAIN
        
        'put do loop here that gets filename
        Do
            intStart = intTemp
            intTemp = InStr(intTemp + 1, MMControl1.filename, "\")
        Loop Until intTemp = 0
        strFilename = Right$(MMControl1.filename, Len(MMControl1.filename) - intStart)
        'set the fixed part of the status
        If bDisplayMode Then
            lblStatus.Caption = "left of " & Fix(MMControl1.Length \ 600) & ":" & ((MMControl1.Length Mod 600) / 10) & " in " & strFilename
        Else
            lblStatus.Caption = "% of " & Fix(MMControl1.Length \ 600) & ":" & ((MMControl1.Length Mod 600) / 10) & " in " & strFilename
        End If
        MMControl1.Notify = True
        MMControl1.Wait = False
        MMControl1.Command = "Play"
    Else
        'just rewind the current file, then play again.
        MMControl1.Notify = False
        MMControl1.Wait = True
        MMControl1.To = 0
        MMControl1.Command = "Seek"
        MMControl1.Notify = True
        MMControl1.Wait = False
        MMControl1.Command = "Play"
    End If
End Sub

Private Sub MMControl1_StatusUpdate()
With MMControl1
If Me.WindowState = vbMinimized Then
    'put a percent  in the caption
    If bDisplayMode Then
        Me.Caption = Fix((.Length - .Position) \ 600) & ":" & Fix(((.Length - .Position) Mod 600) / 10) & " - " & strFilename
    Else
        Me.Caption = CInt((.Position / .Length) * 100) & "%" & " - " & strFilename
    End If
Else
    If bDisplayMode Then
        lblPercent.Caption = Fix((.Length - .Position) \ 600) & ":" & Fix(((.Length - .Position) Mod 600) / 10)
        Me.Caption = Fix((.Length - .Position) \ 600) & ":" & Fix(((.Length - .Position) Mod 600) / 10) & " (" & strFilename & ")" & " - " & "Midi List Player"
    Else
        lblPercent.Caption = CInt((.Position / .Length) * 100)
        Me.Caption = CInt((.Position / .Length) * 100) & "%" & " (" & strFilename & ")" & " - " & "Midi List Player"
    End If
End If
End With
End Sub
