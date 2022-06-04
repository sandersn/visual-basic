VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   2070
   ClientTop       =   3030
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   10035
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop &Forever(do not click)"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Shift &Down"
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Shift &Up"
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Re&move All"
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   7320
      TabIndex        =   10
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "&Lock"
      Height          =   855
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove <"
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add >"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdEject 
      Caption         =   "&Eject"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3480
      Pattern         =   "*.mid; *.cda; *.wav"
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      _Version        =   327680
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
   Begin VB.Label lblWarning 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Form1.frx":0000
      Height          =   1215
      Left            =   7320
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim intPos As Integer

Private Sub chkLoop_Click()
    MsgBox "This check box does not work yet, because it would infinitely hang this program--not your whole machine, though."
End Sub

Private Sub cmdAdd_Click()
If File1.ListIndex = -1 Then
        MsgBox "You must select a file before pressing Play"
        Exit Sub
End If
    If Right$(File1.Path, 1) = "\" Then
       List1.AddItem File1.Path & File1.filename
    Else
       List1.AddItem File1.Path & "\" & File1.filename
    End If

End Sub

Private Sub cmdDown_Click()
Dim strTemp As String
Dim lngIndex As Long
With List1
    If .ListIndex <> -1 Then
        lngIndex = .ListIndex
        strTemp = .List(.ListIndex)
        .RemoveItem .ListIndex
        .AddItem strTemp, lngIndex + 1
        .ListIndex = lngIndex + 1
    End If
End With

End Sub

Private Sub cmdLock_Click()
    If cmdLock.Caption = "&Lock" Then
        cmdPlay.Enabled = True
        cmdStop.Enabled = True
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdRemoveAll.Enabled = False
        cmdUp.Enabled = False
        cmdDown.Enabled = False
        cmdLock.Caption = "Un&lock"
    ElseIf cmdLock.Caption = "Un&lock" Then
        cmdPlay.Enabled = False
        cmdStop.Enabled = False
        cmdAdd.Enabled = True
        cmdRemove.Enabled = True
        cmdRemoveAll.Enabled = True
        cmdUp.Enabled = True
        cmdDown.Enabled = True
        cmdLock.Caption = "&Lock"
    End If
End Sub

Private Sub cmdPlay_Click()
Dim intPos As Integer

For intPos = 0 To List1.ListCount - 1 Step 1
    MMControl1.Shareable = False
    MMControl1.filename = List1.List(intPos)
    Select Case LCase$(Right$(MMControl1.filename, 3))
        Case "wav"
            MMControl1.DeviceType = "WaveAudio"
        Case "mid"
            MMControl1.DeviceType = "Sequencer"
        '   case "cda"  'this type unsupported in this program!
    End Select
    'first set the appropriate properties and Open the Mci.
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Command = "Open"
    'then Play the Mci -- after setting the properties AGAIN
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Command = "Play"
    'the Close the Mci to wait for the next file -- setting the properties yet ANOTHER time
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Command = "Close"

Next intPos
''****this part past this is old
'If File1.ListIndex = -1 Then
'        MsgBox "You must select a file before pressing Play"
'        Exit Sub
'End If
''Root directories have a backslash, for example "C:\".  Else, need backslash.
'If Right$(File1.Path, 1) = "\" Then
'   MMControl1.filename = File1.Path & File1.filename
'Else
'   MMControl1.filename = File1.Path & "\" & File1.filename
'End If
'
'' Set properties needed by MCI to open.
'MMControl1.Notify = False
'MMControl1.Shareable = False
'
'Select Case UCase$(Right$(MMControl1.filename, 3))
'    Case "WAV"
'        MMControl1.DeviceType = "WaveAudio"
'        'Root directories have a backslash, for example "C:\".  Else, need backslash.
'        If Right$(File1.Path, 1) = "\" Then
'           MMControl1.filename = File1.Path & File1.filename
'        Else
'           MMControl1.filename = File1.Path & "\" & File1.filename
'        End If
'
'
'    Case "MID"
'       MMControl1.DeviceType = "Sequencer"
'       'Root directories have a backslash, for example "C:\".  Else, need backslash.
'        If Right$(File1.Path, 1) = "\" Then
'           MMControl1.filename = File1.Path & File1.filename
'        Else
'           MMControl1.filename = File1.Path & "\" & File1.filename
'        End If
'
'
'    Case "CDA"
'        MMControl1.DeviceType = "CDAudio"
'
'        Dim i As Integer
'        For i = 1 To File1.ListIndex
'            MMControl1.Next
'        Next
'     Case Else
'        MsgBox "I only handle MID, WAV, and CDA(Cd Audio) files."
'        MMControl1.Command = "Close"
'        Exit Sub
'End Select
'
'' Open the MCI WaveAudio device.
'MMControl1.Command = "Open"
'MMControl1.Command = "Play"
'

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
MMControl1.Command = "Stop"
MMControl1.Command = "Close"
List1.ListIndex = List1.ListCount   'set the focus to the end so that MMControl1_Done() will quit out.
End Sub

Private Sub cmdEject_Click()
MMControl1.Command = "Eject"
End Sub

Private Sub cmdUp_Click()
Dim strTemp As String
Dim lngIndex As Long
With List1
    If .ListIndex <> -1 Then
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

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    cmdAdd_Click
End Sub

Private Sub Form_Load()
File1.Pattern = "*.mid;*.wav"
Dir1.Path = "\Games\Midi Music"
End Sub

Private Sub Form_Resize()
'put code here to resize all the list boxes--height only for the file and dir and both height and width for the true list

    With List1  'listbox first--no good reason.
Dim Rgt As Integer, Bottom As Integer
        Rgt = .Left + .Width
        Bottom = .Top + .Height
        Rgt = (Form1.ScaleWidth - Rgt) + .Width
        If Rgt > 17 Then .Width = Rgt
        Bottom = (Form1.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With
    With Dir1
        Bottom = .Top + .Height
        Bottom = (Form1.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With
    With File1
        Bottom = .Top + .Height
        Bottom = (Form1.ScaleHeight - Bottom) + .Height
        If Bottom > 17 Then .Height = Bottom
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub

'Private Sub MMControl1_Done(NotifyCode As Integer)
'Dim i As Integer
''Dim sngEndTime As Single
''    sngEndTime = Timer + 2
''    Do
''    Loop Until Timer > sngEndTime
''****this part is new
''    Debug.Print "MMControl1_Done"
'     If NotifyCode = mciNotifySuccessful Then
'        Debug.Print "Successful!"
'     ElseIf NotifyCode = mciNotifySuperseded Then
'        Debug.Print "Superseded!"
'    ElseIf NotifyCode = mciAborted Then
'        Debug.Print "Aborted"
'    ElseIf NotifyCode = mciFailure Then
'        Debug.Print "Failed"
'    Else
'        MsgBox "Unknown return code for MMControl status."
'    End If
'    If List1.ListCount = intPos + 1 Then  'all done--quit!
'        MMControl1.Command = "Close"
'        Exit Sub
'    End If
'    intPos = intPos + 1 'move to the next song.
'    ' Set properties needed by MCI to open.
'    MMControl1.Shareable = False
'    MMControl1.filename = List1.List(intPos)
'    Select Case UCase$(Right$(MMControl1.filename, 3))
'        Case "WAV"
''                Dim sngEndTime As Single
''        sngEndTime = Timer + 2
''        Do
''        Loop Until Timer > sngEndTime
'
'            MMControl1.DeviceType = "WaveAudio"
''            'Root directories have a backslash, for example "C:\".  Else, need backslash.
''            If Right$(File2.Path, 1) = "\" Then
''               MMControl1.filename = File2.Path & File1.filename
''            Else
''               MMControl1.filename = File2.Path & "\" & File1.filename
''            End If
'
'
'        Case "MID"
'           MMControl1.DeviceType = "Sequencer"
'           'Dim sngEndTime As Single
''        sngEndTime = Timer + 2
''        Do
''        Loop Until Timer > sngEndTime
'
''           'Root directories have a backslash, for example "C:\".  Else, need backslash.
''            If Right$(File2.Path, 1) = "\" Then
''               MMControl1.filename = File2.Path & File1.filename
''            Else
''               MMControl1.filename = File2.Path & "\" & File1.filename
''            End If
'
'
'        Case "CDA"
'            MMControl1.DeviceType = "CDAudio"
'
''            Dim i As Integer
''            For i = 1 To File1.ListIndex
''                MMControl1.Next
''            Next
'         Case Else
'            MsgBox "I only handle MID, WAV, and CDA(Cd Audio) files."
'            MMControl1.Command = "Close"
'            Exit Sub
'    End Select
'
'    ' Open the MCI WaveAudio device.
'    MMControl1.Command = "Open"
'    MMControl1.Notify = True
'    MMControl1.Wait = True
'    MMControl1.Command = "Play"
'    MMControl1.Command = "Close"
'
'
'End Sub
