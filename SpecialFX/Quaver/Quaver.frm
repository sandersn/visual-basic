VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmQuaver 
   Caption         =   "Quaver Fx(from Snes)"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   643
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   7680
      TabIndex        =   9
      Text            =   "1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   7680
      TabIndex        =   7
      Text            =   "100"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   7680
      TabIndex        =   5
      Text            =   "3"
      Top             =   1320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      Filter          =   "Bitmaps|*.bmp|Icons|*.ico"
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuaver 
      Caption         =   "&Start Quaver"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox picBackBuffer 
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   120
      ScaleHeight     =   204
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   379
      TabIndex        =   1
      Top             =   120
      Width           =   5685
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   120
      Picture         =   "Quaver.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   4680
      Width           =   5625
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Width:"
      Height          =   195
      Left            =   7170
      TabIndex        =   8
      Top             =   1830
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Delay (msec):"
      Height          =   195
      Left            =   6675
      TabIndex        =   6
      Top             =   2310
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Height:"
      Height          =   195
      Left            =   7125
      TabIndex        =   4
      Top             =   1350
      Width           =   510
   End
End
Attribute VB_Name = "frmQuaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bStop As Boolean
Private Const RUNLEN = 5
'warning!! Slope is still not working!!
'next day: I'm commenting out slope until I try something else.

Private Sub cmdLoad_Click()
On Error GoTo Hoho
    CommonDialog1.ShowOpen
    picImage.Picture = LoadPicture(CommonDialog1.filename)
    picBackBuffer.Height = picImage.Height
    picBackBuffer.Width = picImage.Width
Hoho:
End Sub

Private Sub cmdQuaver_Click()
Static bGoing As Boolean
    If bGoing = False Then  'start it up.
        cmdQuaver.Caption = "&Stop Quaver"
        bGoing = True
        Quaver txtHeight.Text, txtWidth.Text, txtDelay.Text
    Else    'stop it!
        cmdQuaver.Caption = "&Start Quaver"
        bStop = True
        bGoing = False
    End If
End Sub

Private Sub Quaver(intHeight As Integer, intWidth As Integer, sngWait As Single)
Dim i As Long
Dim SaveddX As Integer, dX As Integer   'these names come from my calculus background??
Dim SavedXOffset As Integer, XOffset As Integer
Dim sngTime As Single
'Dim intRun As Integer
Dim intDiff As Integer  'the number of '1 bigger values' we need to include to make up the complete
Dim intBottom As Integer, intTop As Integer 'intWidth(versus the 'normal values' which intWidth \ intHeight
Dim dXArray() As Integer   'this to hold the dX values which are pre-calced for no good reason.
    SaveddX = 1 '(1 / sngSlope)  'start at plus one.
    dX = 1
    'convert wait time sec from msec
    sngWait = sngWait / 1000
    If intHeight < -1 Or intHeight > picImage.Height Then intHeight = 3    'default.
    ReDim dXArray(intHeight - 1)
    'now we figure out how much the height and width are and how much the offset should be for
    'each line.
    intBottom = intWidth \ intHeight    ' ex: 11 \ 3 = 3
    intTop = intBottom + 1                  'duh
    intDiff = intWidth - (intBottom * intHeight)    'ex: 11 - (3 * 3) = 2
    'that's how many 'top' values we'll need mixed with the bottom values.
    
    'now we put intDiff number of intTop values in an array of intHeight length spaced at intervals of
    '(iHeight / iDiff). Rounded of course to the *nearest number*. Not truncated as I sometimes do things.
    For i = 0 To intHeight - 1 Step 1
        dXArray(i) = intBottom  'set it to bottom val to start.
        If (i Mod (intHeight / intDiff)) = 0 Then   'it's a top value
            dXArray(i) = intTop
        End If
    Next i
    Do While bStop = False  'loop until they press the start/stop button
        sngTime = Timer
        'here we need to keep track of the current 'global Xoffset' plus the direction to *start* traveling--my code before
        'just used whatever was in at the *end* of the previous cycle.
        'now that we're done with a cycle, we diddle[<- Canadaspeak] the offset by one and then start again.
        If Abs(SavedXOffset) = intHeight Then
            If SavedXOffset = intHeight Then
                SaveddX = -1 '(-1 / sngSlope) 'go backwards; we have been going forwards
            Else 'If XOffset = -intMagnitude Then
                SaveddX = 1 '(1 / sngSlope) 'go forwards instead of backwards
            End If
'            intRun = RUNLEN + 1 'make sure we don't leap back and forth since we're at the edge of intmagnitude still.
        End If
        XOffset = SavedXOffset
'        dX = SaveddX
        SavedXOffset = SavedXOffset + SaveddX * (dXArray(SavedXOffset Mod intHeight))   'so next time we'll know what to do too.
        For i = 0 To picImage.Height Step 1
            If Abs(XOffset) = intHeight Then  'we need to change directions because we're at the
                        'end of a cycle
'                If intRun < RUNLEN Then
'                    dX = 0  'make sure we don't move.
'                    intRun = intRun + 1
                If XOffset = intHeight Then
                    dX = -1 '(-1 / sngSlope) 'go backwards; we have been going forwards
                    'intRun = 0  'reset the run length counter.
                Else 'XOffset = -intMagnitude Then
                    dX = 1 '(1 / sngSlope) 'go forwards instead of backwards
                    'intRun = 0  'reset the run length counter.
                End If
            End If
            XOffset = XOffset + (dX * dXArray(i Mod intHeight))
            Win32.BitBlt picBackBuffer.hDC, XOffset, i, picImage.Width, 1, picImage.hDC, 0, i, Win32.SRCCOPY
        Next i
    
        'delay the right amount of time
        Do Until Timer - sngTime > sngWait
                DoEvents    'don't hang.
        Loop
    Loop
    bStop = False   'reset.
End Sub
Private Sub txtDelay_GotFocus()
With txtDelay
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtHeight_GotFocus()
With txtHeight
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtWidth_GotFocus()
With txtWidth
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
