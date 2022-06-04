VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmFuzz 
   Caption         =   "Fuzzy Fx Demo"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      Filter          =   "Bitmaps|*.bmp|Icons|*.ico"
   End
   Begin VB.TextBox txtBlockSize 
      Height          =   285
      Left            =   6840
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load New"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdFuzz 
      Caption         =   "&Make Fuzzy"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox picBackBuffer 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   120
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   120
      Width           =   5625
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   120
      Picture         =   "Fuzz.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   4440
      Width           =   5625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Block Size:"
      Height          =   195
      Left            =   6000
      TabIndex        =   4
      Top             =   1480
      Width           =   795
   End
End
Attribute VB_Name = "frmFuzz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFuzz_Click()
    cmdFuzz.Enabled = False
    cmdLoad.Enabled = False
    Fuzz txtBlockSize.Text
    cmdFuzz.Enabled = True
    cmdLoad.Enabled = True
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Hoho:
    CommonDialog1.ShowOpen
    picImage.Picture = LoadPicture(CommonDialog1.filename)
    picBackBuffer.Height = picImage.Height
    picBackBuffer.Width = picImage.Width
Hoho:
End Sub

Private Sub Fuzz(intBlockSize As Integer)
Dim X As Integer, Y As Integer
Dim XBlock As Integer, YBlock As Integer
Dim R As Long, G As Long, B As Long, lngcntPixels As Long
Dim lngColor As Long
    'get the block size first:
    If intBlockSize > 50 Or intBlockSize < 2 Then intBlockSize = 3  'if it's out-of-bounds set it to default
    If intBlockSize Mod 2 = 0 Then  'it's even!
        intBlockSize = intBlockSize + 1 'add 1 to it 2 make it odd
    End If
    For Y = 0 To (picImage.Height) Step 1    'loop through all pels
        For X = 0 To (picImage.Width) Step 1
            'for each block, we need to determine the total Rgb for all pixels. Then we divide Rgb by the
            'total number of pixels to determine the average Rgb. Then we draw a point of that color
            'to picBackBuffer.
            R = 0: G = 0: B = 0 'reset
            lngcntPixels = 0
            For YBlock = (Y - (intBlockSize / 2)) To (Y + (intBlockSize / 2)) Step 1
                For XBlock = (X - (intBlockSize / 2)) To (X + (intBlockSize / 2)) Step 1
                    lngColor = picImage.Point(XBlock, YBlock)
                    'now we have to chop the thing up somehow...
'note!!: The comments here appear strange  in memoriam to the huge mix-up where I got confused about the format
'of an RGB long...it &H00BBGGRR not &H00RRGGBB like I thought it was. So the comments make sense in the
'original order.
                    R = R + (lngColor And 255)  'FF0000, FF00, and FF. But I couldn't get the green to work
                    G = G + ((lngColor And 65280) / 256)   'the huge numbers in hex are really:
                    B = B + ((lngColor And 16711680) / 65536)      'will this work ?? Later: Yes it will--I tested it.

                    lngcntPixels = lngcntPixels + 1 'so I used decimal instead(VB is hex unfriendly)
                    DoEvents    'make sure we don't hang.
                Next XBlock
            Next YBlock
            'now average the three colors by the total number of pixels
            R = R \ lngcntPixels
            G = G \ lngcntPixels
            B = B \ lngcntPixels
            'now put a point of that color at X,Y
            picBackBuffer.PSet (X, Y), RGB(R, G, B)
        Next X
    Next Y


End Sub
