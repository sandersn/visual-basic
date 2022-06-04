VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmBlockup 
   Caption         =   "Block Up Special Effect"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin ComCtl2.UpDown updSize 
      Height          =   285
      Left            =   8520
      TabIndex        =   4
      Top             =   2040
      Width           =   195
      _ExtentX        =   265
      _ExtentY        =   503
      _Version        =   327681
      Value           =   10
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtBlockSize"
      BuddyDispid     =   196615
      OrigLeft        =   568
      OrigTop         =   96
      OrigRight       =   581
      OrigBottom      =   115
      Max             =   50
      Min             =   10
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Frame fraType 
      Caption         =   "&Type"
      Height          =   2655
      Left            =   7200
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
      Begin VB.OptionButton optType 
         Caption         =   "Average corners and center"
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optType 
         Caption         =   "Average corners"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optType 
         Caption         =   "Average all"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Caption         =   "Fuzzy Average"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Caption         =   "Point In Center"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Medium"
         Height          =   195
         Left            =   1830
         TabIndex        =   15
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         Height          =   195
         Left            =   2040
         TabIndex        =   14
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         Height          =   195
         Left            =   2085
         TabIndex        =   13
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.TextBox txtBlockSize 
      Height          =   285
      Left            =   7920
      TabIndex        =   3
      Text            =   "4"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load New"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Bitmap|*.bmp|Icon|*.ico"
   End
   Begin VB.CommandButton cmdBlockUp 
      Caption         =   "&Block It!"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   120
      Picture         =   "LieroMatch.frx":0000
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   12
      Top             =   4200
      Width           =   4800
   End
   Begin VB.PictureBox picBackBuffer 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   0
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   0
      Width           =   5625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Block &Size:"
      Height          =   195
      Left            =   7080
      TabIndex        =   2
      Top             =   2070
      Width           =   795
   End
End
Attribute VB_Name = "frmBlockup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum BlockType
    PtInCenter = 0
    Corners = 1
    Corners_Center = 2
    Average = 3
    FuzzyAvg = 4
End Enum
Dim blkType As BlockType
'the Liero colours..for now :)
Private Type Colour
    R As Integer
    G As Integer
    B As Integer
End Type
Const MAX_COLORS = 14
Dim Colors(1 To MAX_COLORS) As Colour


'notes: I can't seem to figure out why the blue seems to be stronger(or else the red/green is weaker)
'later: I figured it out!! The blue and red were REVERSED. If you want to have a cool effect with pictures,
' 1.reverse R & B. 2. remove some or all of the parentheses in the color parsing lines of code
' 3. mess up the for loops so they average 1 pixel too many on each side.
' 4. enjoy the toxic waste look. leave out step 1 to make it look underwater or nighttime.
Private Sub cmdBlockUp_Click()
    cmdBlockUp.Enabled = False
    cmdLoad.Enabled = False
    Select Case blkType
        Case BlockType.Average
            BlockAverage txtBlockSize.Text
        Case BlockType.Corners
            BlockCorner txtBlockSize.Text
        Case BlockType.Corners_Center
            BlockCornerCenter txtBlockSize.Text
        Case BlockType.FuzzyAvg
            BlockFuzzyAvg txtBlockSize.Text
        Case BlockType.PtInCenter
            BlockPtInCenter txtBlockSize.Text
    End Select
    cmdBlockUp.Enabled = True
    cmdLoad.Enabled = True
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Hoho  'quit if they cancel
    CommonDialog1.ShowOpen
    picImage.Picture = LoadPicture(CommonDialog1.filename)
    picBackBuffer.Height = picImage.Height
    picBackBuffer.Width = picImage.Width
Hoho:
End Sub

Private Sub cmdSave_Click()
On Error GoTo Hihi
    CommonDialog1.ShowSave
    SavePicture picBackBuffer.Picture, CommonDialog1.filename
    'this not working because we haven't set the *.picture* property, just drawn on the thing. Not sure how to do this
Hihi:
End Sub

Private Sub Form_Load()
     blkType = Average
    Colors(1).R = 23
    Colors(1).G = 94
    Colors(1).B = 199
    
    Colors(2).R = 59
    Colors(2).G = 234
    Colors(2).B = 193
    
    Colors(3).R = 220
    Colors(3).G = 23
    Colors(3).B = 129
    
    Colors(4).R = 12
    Colors(4).G = 200
    Colors(4).B = 254
    
    Colors(5).R = 0
    Colors(5).G = 0
    Colors(5).B = 255
    
    Colors(6).R = 255
    Colors(6).G = 0
    Colors(6).B = 0
    
    Colors(7).R = 0
    Colors(7).G = 255
    Colors(7).B = 0
    
    Colors(8).R = 128
    Colors(8).G = 128
    Colors(8).B = 128
    
    Colors(9).R = 192
    Colors(9).G = 192
    Colors(9).B = 192
    
    Colors(10).R = 64
    Colors(10).G = 64
    Colors(10).B = 64
    
    Colors(11).R = 32
    Colors(11).G = 32
    Colors(11).B = 32
    
    Colors(12).R = 192
    Colors(12).G = 128
    Colors(12).B = 128
    
    Colors(13).R = 128
    Colors(13).G = 192
    Colors(13).B = 128
End Sub

Private Sub optType_Click(Index As Integer)
    blkType = Index
End Sub

Private Sub BlockAverage(intBlockSize As Integer)
'this is the remastering of the Averaging blockup program to create a Liero colour matcher
Dim X As Integer, Y As Integer
Dim XBlock As Integer, YBlock As Integer
Dim Col As Colour
Dim dCol As Colour
Dim dTotal As Long, lowdTotal As Long, lowIndex As Integer
Dim lngColor As Long
Dim i As Long
    'get the block size first:

    For Y = 0 To picImage.Height Step 1    'loop through all of the blocks
        For X = 0 To picImage.Width Step 1
            'for each pixel, compare to all 'Liero' colours
            lowdTotal = 256 * 3 'reset
            lngColor = picImage.Point(X, Y)
            'now we have to chop the thing up somehow...
'note!!: The comments here appear strange  in memoriam to the huge mix-up where I got confused about the format
'of an RGB long...it &H00BBGGRR not &H00RRGGBB like I thought it was. So the comments make sense in the
'original order.
            Col.R = (lngColor And 255)  'FF0000, FF00, and FF. But I couldn't get the green to work
            Col.G = ((lngColor And 65280) / 256)   'the huge numbers in hex are really:
            Col.B = ((lngColor And 16711680) / 65536)      'will this work ?? Later: Yes it will--I tested it.
            For i = 1 To MAX_COLORS Step 1
                dCol.R = Abs(Col.R - Colors(i).R)
                dCol.G = Abs(Col.G - Colors(i).G)
                dCol.B = Abs(Col.B - Colors(i).B)
                dTotal = dCol.R + dCol.G + dCol.B
                If dTotal < lowdTotal Then
                    lowdTotal = dTotal
                    lowIndex = i
                End If
            Next i
            
            picBackBuffer.PSet (X, Y), RGB(Colors(lowIndex).R, Colors(lowIndex).G, Colors(lowIndex).B)
            
            DoEvents    'make sure we don't hang.
        Next X
    Next Y

End Sub

Private Sub BlockCorner(intBlockSize As Integer)
'the only difference from this procedure from BlockAverage is that the for loops are modified only
'to check the four corners.
Dim X As Integer, Y As Integer
Dim XBlock As Integer, YBlock As Integer
Dim R As Long, G As Long, B As Long, lngcntPixels As Long
Dim lngColor As Long
    'get the block size first:
    If intBlockSize > 50 Or intBlockSize < 2 Then intBlockSize = 4  'if it's out-of-bounds set it to default
    For Y = 0 To (picImage.Height \ intBlockSize) Step 1    'loop through all of the blocks
        For X = 0 To (picImage.Width \ intBlockSize) Step 1
            R = 0: G = 0: B = 0 'reset
            lngcntPixels = 0
            For YBlock = 0 To intBlockSize - 1 Step intBlockSize - 1
                For XBlock = 0 To intBlockSize - 1 Step intBlockSize - 1
                    lngColor = picImage.Point((XBlock + (X * intBlockSize)), (YBlock + (Y * intBlockSize)))
                    R = R + (lngColor And 255)
                    G = G + ((lngColor And 65280) / 256)
                    B = B + ((lngColor And 16711680) / 65536)

                    lngcntPixels = lngcntPixels + 1
                    DoEvents    'make sure we don't hang.
                Next XBlock
            Next YBlock
            'now average the three colors by the total number of pixels
            R = R \ lngcntPixels
            G = G \ lngcntPixels
            B = B \ lngcntPixels
            'now let's draw a boxfill of size intblocksize ^ 2
            picBackBuffer.Line ((X * intBlockSize), (Y * intBlockSize))-Step(intBlockSize - 1, intBlockSize - 1), RGB(R, G, B), BF
        Next X
    Next Y
End Sub

Private Sub BlockPtInCenter(intBlockSize As Integer)
Dim X As Integer, Y As Integer
Dim XBlock As Integer, YBlock As Integer
Dim R As Long, G As Long, B As Long, lngcntPixels As Long
Dim lngColor As Long
    'get the block size first:
    If intBlockSize > 50 Or intBlockSize < 1 Then intBlockSize = 4  'if it's out-of-bounds set it to default
    For Y = 0 To (picImage.Height \ intBlockSize) Step 1    'loop through all of the blocks
        For X = 0 To (picImage.Width \ intBlockSize) Step 1
            'for each block, we need to determine the color in the center of the block.
            'Then we draw a Line...BF of that color
            'to picBackBuffer.
            lngColor = picImage.Point(((intBlockSize / 2) + (X * intBlockSize)), ((intBlockSize / 2) + (Y * intBlockSize)))
            DoEvents    'make sure we don't hang.
            'now average the three colors by the total number of pixels
            picBackBuffer.Line ((X * intBlockSize), (Y * intBlockSize))-Step(intBlockSize - 1, intBlockSize - 1), lngColor, BF
        Next X
    Next Y
End Sub

Private Sub txtBlockSize_GotFocus()
    'swipe the text for the user.
    txtBlockSize.SelStart = 0
    txtBlockSize.SelLength = Len(txtBlockSize.Text)
End Sub

Private Sub BlockFuzzyAvg(intBlockSize As Integer)
Dim X As Integer, Y As Integer
Dim XBlock As Integer, YBlock As Integer
Dim R As Long, G As Long, B As Long, lngcntPixels As Long
Dim lngColor As Long
    'get the block size first:
    If intBlockSize > 50 Or intBlockSize < 1 Then intBlockSize = 4
    For Y = 0 To (picImage.Height \ intBlockSize) Step 1    'loop through all of the blocks
        For X = 0 To (picImage.Width \ intBlockSize) Step 1
            R = 0: G = 0: B = 0 'reset
            lngcntPixels = 0
            For YBlock = 0 To intBlockSize Step 1   'the only diff between Fuzzy and Normal Avg. is
                For XBlock = 0 To intBlockSize Step 1 'this one's intBlockSize is not -1
                    lngColor = picImage.Point((XBlock + (X * intBlockSize)), (YBlock + (Y * intBlockSize)))
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
            'now let's draw a boxfill of size intblocksize ^ 2
            picBackBuffer.Line ((X * intBlockSize), (Y * intBlockSize))-Step(intBlockSize - 1, intBlockSize - 1), RGB(R, G, B), BF
        Next X
    Next Y

End Sub

Private Sub BlockCornerCenter(intBlockSize As Integer)
Dim X As Integer, Y As Integer
Dim XBlock As Integer, YBlock As Integer
Dim R As Long, G As Long, B As Long, lngcntPixels As Long
Dim lngColor As Long
    'get the block size first:
    If intBlockSize > 50 Or intBlockSize < 2 Then intBlockSize = 4  'if it's out-of-bounds set it to default
    For Y = 0 To (picImage.Height \ intBlockSize) Step 1    'loop through all of the blocks
        For X = 0 To (picImage.Width \ intBlockSize) Step 1
            R = 0: G = 0: B = 0 'reset
            lngcntPixels = 0
            For YBlock = 0 To intBlockSize - 1 Step intBlockSize - 1
                For XBlock = 0 To intBlockSize - 1 Step intBlockSize - 1
                    lngColor = picImage.Point((XBlock + (X * intBlockSize)), (YBlock + (Y * intBlockSize)))
                    R = R + (lngColor And 255)
                    G = G + ((lngColor And 65280) / 256)
                    B = B + ((lngColor And 16711680) / 65536)

                    lngcntPixels = lngcntPixels + 1
                    DoEvents    'make sure we don't hang.
                Next XBlock
            Next YBlock
            'now get the center's color as well.
            lngColor = picImage.Point(((intBlockSize / 2) + (X * intBlockSize)), ((intBlockSize / 2) + (Y * intBlockSize)))
            R = R + (lngColor And 255)
            G = G + ((lngColor And 65280) / 256)
            B = B + ((lngColor And 16711680) / 65536)

            lngcntPixels = lngcntPixels + 1

            'now average the three colors by the total number of pixels
            R = R \ lngcntPixels
            G = G \ lngcntPixels
            B = B \ lngcntPixels
            'now let's draw a boxfill of size intblocksize ^ 2
            picBackBuffer.Line ((X * intBlockSize), (Y * intBlockSize))-Step(intBlockSize - 1, intBlockSize - 1), RGB(R, G, B), BF
        Next X
    Next Y

End Sub

Private Sub updSize_Change()
    If blkType = PtInCenter Then    'let's update with every flip of the spin control(that's the real name--really!)
        updSize.Enabled = False
        cmdBlockUp_Click
        updSize.Enabled = True
    End If
End Sub
