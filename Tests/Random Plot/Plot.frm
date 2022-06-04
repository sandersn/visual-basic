VERSION 5.00
Begin VB.Form frmPlot 
   Caption         =   "Random Point Plotting(double click to start)"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Point
    x As Integer
    y As Integer
End Type
Dim points() As Point

Private Sub Form_DblClick()
Dim intType As Integer, cPoints As Long, lngMinDist As Long, intHiContrast As Integer
'these are the options returned by frmOptions.
Dim i As Long, j As Long
Dim bTooClose As Boolean
Dim Count As Long
Dim numUnplaced As Long
Dim fColor As Long
    'start new sequence of randomness(show frmOptions)
    frmOptions.Show vbModal, Me
    'now get the options back
    frmOptions.ReturnOptions cPoints, intType, lngMinDist, intHiContrast
    If intHiContrast = vbChecked Then
        fColor = vbWhite
        Me.BackColor = RGB(0, 0, 0)
    Else
        Me.foreColor = vbButtonText
        fColor = vbButtonText
        Me.BackColor = vbButtonFace
    End If
    Me.Cls  'clear. This maybe should be an option in there eventually.
    If intType = 0 Then 'ranDumb
        For i = 1 To cPoints Step 1
            'plot a point!
            Me.PSet (Int(Rnd * Me.ScaleWidth), Int(Rnd * Me.ScaleHeight)), fColor
            DoEvents
        Next i
        MsgBox cPoints & " points plotted.", vbInformation, "Infomeeshon"
    Else    'we hope minimum space
        ReDim points(1 To cPoints) As Point
        For i = 1 To cPoints Step 1
            Count = 0
            Do
                bTooClose = False   'oops have to reset at beginning of every loop the boolean
                'plot a point!
                points(i).x = Int(Rnd * Me.ScaleWidth)  'make up a new point.
                points(i).y = Int(Rnd * Me.ScaleHeight)
                For j = 1 To i - 1 Step 1 'compare this point to all previous points and make sure it's not too close.
                    If dist(points(j), points(i)) < lngMinDist Then
                        bTooClose = True
                        Exit For
                    End If
                Next j
                If Count > 200 Then
                    points(i).x = -99
                    points(i).y = -99    'zero them out(far enough out so as not to harm the rest)
                    numUnplaced = numUnplaced + 1
                    Exit Do
                End If
                Count = Count + 1
            Loop While bTooClose = True 'loop while the point is too close OR it's been tried to plotted 200 times
            Me.PSet (points(i).x, points(i).y), fColor
            Me.Caption = "Processing: " & i & "/" & cPoints
            DoEvents
        Next i
        MsgBox cPoints & " possible points plotted." & vbCrLf & numUnplaced & " unplaced points", vbInformation, "Infomeeshon"
        Me.Caption = "Random Point Plotting(double click to start)"
    End If
    'now that we're done, give a msgbox of statistics.
End Sub

Private Sub Form_Load()
    Randomize Timer
End Sub

Private Function dist(point1 As Point, point2 As Point) As Single
'return the distance between two points
'pass two points(order does not matter)
'return an integer
'how easy.
    dist = (((point1.x - point2.x) ^ 2) + ((point1.y - point2.y) ^ 2)) ^ 0.5
    'note: if you haven't had this in math, too bad. learn it here.
End Function
