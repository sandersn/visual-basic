VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmSpatter 
   Caption         =   "Spatter Special Effect"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIterations 
      Height          =   285
      Left            =   8040
      MaxLength       =   9
      TabIndex        =   7
      Text            =   "10000"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Style"
      Height          =   1095
      Left            =   7200
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
      Begin VB.CheckBox chkFromPicture 
         Caption         =   "Spatter from &Picture"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkZoned 
         Caption         =   "&Zoned"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSpatter 
      Caption         =   "&Spatter"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.PictureBox picBackBuffer 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   240
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   120
      Width           =   4800
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   240
      Picture         =   "Spatter.frx":0000
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   4200
      Width           =   4800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of Iterations:"
      Height          =   195
      Left            =   6480
      TabIndex        =   8
      Top             =   1350
      Width           =   1470
   End
End
Attribute VB_Name = "frmSpatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bZoned As Boolean
Dim bFromPic As Boolean

Private Sub chkFromPicture_Click()
    If chkFromPicture.Value = vbChecked Then
        bFromPic = True
    Else
        bFromPic = False
    End If
End Sub

Private Sub chkZoned_Click()
    If chkZoned.Value = vbChecked Then
        bZoned = True
    Else
        bZoned = False
    End If
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Hoho  'quit if they cancel
    CommonDialog1.ShowOpen
    picImage.Picture = LoadPicture(CommonDialog1.filename)
    picBackBuffer.Height = picImage.Height
    picBackBuffer.Width = picImage.Width
Hoho:
End Sub
End Sub

Private Sub cmdSpatter_Click()
    cmdSpatter.Enabled = False
    cmdLoad.Enabled = False
    picBackBuffer.Cls
    If bZoned Then
        SpatterZoned txtIterations.Text
    Else
        Spatter txtIterations.Text
    End If
    cmdSpatter.Enabled = True
    cmdLoad.Enabled = True
End Sub

Private Sub Form_Load()
    Randomize Timer
End Sub

Private Sub Spatter(lngIterate As Long)
Dim lngColor As Long, i As Long
Dim X As Integer, Y As Integer
        lngColor = QBColor(1)   'set it to blue. Then in the loop we'll check to see if frompic is checked
        For i = 0 To lngIterate Step 1
            X = Int(Rnd * picImage.Width) + 1
            Y = Int(Rnd * picImage.Height) + 1
            If bFromPic Then    'we have to get the actual color from the picture
                lngColor = picImage.Point(X, Y)
            End If
            picBackBuffer.PSet (X, Y), lngColor
        Next i
        If bFromPic = True Then
        'finish by blitting the picture onto the Buffer
            Win32.BitBlt picBackBuffer.hDC, 0, 0, picImage.Width, picImage.Height, picImage.hDC, 0, 0, Win32.SRCCOPY
        Else
            'or else just bf the whole thing solid
            picBackBuffer.Line (0, 0)-(picImage.Width, picImage.Height), lngColor, BF
        End If
End Sub

Private Sub SpatterZoned(lngIterate As Long)
Dim lngColor As Long, i As Long
Dim X As Integer, Y As Integer
        lngColor = QBColor(1)   'set it to blue. Then in the loop we'll check to see if frompic is checked
        For i = 0 To lngIterate Step 1
        'we need to divide up the picbox into 5 zones.
            If i < lngIterate / 10 Then 'start.
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * (picImage.Height / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor

            ElseIf i < lngIterate / 5 Then 'half way thru first zone--start second zone(i.e. (2 * lngIterate) / 10)
                'but keep going at it with the first zone still
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * (picImage.Height / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
                'now get a pixel into the second zone, too.
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((2 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i = lngIterate / 5 Then 'done with first zone--stop it and blit the picture onto that part
                If bFromPic = True Then
                    Win32.BitBlt picBackBuffer.hDC, 0, 0, picImage.Width, picImage.Height / 5, picImage.hDC, 0, 0, Win32.SRCCOPY
                Else
                    picBackBuffer.Line (0, 0)-(picImage.Width, picImage.Height / 5), lngColor, BF
                End If
                
            ElseIf i < (lngIterate * 3) / 10 Then   'just spatter the second section.
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((2 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i < (lngIterate * 2) / 5 Then    'i.e. less than 4/10 thru(spatter on 2nd & 3rd zones)
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((2 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((3 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i = (lngIterate * 2) / 5 Then
                If bFromPic = True Then
                    Win32.BitBlt picBackBuffer.hDC, 0, 0, picImage.Width, (2 * picImage.Height) / 5, picImage.hDC, 0, 0, Win32.SRCCOPY
                Else
                    picBackBuffer.Line (0, 0)-(picImage.Width, (2 * picImage.Height) / 5), lngColor, BF
                End If
            ElseIf i < lngIterate / 2 Then  'less than 1/2 way thru--just spatter the 3rd section.
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((3 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i < (lngIterate * 3) / 5 Then    'parts 3 & 4
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((3 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((4 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i = (lngIterate * 3) / 5 Then
                If bFromPic = True Then
                    Win32.BitBlt picBackBuffer.hDC, 0, 0, picImage.Width, (3 * picImage.Height) / 5, picImage.hDC, 0, 0, Win32.SRCCOPY
                Else
                    picBackBuffer.Line (0, 0)-(picImage.Width, (3 * picImage.Height) / 5), lngColor, BF
                End If
            ElseIf i < (7 * lngIterate) / 10 Then 'less than 7/10 way thru--just spatter the 4th section.
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((4 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i < (lngIterate * 4) / 5 Then    'parts 4 & 5
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * ((4 * picImage.Height) / 5))
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * picImage.Height)
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            ElseIf i = (lngIterate * 4) / 5 Then
                If bFromPic = True Then
                    Win32.BitBlt picBackBuffer.hDC, 0, 0, picImage.Width, (4 * picImage.Height) / 5, picImage.hDC, 0, 0, Win32.SRCCOPY
                Else
                    picBackBuffer.Line (0, 0)-(picImage.Width, (4 * picImage.Height) / 5), lngColor, BF
                End If
            Else    'just spatter to the bottom section.
                X = Int(Rnd * picImage.Width)
                Y = Int(Rnd * picImage.Height)
                If bFromPic Then    'we have to get the actual color from the picture
                    lngColor = picImage.Point(X, Y)
                End If
                picBackBuffer.PSet (X, Y), lngColor
            End If
            DoEvents
        Next i
        If bFromPic = True Then
        'finish by blitting the picture onto the Buffer
            Win32.BitBlt picBackBuffer.hDC, 0, 0, picImage.Width, picImage.Height, picImage.hDC, 0, 0, Win32.SRCCOPY
        Else
            'or else just bf the whole thing solid
            picBackBuffer.Line (0, 0)-(picImage.Width, picImage.Height), lngColor, BF
        End If

End Sub
