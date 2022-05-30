VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7035
   Height          =   5385
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7035
   Top             =   1140
   Width           =   7155
   Begin CommandButton cmdCopy 
      Caption         =   "&Just Copy the Picture Straight to the BackGround"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3720
      Width           =   5055
   End
   Begin CommandButton cmdPicMask 
      Caption         =   "&Both Pic and Mask"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   4080
      Width           =   2535
   End
   Begin CommandButton cmdPic 
      Caption         =   "&Picture Only"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin CommandButton cmdMask 
      Caption         =   "&Mask Only"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin PictureBox picZack_m 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   720
      Picture         =   TSTTRANS.FRX:0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   4080
      Width           =   495
   End
   Begin PictureBox picZack 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   240
      Picture         =   TSTTRANS.FRX:026A
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   4080
      Width           =   495
   End
   Begin PictureBox picBG 
      AutoSize        =   -1  'True
      Height          =   3630
      Left            =   0
      Picture         =   TSTTRANS.FRX:04D4
      ScaleHeight     =   3600
      ScaleWidth      =   6990
      TabIndex        =   0
      Top             =   0
      Width           =   7020
   End
End
Option Explicit
Declare Function BitBlt Lib "GDI" (ByVal hDestdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcdc As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Const PIXEL = 3
Const TWIP = 1
Const SRCAND = &H8800C6
Const SRCCOPY = &HCC0020
Const SRCPAINT = &HEE0086
Dim Pos As Integer

Sub cmdCopy_Click ()
Dim Dummy
    Pos = Pos + 20
    picBG.ScaleMode = PIXEL
    picZack.ScaleMode = PIXEL
    picZack_m.ScaleMode = PIXEL
    Dummy = BitBlt(picBG.hDC, Pos, Pos, picZack.ScaleWidth, picZack.ScaleHeight, picZack.hDC, 0, 0, SRCCOPY)
    picZack.ScaleMode = TWIP
    picBG.ScaleMode = TWIP
    picZack_m.ScaleMode = TWIP
End Sub

Sub cmdMask_Click ()
Dim Dummy
    Pos = Pos + 20
    picBG.ScaleMode = PIXEL
    picZack.ScaleMode = PIXEL
    picZack_m.ScaleMode = PIXEL
    Dummy = BitBlt(picBG.hDC, Pos, Pos, picZack_m.ScaleWidth, picZack_m.ScaleHeight, picZack_m.hDC, 0, 0, SRCAND)
    picZack.ScaleMode = TWIP
    picBG.ScaleMode = TWIP
    picZack_m.ScaleMode = TWIP
End Sub

Sub cmdPic_Click ()
Dim Dummy
    Pos = Pos + 20
    picBG.ScaleMode = PIXEL
    picZack.ScaleMode = PIXEL
    picZack_m.ScaleMode = PIXEL
    Dummy = BitBlt(picBG.hDC, Pos, Pos, picZack.ScaleWidth, picZack.ScaleHeight, picZack.hDC, 0, 0, SRCPAINT)
    picZack.ScaleMode = TWIP
    picBG.ScaleMode = TWIP
    picZack_m.ScaleMode = TWIP

End Sub

Sub cmdPicMask_Click ()
Dim Dummy
    Pos = Pos + 20
    picBG.ScaleMode = PIXEL
    picZack.ScaleMode = PIXEL
    picZack_m.ScaleMode = PIXEL
    Dummy = BitBlt(picBG.hDC, Pos, Pos, picZack_m.ScaleWidth, picZack_m.ScaleHeight, picZack_m.hDC, 0, 0, SRCAND)
    Dummy = BitBlt(picBG.hDC, Pos, Pos, picZack.ScaleWidth, picZack.ScaleHeight, picZack.hDC, 0, 0, SRCPAINT)
    picZack.ScaleMode = TWIP
    picBG.ScaleMode = TWIP
    picZack_m.ScaleMode = TWIP

End Sub

