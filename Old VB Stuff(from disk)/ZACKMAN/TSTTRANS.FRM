VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "f"
   ClientHeight    =   4020
   ClientLeft      =   1080
   ClientTop       =   1485
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Begin VB.PictureBox picZackmov 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   1
      Left            =   120
      Picture         =   "TSTTRANS.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackmov 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   0
      Left            =   120
      Picture         =   "TSTTRANS.frx":0272
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackmm 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   0
      Left            =   840
      Picture         =   "TSTTRANS.frx":04E4
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZackmm 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   1
      Left            =   840
      Picture         =   "TSTTRANS.frx":0756
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picZack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1440
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Private Declare Function BitBlt Lib "Gdi" (ByVal destHdc, ByVal X, ByVal Y, ByVal w, ByVal h, ByVal srcHdc, ByVal srcX, ByVal srcY, ByVal rop As Long)
Const srccopy = &HCC0020
Const srcand = &H8800C6
Const SRcpaint = &HEE0086

Dim leftedge, topedge
Dim current As Integer
Dim currentm As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim dummy
    Select Case KeyCode
        Case KEY_RIGHT
          Beep
                'current = current + 1
                'currentm = currentm + 1
            
            leftedge = leftedge + 60
            dummy = BitBlt(picZack.hDC, 0, 0, picZack.ScaleWidth, picZack.ScaleHeight, Form1.hDC, leftedge, topedge, srccopy)
            dummy = BitBlt(picZack.hDC, 60, 60, picZackmm(0).ScaleWidth, picZackmm(0).ScaleHeight, picZackmm(0).hDC, 0, 0, srcand)
            dummy = BitBlt(picZack.hDC, 60, 60, picZackmm(0).ScaleWidth, picZackmm(0).ScaleHeight, picZackmov(0).hDC, 0, 0, SRcpaint)
            
            If leftedge > ScaleWidth Then
                leftedge = leftedge - 60
            End If
            
            leftedge = leftedge + 60
            dummy = BitBlt(picZack.hDC, 0, 0, picZack.ScaleWidth, picZack.ScaleHeight, Form1.hDC, leftedge, topedge, srccopy)
            dummy = BitBlt(picZack.hDC, 60, 60, picZackmm(1).ScaleWidth, picZackmm(1).ScaleHeight, picZackmm(1).hDC, 0, 0, srcand)
            dummy = BitBlt(picZack.hDC, 60, 60, picZackmm(1).ScaleWidth, picZackmm(1).ScaleHeight, picZackmov(1).hDC, 0, 0, SRcpaint)
            
            If leftedge > ScaleWidth Then
                leftedge = leftedge - 120
            End If
                
            If current = 1 Then
            current = 0
            End If
                
            If currentm = 1 Then
            currentm = 0
            End If
    
    End Select
    
End Sub

Private Sub Form_Load()
    'Center the window
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    leftedge = Width \ 2
    topedge = Height \ 2
End Sub

