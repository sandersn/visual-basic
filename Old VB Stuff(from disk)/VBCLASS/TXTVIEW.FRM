VERSION 2.00
Begin Form frmTextViewer 
   Caption         =   "Text Viewer"
   ClientHeight    =   5130
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7785
   Height          =   5535
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7785
   Top             =   1140
   Width           =   7905
   Begin CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin CommandButton cmdOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin TextBox txtOpen 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin TextBox txtView 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   7575
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Text File to Open:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   1845
   End
End
Option Explicit
Dim Fileno As Single
Dim Opener As String
Dim LineText As String
Dim Crlf As String
Dim Msg As String

Sub cmdExit_Click ()
Dim Msg As String, Crlf As String
Crlf = Chr(10) & Chr(13)
    Msg = "Goodbye, and thankyou for using"
    Msg = Msg & " The Text Viewer"
    Msg = Msg & Crlf & Crlf & Crlf & "(Note:) None of the files"
    Msg = Msg & " you edited actually saved to disk!!!"
    MsgBox Msg
    End
End Sub

Sub cmdOpen_Click ()
Crlf = Chr(13) + Chr(10)
    Fileno = FreeFile
    Opener = txtOpen.Text
    Open Opener For Input As #Fileno
    Do Until EOF(Fileno)
	Input #Fileno, LineText  ' Read line of data.
	Msg = Msg & LineText & Crlf   ' Construct message from data.
    Loop
    Close #Fileno    ' Close file.
    txtView.Text = Msg  ' Display message.
End Sub

