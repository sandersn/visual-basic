VERSION 2.00
Begin Form frmMorg 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Mortgage Calculator"
   ClientHeight    =   6270
   ClientLeft      =   1095
   ClientTop       =   645
   ClientWidth     =   4515
   Height          =   6675
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   4515
   Top             =   300
   Width           =   4635
   Begin CommandButton cmdClear 
      Caption         =   "C&lear"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin TextBox txtEnd 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   420
      Left            =   2640
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin TextBox txtYears 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   420
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin TextBox txtRate 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   420
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin TextBox txtStart 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   420
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ending Balance:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   315
      TabIndex        =   3
      Top             =   4560
      Width           =   1980
   End
   Begin Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Years:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   2520
      Width           =   2115
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Interest Rate:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Starting Balance:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2100
   End
End
Option Explicit
Dim StartBal As Single
Dim Irate As Single
Dim Years As Single
Dim EndBal As Single

Sub cmdCalculate_Click ()
    StartBal = Val(txtStart.Text)
    Irate = Val(txtRate.Text)
    Years = Val(txtYears.Text)
    EndBal = StartBal * (1 + Irate / 36500) ^ (Years * 365)
    txtEnd.Text = Format$(EndBal, "###,###,###0.00")
End Sub

Sub cmdClear_Click ()
    txtStart.Text = ""
    txtRate.Text = ""
    txtYears.Text = ""
End Sub

Sub cmdExit_Click ()
    Dim Msg As String
    Msg = "Come back and compute another mortgage"
    Msg = Msg + " sometime!"
    MsgBox Msg
    End
End Sub

