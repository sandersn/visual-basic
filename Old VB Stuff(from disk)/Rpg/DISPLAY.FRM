VERSION 2.00
Begin Form Display 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   3675
   ClientLeft      =   2565
   ClientTop       =   2265
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Height          =   4080
   Left            =   2505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5955
   Top             =   1920
   Width           =   6075
   Begin CommandButton Choice 
      BackColor       =   &H00404000&
      Caption         =   "Choice(1)"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin CommandButton Choice 
      BackColor       =   &H00404000&
      Caption         =   "Choice(0)"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin PictureBox Pict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3875
      Left            =   -60
      ScaleHeight     =   3840
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   -60
      Width           =   1945
   End
   Begin Label Description 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3735
   End
End
Option Explicit
'------------------------------------------------------------
' DISPLAY.FRM - used when the player is presented with a choice.
'------------------------------------------------------------

Sub Choice_Click (Index As Integer)
'------------------------------------------------------------
' Return the user's choice via the global DisplayChoice.
'------------------------------------------------------------
    
    ' return choice to calling routine
    DisplayChoice = Index

    ' remove display form
    Display.Hide
End Sub

