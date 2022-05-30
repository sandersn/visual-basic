VERSION 2.00
Begin Form frmConvert 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Conversion Program"
   ClientHeight    =   2745
   ClientLeft      =   1410
   ClientTop       =   2130
   ClientWidth     =   4260
   Height          =   3150
   Icon            =   FRMCONVE.FRX:0000
   Left            =   1350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4260
   Top             =   1785
   Width           =   4380
   Begin CommandButton cmdExit 
      Caption         =   "E&xit"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin CommandButton cmdClear 
      Caption         =   "&Clear"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin TextBox txtKilometers 
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin TextBox txtMiles 
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin Label lblKilometers 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Kilometers:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   240
      Left            =   765
      TabIndex        =   1
      Top             =   1080
      Width           =   1170
   End
   Begin Label lblMiles 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Miles:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   240
      Left            =   1305
      TabIndex        =   0
      Top             =   480
      Width           =   630
   End
End
Option Explicit

Sub cmdClear_Click ()
    'Resets the Miles and Kilometers text boxes.
    txtMiles.Text = 0
    txtKilometers.Text = 0
End Sub

Sub cmdExit_Click ()
    'Ends the program. (Pretty straightfoward.)
    End
End Sub

Sub Form_Load ()
    'Ah-ha! Here's the extra credit(I think).
    'Basically does the same thing as the Mi. and Km text
    'boxes do, except it does it ahead of time.
    txtMiles.SelStart = 0
    txtMiles.SelLength = 5
End Sub

Sub txtKilometers_LostFocus ()
    'Multiplies the numbers in Km text box and puts the
    'answer in the Mi. text box.
    txtMiles.Text = txtKilometers.Text * .621
    'Tells the comp. where to start and stop.
    txtMiles.SelStart = 0
    txtMiles.SelLength = 5
End Sub

Sub txtMiles_LostFocus ()
    'This part takes the Kilometers and multiplies it to
    'make it into Miles.
    txtKilometers.Text = txtMiles.Text * 1.609
    'Tells the Kilometers text box where to start and how
    'far to go.
    txtKilometers.SelStart = 0
    txtKilometers.SelLength = 5
End Sub

