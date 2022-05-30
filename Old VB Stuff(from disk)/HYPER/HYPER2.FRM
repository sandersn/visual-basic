VERSION 5.00
Begin VB.Form frmHyper 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "The Hyperator"
   ClientHeight    =   4020
   ClientLeft      =   1920
   ClientTop       =   1530
   ClientWidth     =   3990
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
   Icon            =   "HYPER2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   3990
   Begin VB.TextBox txtHyper 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.HScrollBar hsbHyper 
      Height          =   255
      LargeChange     =   5
      Left            =   600
      Max             =   100
      TabIndex        =   3
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   2520
      Top             =   1560
   End
   Begin VB.TextBox txtOunces 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.VScrollBar vsbOunces 
      Height          =   3735
      LargeChange     =   5
      Left            =   120
      Max             =   250
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgEyes5 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3480
      Picture         =   "HYPER2.frx":030A
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgEyes4 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3480
      Picture         =   "HYPER2.frx":0F8C
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgEyes3 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3480
      Picture         =   "HYPER2.frx":1C0E
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgEyes2 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3480
      Picture         =   "HYPER2.frx":2890
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgEyes1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3480
      Picture         =   "HYPER2.frx":3512
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSoda 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   4
      Left            =   2040
      Picture         =   "HYPER2.frx":4194
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSoda 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   3
      Left            =   2040
      Picture         =   "HYPER2.frx":50B6
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSoda 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   2
      Left            =   1680
      Picture         =   "HYPER2.frx":5FD8
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSoda 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   1680
      Picture         =   "HYPER2.frx":6EFA
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSoda 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   1680
      Picture         =   "HYPER2.frx":7E1C
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHyper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Percent of Total Hyperness:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image imgMainEyes 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   1800
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgMainSoda 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   480
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblOunces 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ounces of soda drunk per day:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmHyper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentSoda As Integer
'Mr. Spellmann and/or anybody else reading this!!
'This is my extra credit program for post lesson 4.
'It started out pretty simpliy, but I got the idea of two
'scroll bars, and it kind gained momentum from there. I
'have just gotten finished trapping the errors like:
'   1. Before, when you went to the text box, unless you
'   swiped the text before you typed, you would have to
'   delete it. This caused an error. I solved it by
'   swiping the text in the GotFocus event. If you STILL
'   try to delete the text, the program will type a 0 for
'   you, and swipe it.
'   2. Also before, if typed in a number too big for the
'   scroll bar to handle, it would stop the program to
'   give you an error message. Now it gives my error
'   message, and sets the scroll bars to their highest
'   capacity.
'I was going to make the eyes bounce iff the walls(from
'hyperness) when the hsbHyper.Value got up to 90. I didn't
'because I ran into WAY too many problems to make it
'feasible.  I also had kind of a hard time with the
'pictures, but they weren't that hard. I think total I've
'spent about 2 hours on this.(Not counting drawing the
'picture.)

Private Sub Form_Load()
    CurrentSoda = 0
End Sub

Private Sub hsbHyper_Change()
     'Make sure the text box shows the value of the scroll
     'bar.
     txtHyper.Text = hsbHyper.Value
     'Change the value of the vsbOunces scroll bar and
     'text box.
     vsbOunces.Value = hsbHyper.Value * 2.5
End Sub

Private Sub hsbHyper_Scroll()
    hsbHyper_Change
End Sub

Private Sub imgMainEyes_Click()
Dim Msg As String
Dim Crlf As String
Crlf = Chr(10) & Chr(13)
    Msg = "This picture shows how hyper you would be if "
    Msg = Msg & Crlf & "you actually drank that much soda!!"
    MsgBox Msg
End Sub

Private Sub imgMainSoda_Click()
    MsgBox "Glug, glug, glug..."
End Sub

Private Sub Timer1_Timer()
    If hsbHyper.Value < 20 Then
        Timer1.Interval = 1500
        imgMainEyes.Picture = imgEyes1.Picture
    End If
    If 20 <= hsbHyper.Value Then
        Timer1.Interval = 1000
        imgMainEyes.Picture = imgEyes2.Picture
    End If
    If 40 <= hsbHyper.Value Then
        Timer1.Interval = 500
        imgMainEyes.Picture = imgEyes3.Picture
    End If
    If 60 <= hsbHyper.Value Then
        Timer1.Interval = 100
        imgMainEyes.Picture = imgEyes4.Picture
    End If
    If 80 <= hsbHyper.Value Then
        Timer1.Interval = 50
        imgMainEyes.Picture = imgEyes5.Picture
    End If
    imgMainSoda.Picture = imgSoda(CurrentSoda).Picture
    CurrentSoda = CurrentSoda + 1
    If CurrentSoda = 5 Then
        CurrentSoda = 0
    End If
End Sub

Private Sub txtHyper_Change()
    'Make sure that the user cannot delete all characters
    'from txtHyper, thereby causing an error.
    If txtHyper.Text = "" Then
        txtHyper.Text = "0"
        txtHyper.SelStart = 0
        txtHyper.SelLength = 1
    End If
    'Make sure that the user cannot enter a value higher
    'that 100 in the txtHyper text box.
    If Val(txtHyper.Text) > 100 Then
        Beep
        txtHyper.Text = 100
        txtHyper.SelStart = 0
        txtHyper.SelLength = 3
    End If
    'Change the current value of hsbHyper to what is in
    'txtHyper.Text.
    hsbHyper.Value = txtHyper.Text
End Sub

Private Sub txtHyper_GotFocus()
    txtHyper.SelStart = 0
    txtHyper.SelLength = 3
End Sub

Private Sub txtOunces_Change()
    'Make sure that the user cannot delete all characters
    'from txtOunces, thereby causing an error.
    If txtOunces.Text = "" Then
        txtOunces.Text = "0"
        txtOunces.SelStart = 0
        txtOunces.SelLength = 1
    End If
    'Make sure that the user cannot enter a value higher
    'that 250 in the txtOunces text box.
    If Val(txtOunces.Text) > 250 Then
        Beep
        txtOunces.Text = 250
        txtOunces.SelStart = 0
        txtOunces.SelLength = 3
    End If
    'Move the value of the scroll bar to the current value
    'of the text box.
    vsbOunces.Value = txtOunces.Text
End Sub

Private Sub txtOunces_GotFocus()
    txtOunces.SelStart = 0
    txtOunces.SelLength = 3
End Sub

Private Sub vsbOunces_Change()
     'Make sure the text box shows the value of the scroll
     'bar.
     txtOunces.Text = vsbOunces.Value
     'Change the value of the Hyperness scroll bar and
     'text box.
     hsbHyper.Value = vsbOunces.Value / 2.5
End Sub

Private Sub vsbOunces_Scroll()
    'Call the vsbOunces_Change procedure.
    vsbOunces_Change
End Sub

