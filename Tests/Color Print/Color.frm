VERSION 5.00
Begin VB.Form frmColor 
   Caption         =   "Color Printing Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Findings: .BackColor is universal. You change it once from a previous change and the whole mess changes.

'This is a test to see if you can feasibly *directly* port thingy's color printing to Windows--it works pretty well, but I
'suspect flicker would be bad. I'd really have to test it out with the whole code--and that means a lot of work...maybe
'I can copy the already written code and just modify it not to use text boxes any more...but then you lose the up/down
'scroll bar advantage...hmm but I *could* leave that around, I just would put it on one edge. OK. Maybe I'll try later if the bug
'fixes don't work.
Private Sub Form_Resize()
Static cPaints As Long
    Me.ForeColor = QBColor(7)
    Me.BackColor = QBColor(0)
    Me.Cls
    Me.Print "Color NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    Me.ForeColor = QBColor(8)
    Me.Print "Gray NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    Me.Print "Gray NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    Me.ForeColor = QBColor(5)
    Me.Print "Gray NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    Me.Print "Gray NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    Me.Print "Gray NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    Me.Print "Gray NunsenseNunsenseNonsenseNinisenseEquis;asdlfj"
    cPaints = cPaints + 1
    Me.Print cPaints
End Sub
