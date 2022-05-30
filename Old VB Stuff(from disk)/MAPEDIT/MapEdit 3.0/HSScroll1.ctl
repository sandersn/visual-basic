VERSION 5.00
Begin VB.UserControl HSScroll 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LockControls    =   -1  'True
   ScaleHeight     =   255
   ScaleWidth      =   3615
   Begin VB.HScrollBar hsbHScroll 
      Height          =   255
      Left            =   0
      Max             =   15
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblToolTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ToolTipText"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "HSScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Change()
Public Event Scroll()
Private bCode As Boolean
Private strToolTip As String
Public Property Get Value()
    Value = hsbHScroll.Value
    'again, pretty easy(I'm coding this from the bottom up so some of my comments won't make a whole
    'lot of sense read top 2 bottom.
End Property
Public Property Let Value(intValue)
'this is the programmer changing it--don't tell him about changing it himself!!
    bCode = True
    hsbHScroll.Value = intValue
    'pretty easy.
End Property
Public Property Get Max()
    Max = hsbHScroll.Max
End Property
Public Property Let Max(intMax)
    hsbHScroll.Max = intMax
End Property
Public Property Get Min()
    Min = hsbHScroll.Min
End Property
Public Property Let Min(intMin)
    hsbHScroll.Min = intMin
End Property
Public Property Get LargeChange()
    LargeChange = hsbHScroll.LargeChange
End Property
Public Property Let LargeChange(intLargeChange)
    hsbHScroll.LargeChange = intLargeChange
End Property

Private Sub hsbHScroll_Scroll()
    RaiseEvent Scroll
End Sub
Private Sub hsbHScroll_Change()
    If Not bCode Then 'manipulated by user--throw event.
        RaiseEvent Change
        'now the programmer needs to check the value of the scroll bar or whatever...
    End If
    bCode = False 'remember to reset the variable.
End Sub

Private Sub UserControl_Resize()
    'not too hard 'u'
    hsbHScroll.Height = UserControl.Height
    hsbHScroll.Width = UserControl.Width
End Sub
