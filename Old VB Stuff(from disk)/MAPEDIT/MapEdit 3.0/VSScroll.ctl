VERSION 5.00
Begin VB.UserControl VSScroll 
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   ScaleHeight     =   3615
   ScaleWidth      =   255
   Begin VB.VScrollBar vsbVScroll 
      Height          =   3615
      Left            =   0
      Max             =   15
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "VSScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Change()
Public Event Scroll()
Private bCode As Boolean
Public Property Get Value()
    Value = vsbVScroll.Value
    'again, pretty easy(I'm coding this from the bottom up so some of my comments won't make a whole
    'lot of sense read top 2 bottom.
End Property
Public Property Let Value(intValue)
'this is the programmer changing it--don't tell him about changing it himself!!
    bCode = True
    vsbVScroll.Value = intValue
    'pretty easy.
End Property
Public Property Get Max()
    Max = vsbVScroll.Max
End Property
Public Property Let Max(intMax)
    vsbVScroll.Max = intMax
End Property
Public Property Get Min()
    Min = vsbVScroll.Min
End Property
Public Property Let Min(intMin)
    vsbVScroll.Min = intMin
End Property
Public Property Get LargeChange()
    LargeChange = vsbVScroll.LargeChange
End Property
Public Property Let LargeChange(intLargeChange)
    vsbVScroll.LargeChange = intLargeChange
End Property

Private Sub vsbVScroll_Scroll()
    RaiseEvent Scroll
End Sub
Private Sub vsbVScroll_Change()
    If Not bCode Then 'manipulated by user--throw event.
        RaiseEvent Change
        'now the programmer needs to check the value of the scroll bar or whatever...
    End If
    bCode = False 'remember to reset the variable.
End Sub

Private Sub UserControl_Resize()
    'not too hard 'u'
    vsbVScroll.Height = UserControl.Height
    vsbVScroll.Width = UserControl.Width
End Sub

