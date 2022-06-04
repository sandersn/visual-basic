VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HufNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarFreq As Integer 'local copy
Private mvarString As String
Private mvarLeft As HufNode 'local copy
Private mvarRight As HufNode 'local copy
Private mvarIsLeft As Boolean
Private mvarIsRight As Boolean
Public Property Set Right(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Right = Form1
    Set mvarRight = vData
End Property


Public Property Get Right() As HufNode
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Right
    Set Right = mvarRight
End Property



Public Property Set Left(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Left = Form1
    Set mvarLeft = vData
End Property


Public Property Get Left() As HufNode
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Left
    Set Left = mvarLeft
End Property
Public Property Let IsLeft(ByVal vData As Boolean)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Left = Form1
    mvarIsLeft = vData
End Property


Public Property Get IsLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Left
    IsLeft = mvarIsLeft
End Property
Public Property Let IsRight(ByVal vData As Boolean)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Left = Form1
    mvarIsRight = vData
End Property


Public Property Get IsRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Left
    IsRight = mvarIsRight
End Property



Public Property Let Freq(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Freq = 5
    mvarFreq = vData
End Property


Public Property Get Freq() As Integer
Attribute Freq.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Freq
    Freq = mvarFreq
End Property
Public Property Let Strin(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Strin = "xx"
    mvarString = vData
End Property


Public Property Get Strin() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Freq
    Strin = mvarString
End Property



Private Sub Class_Initialize()
    mvarIsLeft = False
    mvarIsRight = False
End Sub
