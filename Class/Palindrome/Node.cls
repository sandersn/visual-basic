VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The basic node building block for a linked list"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'REQUIRES ALL VARIABLES TO BE DECLARED BEFORE USE

'LOCAL VARIABLES
'VB'S CLASS NOTATION PREFIX OF mvar- AVOID FREQUENT NAMING CONFLICTS
'SINCE VB DOES NOT SUPPORT PACKAGES OR NAMESPACES
'm=member
'var=variable
Private mvarNext As Node 'THE NEXT POINTER
Private mvarIsNext As Boolean   'THE BOOLEAN HELPER WHICH INDICATES WHETHER OR NOT
                                            'THERE IS A NEXT POINTED TO OR NOT
Private mvarChData As String    'THE CHARACTER DATA

Private Sub Class_Initialize()
'Class_Initialize IS THE VB VERSION OF A CONSTRUCTOR.
'IT CANNOT BE OVERLOADED AND TAKES NO ARGUMENTS
'SINCE VB DOES NOT SUPPORT OVERLOADING
    Set mvarNext = Nothing
    mvarIsNext = False
End Sub 'INITIALIZE

Public Property Set NextLink(newNext As Object)
'Property Set IS CALLED WHENEVER AN OBJECT'S ADDRESS IS ASSIGNED VIA Set
'EX: Set Node.Next = New Node
    Set mvarNext = newNext    'Set IS REQUIRED WHENEVER POINTER ASSIGNMENT TAKES PLACE
End Property
Public Property Get NextLink() As Node
'Property Get IS CALLED WHENEVER AN OBJECT'S ADDRESS IS REQUESTED
'EX: Print Node.Next
    Set NextLink = mvarNext
End Property


Public Property Let IsNext(newIsNext As Boolean)
'Property Let IS THE PRIMITIVE COUNTERPART OF Property Get
    mvarIsNext = newIsNext
End Property
Public Property Get IsNext() As Boolean
    IsNext = mvarIsNext
End Property


Public Property Let ChData(ByVal newData As String)
'ChData IS THE ACTUAL DATA PORTION OF THE NODE
    mvarChData = newData
End Property
Public Property Get ChData() As String
Attribute ChData.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChData
    ChData = mvarChData
End Property

