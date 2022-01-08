VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'SINCE I WAS UNABLE TO DIRECTLY CREATE A POINTER WHICH WAS
'ABLE TO POINT TO NULL, I HAD TO CREATE THE TOP
'AS A DUMMY POINTER--IT HOLDS TRASH VALUES AND ITS NEXT
'POINTER HOLDS THE ACTUAL TOP OF THE STACK
'THIS REQUIRED MORE COMPLICATED CODE THAN THE OTHER TWO
'LANGUAGES, WHICH ALLOW PASSING AND COMPARISON TO NULL
'LOCAL VARIABLES
Private mvarTop As Node 'POINTER TO TOP OF STACK

Private Sub Class_Initialize()  'EMPTY(AND ONLY) CONSTRUCTOR
    Set mvarTop = New Node
    mvarTop.ChData = "trash value"
End Sub 'INITIALIZE

Public Function IsEmpty() As Boolean
    IsEmpty = Not mvarTop.IsNext  'IF THERE IS A NEXT (THAT IS, AT LEAST ONE ELEMENT)
                                               'WE'RE NOT EMPTY YET
End Function 'ISEMPTY

Public Function Peek() As String
    Peek = mvarTop.NextLink.ChData  'ACCESS IS THROUGH THE NEXTLINK OF mvarTop
                                                  'DUE TO ROUNDABOUT POINTER CODE
End Function 'PEEK

Public Function Pop() As String
Dim temp As String
    If IsEmpty Then
        MsgBox "Stack Empty!", vbCritical
        Pop = ""
    Else
        temp = mvarTop.NextLink.ChData
        mvarTop.IsNext = mvarTop.NextLink.IsNext
        If mvarTop.IsNext Then    'IF THERE'S AN ELEMENT BEYOND THIS ONE,
            Set mvarTop.NextLink = mvarTop.NextLink.NextLink    'LINK THEM
        Else                                'IF NOT, DELETE THE CURRENT ELEMENT ANYWAY
            Set mvarTop.NextLink = Nothing  'STACK IS NOW EMPTY
        End If  'THERE IS MORE THAN ONE ELEMENT IN STACK
        Pop = temp
    End If
End Function 'POP

Public Sub Push(newData As String)
Dim temp As Node
    If IsEmpty Then 'ADD FIRST ELEMENT TO STACK
        'INITIALIZE A NEW NODE
        Set temp = New Node
        temp.ChData = newData
        temp.IsNext = False
        'NOW LINK THE TOP POINTER TO THE TOP ELEMENT
        'OF THE STACK(CURRENTLY THE ONLY ELEMENT)
        mvarTop.IsNext = True
        Set mvarTop.NextLink = temp
    Else
        'INITIALIZE A NEW NODE
        Set temp = New Node
        temp.IsNext = mvarTop.IsNext
        'LINK NEW NODE TO TOP POINTER AND PREVIOUS TOP ELEMENT
        Set temp.NextLink = mvarTop.NextLink
        temp.ChData = newData
        Set mvarTop.NextLink = temp
    End If  'FIRST ELEMENT
End Sub 'PUSH


