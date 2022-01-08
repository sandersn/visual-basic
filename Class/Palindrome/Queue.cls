VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The queue class that uses a linked list for implementation"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'SINCE I WAS UNABLE TO DIRECTLY CREATE A POINTER WHICH WAS
'ABLE TO POINT TO NULL, I HAD TO CREATE THE FRONT AND REAR
'AS DUMMY POINTERS--THEY HOLD TRASH VALUES AND THEIR NEXT
'POINTER HOLDS THE ACTUAL FRONT/REAR OF THE QUEUE
'THIS REQUIRED MORE COMPLICATED CODE THAN THE OTHER TWO
'LANGUAGES, WHICH ALLOW PASSING AND COMPARISON TO NULL
'LOCAL VARIABLES
Private mvarFront As Node 'THE FRONT POINTER
Private mvarRear As Node  'THE REAR POINTER

Private Sub Class_Initialize()
Set mvarFront = New Node
Set mvarRear = New Node
    mvarFront.ChData = "trash value front"  'THIS WILL BE AN OBVIOUS SIGN
    mvarRear.ChData = "trash value rear"    'IF THERE IS A BUG IN THE QUEUE LOGIC
End Sub 'INITIALIZE

Public Function Front() As String
    Front = mvarFront.NextLink.ChData 'ACCESS IS THROUGH THE NEXTLINK OF mvarTop
                                                    'DUE TO ROUNDABOUT POINTER CODE
End Function 'FRONT

Public Function Remove() As String
Dim temp As String
    If IsEmpty Then
        MsgBox "Queue Empty!", vbCritical
        Remove = ""
    Else
        Remove = mvarFront.NextLink.ChData
        mvarFront.IsNext = mvarFront.NextLink.IsNext
        If mvarFront.IsNext Then
            Set mvarFront.NextLink = mvarFront.NextLink.NextLink
        Else    'that was the last element, so now we have an empty queue
            mvarFront.IsNext = False
            mvarRear.IsNext = False
            Set mvarFront.NextLink = Nothing
            Set mvarRear.NextLink = Nothing
        End If
    End If 'QUEUE EMPTY
End Function 'REMOVE

Public Sub Insert(newData As String)
Dim temp As Node
    If IsEmpty Then 'FIRST ELEMENT IN QUEUE
        'INITIALIZE A NEW NODE
        Set temp = New Node
        temp.ChData = newData
        temp.IsNext = False
        'NOW LINK THE FRONT/REAR POINTERS TO THIS
        'ELEMENT SINCE IT IS THE ONLY ONE IN THE QUEUE
        mvarRear.IsNext = True
        mvarFront.IsNext = True
        Set mvarRear.NextLink = temp
        Set mvarFront.NextLink = temp
    Else
        'INITIALIZE A NEW NODE
        Set temp = New Node
        temp.IsNext = False
        temp.ChData = newData
        'AND LINK THE FORMER REAR TO IT
        Set temp.NextLink = Nothing
        mvarRear.NextLink.IsNext = True
        Set mvarRear.NextLink.NextLink = temp
        'POINT THE REAR POINTER TO THE NEW REAR
        Set mvarRear.NextLink = temp
    End If  'FIRST ELEMENT
End Sub 'INSERT

Public Function IsEmpty() As Boolean
    IsEmpty = Not mvarFront.IsNext  'IF THERE IS A FRONT ELEMENT, WE'RE NOT EMPTY YET
End Function 'ISEMPTY

