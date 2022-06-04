VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Async"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim CurrForm As Form
Public Event Complete(Result As String)
Public Sub AsyncMethod()
    Set CurrForm = New frmTimer
    Set CurrForm.Callback = Me
    CurrForm.Timer1.Enabled = True
End Sub
Friend Sub DoWork()
Dim sngEndTime As Single
    Set CurrForm = Nothing
    sngEndTime = Timer + 5
    Do While Timer < sngEndTime
        DoEvents    'do nothing--and tell Windows™ that that is what we're doing.
    Loop
    RaiseEvent Complete("Hello, world!")
End Sub
