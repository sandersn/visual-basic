VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreditCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public CardNumber As Integer    'Public, so show up as properties...
Attribute CardNumber.VB_VarUserMemId = 0
Public ExpireDate As Date
Private lngPurchaseAmount As Long   'private, so THIS IS NOT THE PROPERTY! It is a holding place for the property.
Public Event Status(ByVal strName As String)
Property Get PurchaseAmount() As Long 'note:PurchaseAmount must be the same for both variable name and function name
Attribute PurchaseAmount.VB_Description = "The amount of credit to be put on the credit card."
Attribute PurchaseAmount.VB_HelpID = 2
    PurchaseAmount = lngPurchaseAmount
End Property
Property Let PurchaseAmount(lngAmt As Long)
    If lngAmt <= 0 Then
        Err.Raise vbObjectError + 1000, , "Purchase amount must be greater than zero."
        'lngPurchaseAmount = 0
    Else
        lngPurchaseAmount = lngAmt
    End If
End Property
Public Function Approve() As Boolean
Attribute Approve.VB_Description = "Approve the current transaction.(after Credit Card object's fields are filled) "
Attribute Approve.VB_HelpID = 1
Dim sngEndTime As Single
    RaiseEvent Status("Dialing bank...")
    sngEndTime = Timer + 2
    Do While Timer < sngEndTime
        DoEvents
    Loop
    If lngPurchaseAmount < 1000 And ExpireDate > Now() Then
        Approve = True
    ElseIf lngPurchaseAmount >= 1000 Or ExpireDate <= Now() Then
        Approve = False
    End If
    RaiseEvent Status("Processing card...")
        sngEndTime = Timer + 2
    Do While Timer < sngEndTime
        DoEvents
    Loop

End Function
