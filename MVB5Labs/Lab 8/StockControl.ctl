VERSION 5.00
Begin VB.UserControl StockControl 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   PropertyPages   =   "StockControl.ctx":0000
   ScaleHeight     =   1110
   ScaleWidth      =   2565
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   240
   End
   Begin VB.TextBox txtPrice 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   780
      Width           =   1455
   End
   Begin VB.TextBox txtTicker 
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Stock Price"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   780
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Stock Ticker"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "StockControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event TickerKeyPress(KeyAscii As Integer)
Public Property Get Active() As Boolean
Attribute Active.VB_ProcData.VB_Invoke_Property = "Active"
    Active = Timer1.Enabled
End Property
Public Property Let Active(ByVal NewValue As Boolean)
    Timer1.Enabled = NewValue
    PropertyChanged "Active"
End Property
Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont"
    Set Font = txtPrice.Font
End Property
Public Property Let Font(ByVal NewFont As Font)
Dim x As Control
    Set txtPrice.Font = NewFont
    Set txtTicker.Font = NewFont
    Set Label1.Font = NewFont
    Set Label2.Font = NewFont
    For Each x In Controls
        If (TypeOf x Is Label) Or (TypeOf x Is TextBox) Then
            Set x.Font = NewFont
        End If
    Next
    PropertyChanged "Font"
    'why doesn't (me.) let me access the control on the user control?!?
End Property


Private Sub txtTicker_KeyPress(KeyAscii As Integer)
    RaiseEvent TickerKeyPress(KeyAscii)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Active", Timer1.Enabled, True
    PropBag.WriteProperty "Font", txtPrice.Font
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Active = PropBag.ReadProperty("Active", True)
    Font = PropBag.ReadProperty("Font")

End Sub
Public Sub Refresh()
    'sim a timer tick
    Timer1_Timer
End Sub
Private Sub Timer1_Timer()
    If txtTicker = "MSFT" Then
        'return sim stock price.
        txtPrice.Text = Rnd() * 200
    Else
        txtPrice.Text = 0
    End If
End Sub

