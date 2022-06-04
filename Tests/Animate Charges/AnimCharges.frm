VERSION 5.00
Begin VB.Form frmAnimCharges 
   Caption         =   "Animated Electrical Charges"
   ClientHeight    =   6675
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   9525
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pa&use"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picDraw 
      FillStyle       =   0  'Solid
      Height          =   4500
      Left            =   120
      ScaleHeight     =   200
      ScaleLeft       =   -200
      ScaleMode       =   0  'User
      ScaleTop        =   -100
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   9000
   End
   Begin VB.Timer tmrTick 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   6120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"AnimCharges.frx":0000
      Height          =   780
      Left            =   5160
      TabIndex        =   4
      Top             =   4800
      Width           =   3960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAnimCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'headstone:
'you may have noticed that I like IIf a lot. Well, it's true, it's pretty neat. It's a lot liek the C ? : command which is very useful.
'Well, I have fixed two big bugs, so it looks like this algorithm is ready for prime time.
Option Explicit
Private Const PI = 3.14159
Enum ChargeType
    Positive = 1
    Negative = -1
    Test = 0
End Enum
Private Type Charge
    X As Double
    Y As Double
    dX As Double
    dY As Double
    Strength As Double
    Sign As ChargeType
End Type
Dim charges() As Charge
Dim bPlay As Boolean
Dim cCharges As Long

Private Sub cmdPause_Click()
    bPlay = Not bPlay
    tmrTick.Enabled = Not tmrTick.Enabled
    If bPlay Then
        cmdPause.Caption = "Pa&use"
    Else
        cmdPause.Caption = "Unpa&use"
    End If
End Sub

Private Sub cmdPlay_Click()
    bPlay = True
    tmrTick.Enabled = True
    cmdStop.Enabled = True
    cmdPause.Enabled = True
    cmdPlay.Enabled = False
End Sub

Private Sub cmdStop_Click()
    tmrTick.Enabled = False
    bPlay = False
    Erase charges
    cCharges = 0
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    cmdPlay.Enabled = False
    With picDraw
        'reset magnification
        .ScaleLeft = -200
        .ScaleTop = -100
        .ScaleHeight = 200
        .ScaleWidth = 400
        .Cls
    End With
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim stren
    If bPlay = False Then
        If Button = vbKeyLButton Then
            If Shift = 2 Then   'ctrl
                'do Test charge code here
            Else    'normal
                stren = InputBox$("Enter charge of this point:", "Positive Charge", 0.001)
                If Not IsNumeric(stren) Then Exit Sub
                If cCharges = 0 Then cmdPlay.Enabled = True
                cCharges = cCharges + 1
                ReDim Preserve charges(1 To cCharges)
                charges(cCharges).X = X
                charges(cCharges).Y = Y
                charges(cCharges).Sign = Positive
                charges(cCharges).Strength = Abs(stren)
                picDraw.Circle (charges(cCharges).X, charges(cCharges).Y), 10, RGB(0, 0, 255)
            End If
        ElseIf Button = vbKeyRButton Then
                stren = InputBox$("Enter charge of this point:", "Negative Charge", -0.001)
                If Not IsNumeric(stren) Then Exit Sub
                If cCharges = 0 Then cmdPlay.Enabled = True
                cCharges = cCharges + 1
                ReDim Preserve charges(1 To cCharges)
                charges(cCharges).X = X
                charges(cCharges).Y = Y
                charges(cCharges).Sign = Negative
                charges(cCharges).Strength = -Abs(stren)
                picDraw.Circle (charges(cCharges).X, charges(cCharges).Y), 10, RGB(255, 0, 0)
        End If
    Else 'bPlay = true
        If Button = vbKeyLButton Then
            'zoom out
            With picDraw
                .ScaleLeft = (.ScaleLeft * 2) + X
                .ScaleTop = (.ScaleTop * 2) + Y
                .ScaleHeight = .ScaleHeight * 2
                .ScaleWidth = .ScaleWidth * 2
                .Cls
            End With
        ElseIf Button = vbKeyRButton Then
            'zoom in
            With picDraw
                .ScaleLeft = (.ScaleLeft + X) / 2
                .ScaleTop = (.ScaleTop + Y) / 2
                .ScaleHeight = .ScaleHeight / 2
                .ScaleWidth = .ScaleWidth / 2
                .Cls
            End With
        End If
    End If
End Sub

Private Sub tmrTick_Timer()
Static bDir As Integer
Static grayIntensity As Integer
Dim rad As Single
Dim hyp As Single
Dim F As Single
Dim i As Long, j As Long
ReDim tempcharges(1 To cCharges) As Charge
    For i = 1 To UBound(charges)
        charges(i).dX = 0
        charges(i).dY = 0
        For j = 1 To UBound(charges)
            If i = j Then GoTo continue   'C continue; hack
            If charges(i).X = charges(j).X Then
                rad = IIf(charges(i).Y > charges(j).Y, 3 * PI / 2, PI / 2)
            Else
                rad = Atn((charges(i).Y - charges(j).Y) / (charges(i).X - charges(j).X))
                rad = IIf(charges(i).X > charges(j).X, rad + PI, rad)
            End If
            hyp = Sqr(((charges(i).X - charges(j).X) ^ 2) + ((charges(i).Y - charges(j).Y) ^ 2))
            If hyp < 10 And charges(i).Sign <> charges(j).Sign Then   '0 out the force--i.e. these two charges become neutralized.
                charges(i).Strength = 0
                charges(j).Strength = 0
            End If
            F = ((9000000000#) * charges(i).Strength * charges(j).Strength) / (hyp ^ 2)
            F = -F  'do NOT ask me why we have to flip signs...
            charges(i).dX = charges(i).dX + (Cos(rad) * F)     '* Sgn(charges(i).x - charges(j).x) on the sin/cos
            charges(i).dY = charges(i).dY + (Sin(rad) * F)   'between those last two parens
continue:
        Next j
        tempcharges(i).X = charges(i).X + charges(i).dX
        tempcharges(i).Y = charges(i).Y + charges(i).dY
        'paint charge on form.
        If charges(i).Sign = Negative Then
            picDraw.Circle (tempcharges(i).X, tempcharges(i).Y), 10, RGB(255, 0, 0)
        ElseIf charges(i).Sign = Positive Then
            picDraw.Circle (tempcharges(i).X, tempcharges(i).Y), 10, RGB(0, 0, 255)
        Else    'test charge
            picDraw.Circle (tempcharges(i).X, tempcharges(i).Y), 5, RGB(0, 0, 0)
        End If
    Next i
    
    For i = 1 To UBound(charges)
        charges(i).X = tempcharges(i).X
        charges(i).Y = tempcharges(i).Y
    Next i
    'change bgcolor
    If grayIntensity = 16 Then bDir = -1
    If grayIntensity <= 1 Then bDir = 1
    grayIntensity = grayIntensity + bDir
    picDraw.FillColor = RGB(grayIntensity * 8, grayIntensity * 8, grayIntensity * 8)
End Sub
