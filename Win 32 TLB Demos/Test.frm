VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   1560
   End
   Begin VB.PictureBox picHolder 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1440
      Picture         =   "Test.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DirectDraw As IDirectDraw4 ' Or up to IDirectDraw4 if running DirectX 6
Dim DDSurfPrim As IDirectDrawSurface4
Dim DDSurfBack As IDirectDrawSurface4

Private Sub Form_Load()
    Dim ddsd As DDSURFACEDESC
    Dim lResult As Long
    Dim hDC As Long
    lResult = DirectX.DirectDrawCreate(ByVal 0&, DirectDraw, Nothing)   'here we create the actual DirectDraw object that we use with everything else
    'this so nobody but US can draw to the screen and we've got the WHOLE screen
    lResult = DirectDraw.SetCooperativeLevel(frmTest.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN)
    If lResult <> DD_OK Then GoTo Mistake
    lResult = DirectDraw.SetDisplayMode(640, 480, 8)
    If lResult <> DD_OK Then GoTo Mistake
    
    ddsd.dwSize = Len(ddsd) 'for some reason(I think version change) we need to tell the SurfaceDesc its own size....
    ddsd.dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT  'what info do we want to give DD for the new surface?
    'what it's properties and capabilities are
    ddsd.DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
'    ddsd.dwHeight = 100
'    ddsd.dwWidth = 200
    lResult = DirectDraw.CreateSurface(ddsd, DDSurfPrim, Nothing)
    If lResult <> DD_OK Then GoTo Mistake
    'now get a back flipping surface
    ddsd.DDSCAPS.dwCaps = DDSCAPS_BACKBUFFER
    lResult = DDSurfPrim.GetAttachedSurface(ddsd, DDSurfBack)
    If lResult <> DD_OK Then GoTo Mistake
    lResult = DDSurfPrim.GetDC(hDC)
    If lResult <> DD_OK Then GoTo Mistake
    Win32.SetBkColor hDC, RGB(0, 0, 255)
    Win32.SetTextColor hDC, RGB(255, 255, 0)
    Win32.TextOut hDC, 0, 0, "This is a sentence", Len("This is a sentence")
    DDSurfPrim.ReleaseDC (hDC)
    lResult = DDSurfBack.GetDC(hDC)
    If lResult <> DD_OK Then GoTo Mistake
    Win32.SetBkColor hDC, RGB(0, 0, 255)
    Win32.SetTextColor hDC, RGB(255, 255, 0)
    Win32.TextOut hDC, 0, 0, "This is a back msg", Len("This is a back msg")
    DDSurfBack.ReleaseDC (hDC)
    Timer1.Enabled = True
    Exit Sub
Mistake:
        MsgBox "Some error happened!"   'give a stupid err msg
        Form_Unload False   'can't remember what happens with the true/false, so I'll try false--probably won't fail anyway
        '   :)
    End If
    'now let's do something interesting--or at least that works!
    DDSurface
End Sub

Private Sub Form_Paint()
Dim ps As PAINTSTRUCT
Dim rc As RECT
Dim sz As Size
    Win32.BeginPaint frmTest.hWnd, ps
    Win32.GetClientRect frmTest.hWnd, rc
    Win32.GetTextExtentPoint ps.hDC, "This is a sentence", Len("This is a sentence"), sz
    Win32.SetBkColor ps.hDC, RGB(0, 0, 255)
    Win32.SetTextColor ps.hDC, RGB(255, 255, 0)
    Win32.TextOut ps.hDC, (rc.Right - sz.cx) / 2, (rc.Bottom - sz.cy) / 2, "This is a sentence", Len("This is a sentence") - 1
    Win32.EndPaint frmTest.hWnd, ps
End Sub

'this is err handling code for the divide by zero bug that's a return value I think from Dx6 tlb(D3D section)
'    On Error GoTo MySub_Error
    'DirectX code...
'   Exit Sub
'   MySub_Error:
'   If Err.Number = 11 Then Resume ' Divide by zero
'    ' Other errors and error handler here
Private Sub Form_Unload(Cancel As Integer)
    Set DirectDraw = Nothing

End Sub

Private Sub Timer1_Timer()
Dim hDC As Long
Dim lResult As Long
Static bPhase As Boolean
    DDSurfBack.GetDC hDC
    'If lResult <> DD_OK Then GoTo Mistake
    Win32.SetBkColor hDC, RGB(0, 0, 255)
    Win32.SetTextColor hDC, RGB(255, 255, 0)
    If (bPhase = True) Then
        Win32.TextOut hDC, 0, 0, "This is a sentence", Len("This is a sentence")
        bPhase = False
    Else
        Win32.TextOut hDC, 0, 0, "This is a back msg", Len("This is a back msg")
        bPhase = True
    End If
    DDSurfBack.ReleaseDC (hDC)
    
    Do While True
        DDSurfBack.Flip Nothing, 0
'        If lResult = DD_OK Then
'            GoTo Continue
'        ElseIf lResult = DDERR_SURFACELOST Then
'            GoTo Continue
'            lResult = DDSurfPrim.Restore
'            If lResult <> DD_OK Then GoTo Mistake
'        ElseIf lResult <> DDERR_WASSTILLDRAWING Then
'            GoTo Continue
'        End If
Continue:
    Loop
Mistake:
    MsgBox "Some error happened"
    Form_Unload False
End Sub
