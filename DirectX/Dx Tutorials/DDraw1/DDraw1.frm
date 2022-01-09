VERSION 5.00
Begin VB.Form frmDDraw1 
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
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmDDraw1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDX As New DirectX7   'hmm fishy I tink this should be in Form_Load
Dim objDD As DirectDraw7
Dim objDDPrimSurf As DirectDrawSurface7
Dim objDDBackSurf As DirectDrawSurface7
Dim ddClipper As DirectDrawClipper
Dim bInit As Boolean

Private Sub Form_Load()
'whoo this is still like Rube Goldberg compared to standard VB code...
Dim ddsd As DDSURFACEDESC2
    Set objDD = objDX.DirectDrawCreate("")  'why not pass GUID? !know
    'oh. to use the active display driver
    objDD.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
    ddsd.lFlags = DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set objDDPrimSurf = objDD.CreateSurface(ddsd)
    FindMediaDir "lake.bmp"
    
    ddsd.lFlags = DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set objDDBackSurf = objDD.CreateSurfaceFromFile("lake.bmp", ddsd)
    Set ddClipper = objDD.CreateClipper(0)  'apparently no flags :)
    ddClipper.SetHWnd Picture1.hWnd
    objDDPrimSurf.SetClipper ddClipper
    bInit = True
    
End Sub

Private Sub Form_Resize()
Dim rcForm As RECT
Dim ddrval As Long
    Picture1.Width = Me.ScaleWidth
    Picture1.Height = Me.ScaleHeight
    objDX.GetWindowRect Me.hWnd, rcForm
    ddrval = objDDPrimSurf.Blt(rcForm, objDDBackSurf, rcForm, DDBLT_WAIT)
End Sub
