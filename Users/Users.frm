VERSION 5.00
Begin VB.Form frmUsers 
   Caption         =   "Change User's Screen Size"
   ClientHeight    =   825
   ClientLeft      =   2895
   ClientTop       =   3315
   ClientWidth     =   3555
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   3555
   Begin VB.CommandButton cmdNathan 
      Caption         =   "&Nathan"
      Height          =   495
      Left            =   1770
      TabIndex        =   1
      Top             =   165
      Width           =   1695
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "&Other"
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   1695
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const DM_BITSPERPEL = &H4000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_DISPLAYFREQUENCY = &H400000
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const DISP_CHANGE_FAILED = -1
Private Const DISP_CHANGE_BADMODE = -2
Private Const DISP_CHANGE_BADFLAGS = -4
Private Const DISP_CHANGE_BADPARAM = -5

Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
Private Declare Function ChangeDisplaySettings Lib "gdi32" (ByRef lpDevMode As DEVMODE, ByVal dwFlags As Integer) As Long

Private Sub cmdNathan_Click()
Dim DMode As DEVMODE
Dim lResult As Long
    DMode.dmFields = DM_BITSPERPEL + DM_PELSWIDTH + DM_PELSHEIGHT + DM_DISPLAYFREQUENCY
    DMode.dmSize = Len(DMode)
    DMode.dmBitsPerPel = 24
    DMode.dmPelsWidth = 1024
    DMode.dmPelsHeight = 768
    DMode.dmDisplayFrequency = 75
    lResult = ChangeDisplaySettings(DMode, 0)
    Select Case lResult
        Case DISP_CHANGE_SUCCESSFUL
            MsgBox "Video Mode Change Successful!"
            'end
        Case DISP_CHANGE_RESTART
            MsgBox "You need to restart Windows to use these settings!"
        Case DISP_CHANGE_BADFLAGS
            MsgBox "The ChangeDisplayMode was passed bad flags!"
        Case DISP_CHANGE_FAILED
            MsgBox "The Video card does not support those settings!"
        Case DISP_CHANGE_BADMODE
            MsgBox "The Video card does not support those settings!"
        Case Else
            MsgBox "Something Wierd happedned!"
    End Select
End Sub

Private Sub cmdOther_Click()
Dim DMode As DEVMODE
Dim lResult As Long
    DMode.dmFields = DM_BITSPERPEL + DM_PELSWIDTH + DM_PELSHEIGHT + DM_DISPLAYFREQUENCY
    DMode.dmSize = Len(DMode)
    DMode.dmBitsPerPel = 24
    DMode.dmPelsWidth = 800
    DMode.dmPelsHeight = 600
    DMode.dmDisplayFrequency = 72
    lResult = ChangeDisplaySettings(DMode, 0)
    Select Case lResult
        Case DISP_CHANGE_SUCCESSFUL
            MsgBox "Video Mode Change Successful!"
            'end
        Case DISP_CHANGE_RESTART
            MsgBox "You need to restart Windows to use these settings!"
        Case DISP_CHANGE_BADFLAGS
            MsgBox "The ChangeDisplayMode was passed bad flags!"
        Case DISP_CHANGE_FAILED
            MsgBox "The Video card does not support those settings!"
        Case DISP_CHANGE_BADMODE
            MsgBox "The Video card does not support those settings!"
        Case Else
            MsgBox "Something Wierd happedned!"
    End Select

End Sub
