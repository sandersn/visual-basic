VERSION 5.00
Begin VB.Form frmWin32 
   Caption         =   "Win32 Demo"
   ClientHeight    =   975
   ClientLeft      =   5865
   ClientTop       =   5595
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3015
   Begin VB.TextBox txtResults 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Click Here to Report"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmWin32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReport_Click()
    'Win32.Beep 1000, 1009 'note that the numbers here only matter in Windows is set up to use your PC speaker instead of
    'a SB card...Yeah, right..........Oops!! I just remembered that Dad's log-in name is set up that way!!
    '(that's right; his PC speaker actually beeps at him rather than the speakers! Remember,  Truth is Stranger than Fiction.)
    'Note that I need to try that while he is logged on to this computer...The results should be *interesting*.  :() wa ha ha
    'OK: I tried it; very boring results...None!! Not a sound!
    
    'try this now:
    'Win32.MessageBeep Win32.MB_ICONQUESTION
    Win32.MessageBox frmWin32.hWnd, "Demo MsgBox by Win32", "This is a test", _
    Win32.MB_ICONQUESTION Or Win32.MB_YESNOCANCEL
    
    
    'now let's do something really worthwhile with this baby...
    'Win32.changedisplaysettings  'or not; I can't remember anything except that C program I wrote to change the screen res!
    'and IT'S not in here!!!!!!!
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set Win32 = Nothing
End Sub
