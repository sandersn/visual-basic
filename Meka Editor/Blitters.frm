VERSION 5.00
Begin VB.Form frmBlitters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meka Blit Mode Editor"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   885
   ClientWidth     =   8085
   Icon            =   "Blitters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWizard 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   8100
      ScaleHeight     =   5895
      ScaleWidth      =   1695
      TabIndex        =   22
      Top             =   0
      Width           =   1695
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Resolution"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   195
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blitter"
         Height          =   195
         Left            =   440
         TabIndex        =   26
         Top             =   1100
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VSync"
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   2235
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flip"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   3165
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show In GUI"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   915
      End
      Begin VB.Image imgShowinGUI 
         Height          =   1560
         Left            =   50
         Top             =   4080
         Width           =   1605
      End
      Begin VB.Image imgFlip 
         Height          =   375
         Left            =   720
         Top             =   3360
         Width           =   615
      End
      Begin VB.Image imgFlipAndVSync 
         Height          =   120
         Left            =   50
         Picture         =   "Blitters.frx":0442
         Top             =   3045
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image imgVSync 
         Height          =   630
         Left            =   120
         Top             =   2400
         Width           =   1245
      End
      Begin VB.Image imgBlitter 
         Height          =   900
         Left            =   285
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Shape shpGG 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   780
         Top             =   645
         Width           =   255
      End
      Begin VB.Shape shpSMS 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   600
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape shpScreen 
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   480
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdShowWizard 
      Caption         =   ">> Mo&re"
      Height          =   975
      Left            =   7395
      TabIndex        =   21
      Top             =   4920
      Width           =   660
   End
   Begin VB.CommandButton cmdUp 
      Height          =   855
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Blitters.frx":0984
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdDown 
      Height          =   855
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Blitters.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1875
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Mode"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtComments 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Text            =   "Blitters.frx":0D98
      Top             =   4920
      Width           =   7215
   End
   Begin VB.CheckBox chkVSync 
      Caption         =   "&VSync"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.CheckBox chkFlip 
      Caption         =   "&Flip"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   4560
      Width           =   765
   End
   Begin VB.CheckBox chkShowInGUI 
      Caption         =   "Show In &GUI"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   4560
      Width           =   1245
   End
   Begin VB.ComboBox cboYRes 
      Height          =   1515
      IntegralHeight  =   0   'False
      Left            =   1320
      Style           =   1  'Simple Combo
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox cboXRes 
      Height          =   1515
      IntegralHeight  =   0   'False
      ItemData        =   "Blitters.frx":0DAC
      Left            =   120
      List            =   "Blitters.frx":0DAE
      Style           =   1  'Simple Combo
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Meka.BLT"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdNewMode 
      Caption         =   "&New Mode"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstModes 
      Height          =   2010
      ItemData        =   "Blitters.frx":0DB0
      Left            =   120
      List            =   "Blitters.frx":0DB2
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.ListBox lstDrivers 
      Height          =   1500
      IntegralHeight  =   0   'False
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ListBox lstBlitters 
      Height          =   2010
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image imgGUIs 
      Height          =   1560
      Index           =   1
      Left            =   2520
      Picture         =   "Blitters.frx":0DB4
      Top             =   0
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   0
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      Caption         =   "&Creator's comments:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4600
      Width           =   1425
   End
   Begin VB.Label lblHelp 
      Caption         =   "To change a mode name, double click it."
      Height          =   4695
      Left            =   4680
      TabIndex        =   20
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblDrivers 
      Caption         =   "Dri&vers:"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblYRes 
      AutoSize        =   -1  'True
      Caption         =   "&Y Resolution:"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblXRes 
      AutoSize        =   -1  'True
      Caption         =   "&X Resolution:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblBlitters 
      Caption         =   "&Blitters:"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblModes 
      Caption         =   "&Modes:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image imgBlitters 
      Height          =   900
      Index           =   6
      Left            =   7200
      Picture         =   "Blitters.frx":9196
      Top             =   3720
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgBlitters 
      Height          =   900
      Index           =   1
      Left            =   7080
      Picture         =   "Blitters.frx":CBF8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgBlitters 
      Height          =   900
      Index           =   4
      Left            =   7440
      Picture         =   "Blitters.frx":1065A
      Top             =   3000
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgBlitters 
      Height          =   900
      Index           =   3
      Left            =   7320
      Picture         =   "Blitters.frx":140BC
      Top             =   3240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgBlitters 
      Height          =   450
      Index           =   0
      Left            =   8040
      Picture         =   "Blitters.frx":17B1E
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlitters 
      Height          =   900
      Index           =   2
      Left            =   4320
      Picture         =   "Blitters.frx":189E8
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgBlitters 
      Height          =   450
      Index           =   5
      Left            =   4440
      Picture         =   "Blitters.frx":1C44A
      Top             =   840
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image imgVSyncs 
      Height          =   630
      Index           =   1
      Left            =   4320
      Picture         =   "Blitters.frx":1E19C
      Top             =   2520
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image imgVSyncs 
      Height          =   630
      Index           =   0
      Left            =   4320
      Picture         =   "Blitters.frx":20B36
      Top             =   1920
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image imgFlips 
      Height          =   360
      Index           =   1
      Left            =   4680
      Picture         =   "Blitters.frx":234D0
      Top             =   3360
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Image imgFlips 
      Height          =   360
      Index           =   0
      Left            =   4680
      Picture         =   "Blitters.frx":24592
      Top             =   3840
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgGUIs 
      Height          =   1560
      Index           =   0
      Left            =   960
      Picture         =   "Blitters.frx":24DB4
      Top             =   0
      Visible         =   0   'False
      Width           =   1605
   End
End
Attribute VB_Name = "frmBlitters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Started this project as a result of downloading Meka 0.53c. Noted new blitter caps. Ran doubled(already done
'once). Liked it very much. Enabled double Game gear *manually*. Played SF2G in double game gear. LOVED it.
'Spent 15 minutes enabling new colour schemes using Wondertheme. Decided it would be cool to have a blitter editor.
'Decided it would pretty easy to implement. Got fired up about it. In a moment of thoughtfulness thought maybe Zoop
'would decide it was worth a free registered copy of Meka. So there you have it ^_^
'Special note: Non-commented lines are ignored when parsing meka.blt. That means they won't bug anything up, but
'they won't show up in the comments box, either.
'This program requires the Visual Basic 5.0(or greater) Runtimes. But if you can run Wonder Theme, you can run this.
'If you don't have them, go to download.com and search for Visual Basic Runtimes or something similar.
'Note to self: Note to Zoop: Add caps to Meka to allow the blitters to be 1.used in GUI mode 2.displayed on screenshots.
'I had to eagle the picture of Lizard Man manually!
'TODO:
                'Add save. DONE
                'Add New Mode caps.DONE
                'Add error checking when loading. DONE
                'Add ability to choose Meka.BLT file. DONE
                'Add saving of comments and display. DONE
                'Add roll over help.DONE
                'Change Tab order again.DONE
                'Add listing of used X, Y resolutions for the timid and stupid. DONE
                'Add syncing of X,Y combo boxes. DONE
                'Add Delete with confirmation. DONE
                'Add Up/Down buttons. DONE plus KB interface as well.
                'Add comments label. DONE
                'Add rollover for Resolution combo. DONE with a background imagebox that just fakes you out becuase
'it's big enough that you almost ALWAYS get a mousemove for it before going over the actual combo boxes.
'--------Version: 1.1
                'Add command line ability to choose Meka Dir. DONE
                'Add 'AppWizard' caps with a picture box containing other images
                'Add a little more mouse-over help(or improve already existing help).

'Fixes:
                'Fixed Saving to Mekablt.txt to Meka.blt. Oops
                'Fixed Saving from App.Path to MekaPath(a string var)
'-------new ver 1.1
                'Fixed bug where if you specified the wrong path, MekaBlt got confused and never got the new directory you specified.
Option Explicit
Public Enum Blitters
    Normal = 0
    DoubleSize = 1
    Scanlines = 2
    TVMode = 3
    Eagle = 4
    Parallel = 5
    TVMode_Double = 6
End Enum
Public Enum Drivers
    Auto = 0
    Safe = 1
    VGA = 2
    ModeX = 3
    Vesa1 = 4
    Vesa2b = 5  '??!?!?
    Vesa2l = 6  'ditto
    Vesa3 = 7   'even more ditto
    VBEAF = 8
End Enum
Private Type BlitModes
    Name As String
    XRes As Integer
    YRes As Integer
    Driver As Drivers
    Blitter As Blitters
    Flip As Boolean
    VSync As Boolean
    Comments As String
    ShowInGUI As Boolean    'only non-file inherent property; determined by whether or not the Name is commented.
End Type
Private Blits() As BlitModes
Private BlitterList(0 To 6) As String
Private DriverList(0 To 8) As String
Private cModes As Integer   'sorry only 32767 modes ^^
Private ZoopComments As String
Private MekaPath As String

Private Sub cboXRes_Change()
    If IsNumeric(cboXRes.Text) And Len(cboXRes.Text) < 5 Then
        Blits(lstModes.ListIndex).XRes = cboXRes.Text
        UpdateWizard
    End If
End Sub

Private Sub cboXRes_Click()
    Blits(lstModes.ListIndex).XRes = cboXRes.Text
    cboXRes_GotFocus
End Sub

Private Sub cboXRes_GotFocus()
lblHelp.Caption = "Set the horizontal resolution. Be careful, however: not all numbers will work. If you don't know which numbers to use, see what previously created video modes have used. Note also that for some computers the colors in 512x384 resolution are messed up."
End Sub

Private Sub cboYRes_Change()
    If IsNumeric(cboYRes.Text) And Len(cboYRes.Text) < 5 Then
        Blits(lstModes.ListIndex).YRes = cboYRes.Text
        UpdateWizard
    End If

End Sub

Private Sub cboYRes_Click()
    Blits(lstModes.ListIndex).YRes = cboYRes.Text
    cboYRes_GotFocus
End Sub

Private Sub cboYRes_GotFocus()
lblHelp.Caption = "Set the vertical resolution. Be careful, however: not all numbers will work. If you don't know which numbers to use, see what previously created video modes have used. Note also that for some computers the colors in 512x384 resolution are messed up."
End Sub

Private Sub chkFlip_Click()
    Blits(lstModes.ListIndex).Flip = IIf(chkFlip.Value = vbChecked, True, False)
    'if check, turn mode on. else not. DUUUUUUH -ahem-
    UpdateWizard
End Sub

Private Sub chkFlip_GotFocus()
    chkFlip_MouseMove 0, 0, 1, 1
End Sub

Private Sub chkFlip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "When Flip is on, Meka creates a back buffer of memory on which it draws the screen. This reduces the possibility of 'tearing' or flickering. But it slows down emulation speed. Meka's GUI uses flip by default. Note that enabling both Flip and VSync will cause speed to be cut in half."
End Sub

Private Sub chkShowInGUI_Click()
    Blits(lstModes.ListIndex).ShowInGUI = IIf(chkShowInGUI.Value = vbChecked, True, False)
    UpdateWizard
End Sub

Private Sub chkShowInGUI_GotFocus()
chkShowInGUI_MouseMove 0, 0, 1, 1
End Sub

Private Sub chkShowInGUI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Checked modes appear in Meka's menu. Unchecked modes sit in Meka.BLT providing amusement for the users of this program." & vbCrLf & vbCrLf & "Meka will display only a maximum of 20 video modes, even if more are checked."
End Sub

Private Sub chkVSync_Click()
    Blits(lstModes.ListIndex).VSync = IIf(chkVSync.Value = vbChecked, True, False)
    UpdateWizard
End Sub

Private Sub chkVSync_GotFocus()
chkVSync_MouseMove 0, 0, 1, 1
End Sub

Private Sub chkVSync_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "When VSync is on, Meka waits until the electron gun is not 'shooting' before it draws the screen. This reduces the possibility of 'tearing' or flickering. But it slows down emulation speed. Note that enabling both Flip and VSync will cause speed to be cut in half."
End Sub

Private Sub cmdDelete_Click()
'this sub is ugly and slow, but it works
Dim i As Integer, j As Integer
'Delete, but ask for confimatio first.
'Confirmatio demando. Comprenis quis faris?
    If vbYes = MsgBox("Are you sure you want to delete " & Blits(lstModes.ListIndex).Name & "?", vbYesNo + vbQuestion, "Delete Blit Mode") Then
        'create temp array same size as Blits and copy
Dim tempBlits() As BlitModes
        ReDim tempBlits(0 To cModes - 1)
        For i = 0 To cModes Step 1
            If i <> lstModes.ListIndex Then
                tempBlits(j) = Blits(i)
                j = j + 1
            End If
        Next i
        'disappear the mode
        cModes = cModes - 1
        lstModes.RemoveItem (lstModes.ListIndex)    'aw boohoo
        lstModes.ListIndex = 0  'set to first item. tough for usability ^_^
        'copy back
        ReDim Blits(0 To cModes)
        For i = 0 To cModes Step 1
            Blits(i) = tempBlits(i)
        Next i
    End If
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "This button deletes the blit mode currently selected. If you decide that you didn't want to delete a mode after all, exit the program without saving. If you have already saved, you can still restore Meka.blt from the Meka zip file, but any changes you have made will be lost."
End Sub

Private Sub cmdDown_Click()
Dim beginIndex As Integer
Dim tempBlit As BlitModes
    With lstModes
    If .ListIndex < .ListCount - 1 Then
        beginIndex = .ListIndex 'save info we'll need at the end
        tempBlit = Blits(beginIndex + 1)
        lstModes.AddItem Blits(beginIndex).Name, beginIndex + 2
        lstModes.RemoveItem beginIndex
        Blits(beginIndex + 1) = Blits(beginIndex)
        Blits(beginIndex) = tempBlit
        .ListIndex = beginIndex + 1
    End If
    End With
End Sub

Private Sub cmdDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Click to change the position of the selected mode in this program, Meka.BLT, and Meka itself."
End Sub

Private Sub cmdNewMode_Click()
Dim newName As String
    
    With lstModes
    newName = InputBox("Enter new name for mode", "New Name", Blits(.ListIndex).Name)
    If newName <> "" Then   'make sure no blank
        cModes = cModes + 1
        If cModes Mod 100 = 0 Then  'alloc more memory if necessary
            ReDim Preserve Blits(0 To cModes + 100)
        End If
        Blits(cModes).Name = newName    'add to data
        .AddItem Blits(cModes).Name
        .ListIndex = cModes 'setfocus for convenience
        'now init the values to their default
        Blits(cModes).XRes = 320
        cboXRes.Text = 320
        Blits(cModes).YRes = 200
        cboYRes.Text = 200
        Blits(cModes).Blitter = Normal
        lstBlitters.ListIndex = 0
        Blits(cModes).Driver = Auto
        lstDrivers.ListIndex = 0
        Blits(cModes).ShowInGUI = False
        chkShowInGUI.Value = False
        Blits(cModes).Flip = False
        chkFlip.Value = False
        Blits(cModes).VSync = False
        chkVSync.Value = False
    End If
    End With
    
End Sub

Private Sub cmdNewMode_GotFocus()
cmdNewMode_MouseMove 0, 0, 1, 1
End Sub

Private Sub cmdNewMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Create a new mode. You must provide a name and customize it." & vbCrLf & vbCrLf & "Meka will display only a maximum of 20 video modes, even if more are checked."
End Sub

Private Sub cmdSave_Click()
Dim intFileno As Integer
Dim i As Integer
Dim intStart As Integer
Dim intEnd As Integer
Dim strBuff As String
Dim bDone As Boolean
    intFileno = FreeFile
    Open MekaPath & "\" & "Meka.blt" For Output As #intFileno
    'first put in Zoop's comments
    Print #intFileno, ZoopComments
    'now loop through all modes, writing as we go--hoho
    For i = 0 To cModes
        'name
        strBuff = "[" & Blits(i).Name & "]"
        strBuff = CommentOut(strBuff, i)
        Print #intFileno, strBuff
        'resolution
        strBuff = "res = " & Trim$(Str$(Blits(i).XRes)) & "x" & Trim$(Str$(Blits(i).YRes))
        strBuff = CommentOut(strBuff, i)
        Print #intFileno, strBuff
        'blitter
        strBuff = "blitter = " & BlitterList(Blits(i).Blitter)
        strBuff = CommentOut(strBuff, i)
        Print #intFileno, strBuff
        'driver
        strBuff = "driver = " & DriverList(Blits(i).Driver)
        strBuff = CommentOut(strBuff, i)
        Print #intFileno, strBuff
        'options
        If Blits(i).VSync = True Then
            strBuff = "vsync"
            strBuff = CommentOut(strBuff, i)
            Print #intFileno, strBuff
        End If
        If Blits(i).Flip = True Then
            strBuff = "flip"
            strBuff = CommentOut(strBuff, i)
            Print #intFileno, strBuff
        End If
        'comments--this is where things get very chancey.
        If Len(Blits(i).Comments) = 0 Then
            Print #intFileno, ""
        Else
            intEnd = 1 'start at 1
            bDone = False 'and [re]set
            Do
                intStart = InStr(intEnd, Blits(i).Comments, vbCrLf) 'find end of each line
                If intStart = 0 Then
                    intStart = Len(Blits(i).Comments) + 1
                    bDone = True
                End If
                strBuff = Mid$(Blits(i).Comments, intEnd, intStart - intEnd)
                strBuff = CommentOut(strBuff, i)
                Print #intFileno, strBuff
                intEnd = intStart + 2  'account for vbCrlf
            Loop Until bDone = True
            Print #intFileno, ""
        End If
    Next i
    Close intFileno
End Sub

Private Sub cmdSave_GotFocus()
cmdSave_MouseMove 0, 0, 1, 1
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Click to save Meka.BLT. Only do this if you are sure that you like your changes."
End Sub

Private Sub cmdShowWizard_Click()
Static bWizardMode As Boolean
    bWizardMode = Not bWizardMode
    If bWizardMode = True Then
        cmdShowWizard.Caption = "<< &Less"
        frmBlitters.Width = picWizard.Left + picWizard.Width + 100
        UpdateWizard
    Else
        cmdShowWizard.Caption = ">> Mo&re"
        frmBlitters.Width = lblHelp.Left + lblHelp.Width + 100
    End If
End Sub

Private Sub cmdShowWizard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Show a detailed picture which gives a visual representation of all the options in Meka.BLT"
End Sub

Private Sub cmdUp_Click()
Dim beginIndex As Integer
Dim tempBlit As BlitModes
    With lstModes
    If .ListIndex > 0 Then
        beginIndex = .ListIndex 'save info we'll need at the end
        tempBlit = Blits(beginIndex - 1)
        lstModes.AddItem Blits(beginIndex).Name, beginIndex - 1
        lstModes.RemoveItem beginIndex + 1
        Blits(beginIndex - 1) = Blits(beginIndex)
        Blits(beginIndex) = tempBlit
         .ListIndex = beginIndex - 1
    End If
    End With

End Sub

Private Sub cmdUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Click to change the position of the selected mode in this program, Meka.BLT, and Meka itself."
End Sub

Private Sub Form_Load()
    'well, start loading Meka.Blt using Input mode oboyoboy
    On Error GoTo FileNotFoundErr
Dim intFileno As Integer
Dim strBuff As String
Dim strTempType As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim intStart As Integer
Dim intEnd As Integer
ReDim Blits(0 To 100)
    'first, init string-lists of blitters and drivers
    BlitterList(0) = "normal"
    BlitterList(1) = "double"
    BlitterList(2) = "scanlines"
    BlitterList(3) = "tvmode"
    BlitterList(4) = "eagle"
    BlitterList(5) = "parallel"
    BlitterList(6) = "tvmode_double"
    '---------------------
    DriverList(0) = "auto"
    DriverList(1) = "safe"
    DriverList(2) = "vga"
    DriverList(3) = "modex"
    DriverList(4) = "vesa1"
    DriverList(5) = "vesa2b"
    DriverList(6) = "vesa2l"
    DriverList(7) = "vesa3"
    DriverList(8) = "vbeaf"
'load meka.blt and then unload the choosedir dialog
    intFileno = FreeFile
    MekaPath = frmChooseDir.Tag
    Open MekaPath & "\" & "Meka.blt" For Input As #intFileno
    Unload frmChooseDir
    Do  'skip past Zoop's comments
        Line Input #intFileno, strBuff
        ZoopComments = ZoopComments & strBuff & vbCrLf 'save comments for writing back later.
    Loop Until strBuff = ""
        i = 0
    Do Until EOF(intFileno)
    'loop through each blit mode and save
        Do  'get name first
            Do While strBuff = ""
                Line Input #intFileno, strBuff
            Loop
            If Left$(strBuff, 1) = ";" Then 'not ShowInGUI
                Blits(i).ShowInGUI = False
                strBuff = Right$(strBuff, Len(strBuff) - 1)
            Else
                Blits(i).ShowInGUI = True
            End If
            If Left$(Trim$(strBuff), 1) <> ";" Then 'make sure !comment
                intStart = InStr(1, strBuff, "[")
                intEnd = InStr(intStart, strBuff, "]")
                If intStart = 0 Or intEnd = 0 Then  'bad title format
                    MsgBox "Bad title format! Title is supposed to be enclosed in brackets. Using current line as title anyway."
                    Blits(i).Name = strBuff
                Else
                    Blits(i).Name = Mid$(strBuff, intStart + 1, intEnd - intStart - 1)
                End If
            Else
                Blits(i).Comments = Blits(i).Comments & strBuff & vbCrLf
            End If
        Loop While Blits(i).Name = ""
        Do  'get each line of the blit
            Line Input #intFileno, strBuff
            If Blits(i).ShowInGUI = False And strBuff <> "" Then
                strBuff = Right$(strBuff, Len(strBuff) - 1) 'clip the comment(i.e. ';') so we can parse it
            End If
            'read next line.
            If Left$(Trim$(strBuff), 3) = "res" Then 'resolution line--parse 2 pieces
                 intStart = InStr(3, strBuff, "=")  'start and end of XRes
                 intEnd = InStr(intStart, strBuff, "x")
                 If intStart = 0 Or intEnd = 0 Then  'bad title format
                    MsgBox "Bad resolution definition! Title is supposed to be enclosed in brackets. Using default resolution of 320x200."
                    Blits(i).XRes = 320
                    Blits(i).YRes = 200
                Else
                    If Not IsNumeric(Trim(Mid$(strBuff, intStart + 1, intEnd - intStart - 1))) Then
                        MsgBox "Resolution width not a number! Using default width of 320"
                        Blits(i).XRes = 320
                    Else
                        Blits(i).XRes = Trim(Mid$(strBuff, intStart + 1, intEnd - intStart - 1))
                    End If
                    If Not IsNumeric(Trim(Right$(strBuff, Len(strBuff) - intEnd))) Then
                        MsgBox "Resolution height not a number! Using default height of 200"
                        Blits(i).YRes = 200
                    Else
                        Blits(i).YRes = Trim(Right$(strBuff, Len(strBuff) - intEnd))
                    End If
                 End If
            ElseIf Left$(Trim$(strBuff), 7) = "blitter" Then
                intStart = InStr(7, strBuff, "=")
                strTempType = Trim$(LCase$(Right$(strBuff, Len(strBuff) - intStart)))
                For j = 0 To 6 Step 1   'loop through the string lookup of blitters
                    If strTempType = BlitterList(j) Then Exit For
                Next j
                If j = 7 Then 'whoops, bad blitter name!
                    MsgBox "Error on blitter line: '" & strTempType & "' is not a valid blitter name." & vbCrLf & vbCrLf & "Using 'normal' blitter.", vbCritical
                    j = 0 'set to normal
                End If
                Blits(i).Blitter = j    'and find the right one.
            ElseIf Left$(Trim$(strBuff), 6) = "driver" Then
                intStart = InStr(7, strBuff, "=")
                strTempType = Trim$(LCase$(Right$(strBuff, Len(strBuff) - intStart)))
                For j = 0 To 8 Step 1
                    If strTempType = DriverList(j) Then Exit For
                Next j
                If j = 9 Then
                    MsgBox "Error on driver line: '" & strTempType & "' is not a valid driver name." & vbCrLf & vbCrLf & "Using 'auto' driver.", vbCritical
                    j = 0   'set to auto
                End If
                Blits(i).Driver = j
            ElseIf Left$(Trim$(strBuff), 4) = "flip" Then
                Blits(i).Flip = True
            ElseIf Left$(Trim$(strBuff), 5) = "vsync" Then
                Blits(i).VSync = True
            ElseIf Left$(Trim$(strBuff), 1) = ";" Then
                Blits(i).Comments = Blits(i).Comments & strBuff & vbCrLf
            End If
        Loop Until strBuff = "" Or EOF(intFileno)
        i = i + 1
        If i Mod 100 = 0 Then
            'alloc more memory
            ReDim Preserve Blits(i + 100)   'heh pretty easy
        End If
        Do Until strBuff <> "" Or EOF(intFileno) 'skip through the blank spaces
            Line Input #intFileno, strBuff
        Loop
    Loop
Close intFileno
'--
'now insert them into the visual elements
cModes = i - 1 'offset the last one we over-read
With lstBlitters
    .AddItem "Normal"
    .AddItem "Double"
    .AddItem "Scanlines"
    .AddItem "TV Mode"
    .AddItem "Eagle"
    .AddItem "Parallel"
    .AddItem "Double TV Mode"
End With
With lstDrivers
    .AddItem "Auto"
    .AddItem "Safe"
    .AddItem "VGA"
    .AddItem "Mode X"
    .AddItem "VESA 1"
    .AddItem "VESA 2B"
    .AddItem "VESA 2L"
    .AddItem "VESA 3"
    .AddItem "VBEAF"
End With
Dim bFound As Boolean
For j = 0 To cModes Step 1   'add names
    lstModes.AddItem Blits(j).Name
    'add XResolution, but only one of each.
    bFound = False
    For k = 0 To cboXRes.ListCount - 1 Step 1
        If Blits(j).XRes = cboXRes.List(k) Then
            bFound = True
            Exit For
        End If
    Next k
    If bFound = False Then
        cboXRes.AddItem Blits(j).XRes
    End If
    'YRes
    bFound = False
        For k = 0 To cboYRes.ListCount - 1 Step 1
        If Blits(j).YRes = cboYRes.List(k) Then
            bFound = True
            Exit For
        End If
    Next k
    If bFound = False Then
        cboYRes.AddItem Blits(j).YRes
    End If
Next j
'trim trailing vbCrLf from all comments, if present
For j = 0 To cModes Step 1
    If Len(Blits(j).Comments) Then
        Blits(j).Comments = Left$(Blits(j).Comments, Len(Blits(j).Comments) - 2)
    End If
Next j
lstModes.ListIndex = 0
Exit Sub
FileNotFoundErr:
    If Err.Number = 53 Then 'file not found
        MsgBox "Meka.blt not found in that directory! Please choose again. You must choose the directory that contains Meka."
        frmChooseDir.Show vbModal
        MekaPath = frmChooseDir.Tag
        Resume
    Else
        Err.Raise Err.Number
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Ready." & vbCrLf & vbCrLf & "Point the mouse at something to find out how to use it."
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Set the horizontal and vertical resolution. Be careful, however: not all numbers will work. If you don't know which numbers to use, see what previously created video modes have used. Note also that for some computers the colors in 512x384 resolution are messed up."
End Sub

Private Sub imgWizard_Click()
    MsgBox "Version 1.1." & vbCrLf & "By ZackMan" & vbCrLf & vbCrLf & "http://now.at/zackman"
End Sub

Private Sub lblComments_Click()
    MsgBox "Version 1.1." & vbCrLf & "By ZackMan" & vbCrLf & vbCrLf & "http://now.at/zackman"
End Sub

Private Sub lblXRes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Set the horizontal resolution. Be careful, however: not all numbers will work. If you don't know which numbers to use, see what previously created video modes have used. Note also that for some computers the colors in 512x384 resolution are messed up."
End Sub

Private Sub lblYRes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Set the vertical resolution. Be careful, however: not all numbers will work. If you don't know which numbers to use, see what previously created video modes have used. Note also that for some computers the colors in 512x384 resolution are messed up."
End Sub

Private Sub lstBlitters_Click()
    Blits(lstModes.ListIndex).Blitter = lstBlitters.ListIndex
    UpdateWizard
End Sub

Private Sub lstBlitters_GotFocus()
lstBlitters_MouseMove 0, 0, 1, 1
End Sub

Private Sub lstBlitters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "The blitters determine what special effects, if any, Meka uses in full-screen mode." & vbCrLf & vbCrLf & "Normal:  Default one-to-one pixel mode." & vbCrLf & vbCrLf & "Double:  Use to make the image larger for higher resolutions. 512x384 makes the SMS almost full-screen." & vbCrLf & vbCrLf & "Scanlines:  Only draw every other line. Simulates a TV for slow computers." & vbCrLf & vbCrLf & "TV Mode:  Draw every other line at half brightness. Simulates a TV for fast computers." & vbCrLf & vbCrLf & "Eagle:  Rounds edges of graphics. Looks cool, but causes a true sacrilege to good old art." & vbCrLf & vbCrLf & "Parallel: Show even and odd frames side by side. Use with high resolution to simulate 3D." & vbCrLf & vbCrLf & "TV Mode + Double: Combine effects of TV Mode and Double. Use with a fast computer and high resolution."
End Sub

Private Sub lstDrivers_Click()
    Blits(lstModes.ListIndex).Driver = lstDrivers.ListIndex
End Sub

Private Sub lstDrivers_GotFocus()
lstDrivers_MouseMove 0, 0, 1, 1
End Sub

Private Sub lstDrivers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Drivers determine how Meka changes to the X and Y resolution you specify. The VESA drivers are fastest, with the higher numbers being faster, but they are not installed on all computers. Auto is the safest, but not always the fastest since it doesn't automatically use the highest VESA mode." & vbCrLf & vbCrLf & "Even high VESA drivers cannot create all resolutions. Be careful!"
End Sub

Private Sub lstModes_Click()
With lstModes
cboXRes.Text = Blits(.ListIndex).XRes
cboYRes.Text = Blits(.ListIndex).YRes
lstBlitters.ListIndex = Blits(.ListIndex).Blitter
lstDrivers.ListIndex = Blits(.ListIndex).Driver
chkShowInGUI.Value = IIf(Blits(.ListIndex).ShowInGUI, vbChecked, vbUnchecked)
chkFlip.Value = IIf(Blits(.ListIndex).Flip, vbChecked, vbUnchecked)
chkVSync.Value = IIf(Blits(.ListIndex).VSync, vbChecked, vbUnchecked)
txtComments = Blits(.ListIndex).Comments
End With

End Sub

Private Sub lstModes_DblClick()
Dim newName As String
    'here we put change name code
    With lstModes
    newName = InputBox("Enter new name for mode", "New Name", Blits(.ListIndex).Name)
    If newName <> "" Then
        Blits(.ListIndex).Name = newName
        .List(.ListIndex) = newName
    End If
    End With
End Sub

Private Sub lstModes_GotFocus()
lstModes_MouseMove 0, 0, 1, 1
End Sub

Private Sub lstModes_KeyDown(KeyCode As Integer, Shift As Integer)
Dim beginIndex As Integer
Dim tempBlit As BlitModes
    
    If KeyCode = vbKeyReturn Then
        lstModes_DblClick
    ElseIf KeyCode = vbKeyUp Then
        If Shift = vbCtrlMask Then
            'if Ctrl+Arrow, move mode
            With lstModes
            If .ListIndex > 0 Then
                beginIndex = .ListIndex 'save info we'll need at the end
                tempBlit = Blits(beginIndex - 1)
                lstModes.AddItem Blits(beginIndex).Name, beginIndex - 1
                lstModes.RemoveItem beginIndex + 1
                Blits(beginIndex - 1) = Blits(beginIndex)
                Blits(beginIndex) = tempBlit
                 .ListIndex = beginIndex
            End If
            End With
        End If
    ElseIf KeyCode = vbKeyDown Then
        If Shift = vbCtrlMask Then
            With lstModes
            If .ListIndex < .ListCount - 1 Then
                beginIndex = .ListIndex 'save info we'll need at the end
                tempBlit = Blits(beginIndex + 1)
                lstModes.AddItem Blits(beginIndex).Name, beginIndex + 2
                lstModes.RemoveItem beginIndex
                Blits(beginIndex + 1) = Blits(beginIndex)
                Blits(beginIndex) = tempBlit
                .ListIndex = beginIndex
            End If
            End With
        End If
    End If
End Sub

Private Sub lstModes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Choose the mode you want to change. Double click the mode to change its name." & vbCrLf & vbCrLf & "Press Ctrl+Up/Down to change its position in this program, Meka.BLT, and Meka itself."
End Sub

Private Sub picWizard_Click()
    picWizard.Line (50, 50)-(Blits(lstModes.ListIndex).XRes, Blits(lstModes.ListIndex).YRes), QBColor(1), BF
End Sub

Private Sub picWizard_Paint()
    UpdateWizard
End Sub

Private Sub txtComments_Change()
    Blits(lstModes.ListIndex).Comments = txtComments.Text
End Sub

Private Sub txtComments_GotFocus()
txtComments_MouseMove 0, 0, 1, 1
End Sub

Private Sub txtComments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "The comments providing special explanation for this mode. You must prefix each line with ';', the comment indicator." & vbCrLf & vbCrLf & "Note that most modes are self explanatory and do not have comments."
End Sub

Private Function CommentOut(strComment, Index As Integer) As String
'comment the string we're passed out if the blits(index).showingui is false
If Blits(Index).ShowInGUI = False Then
    CommentOut = ";" & strComment
Else
    CommentOut = strComment
End If
End Function

Private Sub UpdateWizard()
        With Blits(lstModes.ListIndex)
        'update virtual screen relative resolutions
        shpGG.Visible = True
        shpSMS.Visible = True
        If .Blitter = DoubleSize Or .Blitter = TVMode_Double Or .Blitter = Eagle Then
            'it's a doubled mode
            'Set internal screen to double
            shpSMS.Width = 512
            shpSMS.Height = 384
            shpGG.Width = 320
            shpGG.Height = 288
        ElseIf .Blitter = Parallel Then
            'it's a doubled horizontal-only mode
            shpSMS.Width = 512
            shpSMS.Height = 192
            shpGG.Width = 320
            shpGG.Height = 144
        ElseIf .Blitter = Scanlines Or .Blitter = TVMode Then
            'it's a doubled vertical-only mode(not sure about this because Perfect 2 doesn't look like it should work)
            shpSMS.Width = 256
            shpSMS.Height = 384
            shpGG.Width = 160
            shpGG.Height = 288
        Else
            'normal
            shpSMS.Width = 256
            shpSMS.Height = 192
            shpGG.Width = 160
            shpGG.Height = 144
        End If
        shpScreen.Width = .XRes
        shpScreen.Height = .YRes
        shpSMS.Left = (shpScreen.Width - shpSMS.Width) / 2 + shpScreen.Left
        shpSMS.Top = (shpScreen.Height - shpSMS.Height) / 2 + shpScreen.Top
        shpGG.Left = (shpScreen.Width - shpGG.Width) / 2 + shpScreen.Left
        shpGG.Top = (shpScreen.Height - shpGG.Height) / 2 + shpScreen.Top
        If shpSMS.Width > shpScreen.Width And shpSMS.Height > shpScreen.Height Then
            shpSMS.Visible = False
        End If
        If shpGG.Width > shpScreen.Width And shpGG.Height > shpScreen.Height Then
            shpGG.Visible = False
        End If
        'end virtual screen res
        'update rest of preview
        imgBlitter.Picture = imgBlitters(.Blitter)
        imgVSync.Picture = imgVSyncs(-.VSync)
        imgFlip.Picture = imgFlips(-.Flip)
        imgShowinGUI = imgGUIs(-.ShowInGUI)
        If .VSync And .Flip Then
            imgFlipAndVSync.Visible = True
        Else
            imgFlipAndVSync.Visible = False
        End If
        'end preview pic
        'draw line from bottom edges of shpScreen to top edges of imgBlitter
        picWizard.Cls
        'show lines from Screen to Blitter
        picWizard.Line (shpScreen.Left, shpScreen.Top + shpScreen.Height)-(imgBlitter.Left, imgBlitter.Top)
        picWizard.Line (shpScreen.Left + shpScreen.Width, shpScreen.Top + shpScreen.Height)-(imgBlitter.Left + imgBlitter.Width, imgBlitter.Top)
        'show lines from VSync to Flip
        If .Flip = True Then
            picWizard.Line (imgVSync.Left + (imgVSync.Width * 2 / 3), imgVSync.Top + imgVSync.Height)-(imgFlip.Left + (imgFlip.Width / 2), imgFlip.Top)
        Else
            picWizard.Line (imgVSync.Left + (imgVSync.Width * 2 / 3), imgVSync.Top + imgVSync.Height)-(imgFlip.Left, imgFlip.Top)
        End If
        picWizard.Line (imgVSync.Left + imgVSync.Width, imgVSync.Top + imgVSync.Height)-(imgFlip.Left + imgFlip.Width, imgFlip.Top)
        End With 'blits(lstMode.listindex)
End Sub
