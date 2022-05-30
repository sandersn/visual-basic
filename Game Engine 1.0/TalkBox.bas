Attribute VB_Name = "TalkBox"
'Gravestone:

'Nathan Sanders
'I got this Module from Ryan about a week ago, and have since made some changes(kompreneble!)
'it is now (25/09/1998)
'These are the changes so far:
'   1.changed the variable names Pic1,2 to more meaningful(Hungarian) names.
'   2.Changed all TalkBox code to be started with a LoadTalkBox, repeatedly OpentalkBoxed, and closed with a
'unload TalkBox. Also added SetPicture subs for both forms; they are not included inside the opentalkbox call any more.
'   3.Changed the Autosize property of the Label to true.
'   4.Moved code to change the picture property of the forms out of OpenTalkBox into separate functions to be called
'before OpenTalkBox and after LoadTalkBox.
'   5.Added(currently uncoded) subs to change the font name and size of 1.the speech label, and 2. the choice buttons.
'       [much later...]oh yeah, I finished those a long time ago, I just have never updated the headstone :)
Public ClickToEnd As Boolean
Public Choice As Integer
' *** CONST ***
Private Const BUTTONSPACE As Integer = 493  'this is how many twips it is for one button plus a gray space at the top.
Private Const BUTTONTOBOTTOM As Integer = 115   'the space from the bottom of a button to the bottom of a form.
'(or close enough)
Private Const MINFONTSIZE As Integer = 6
Private Const NUMBUTTONROWS As Integer = 2  'how many rows(x) of buttons we allow. Right now it's 2 because that's as much
'as a 640x480 screen can support across(untested).
'*** END CONST ***


Public Sub LoadTalkBox()
'load, but not show the two forms. This allows the programmer to call TalkBox
'and any of its functions before unloading it.
Dim Cmd As CommandButton
    Load frmTalkBox
    Load frmButtons
    For Each Cmd In frmButtons.Controls 'init all controls to 0
        With Cmd
            .Caption = ""
            .Enabled = False
        End With
    Next Cmd
    'set an initial position for frmTalkbox
    frmTalkBox.Top = (Screen.Height - (frmTalkBox.Height)) / 2
    frmTalkBox.Left = (Screen.Width - (frmTalkBox.Width + frmButtons.Width)) / 2
    '5721 = frmButtons.left
    'set an initial position for frmButtons
    frmButtons.Left = (frmTalkBox.Left + frmTalkBox.Width) + 1
    frmButtons.Top = frmTalkBox.Top

End Sub
Public Sub UnloadTalkBox()
'unload the forms. Note that this has little effect on the variables in the module, so be careful with your code!
    Unload frmButtons
    Unload frmTalkBox
End Sub
Public Sub SetTalkBoxBackGround(strTalkBoxBackGround As String)
    frmTalkBox.Picture = LoadPicture(strTalkBoxBackGround)
End Sub
Public Sub SetButtonsBackGround(strButtonsBackGround As String)
    frmButtons.Picture = LoadPicture(strButtonsBackGround)
End Sub
Public Sub SetTalkBoxFont(FontName As String, Optional FontSize As Integer = -1)
'note that this function will not let you set the font to less than six point.
    'set the Speech Label's Fontname and optionally fontsize
    With frmTalkBox.lblSpeech
        .FontName = FontName
        If FontSize > MINFONTSIZE Then
            .FontSize = FontSize
        End If
    End With
End Sub
Public Sub SetButtonFont(FontName As String, Optional FontSize As Integer = -1)
'note that this function will not let you set the font to less than six point.
Dim Cmd As CommandButton
    For Each Cmd In frmButtons.Controls
        With Cmd
            .FontName = FontName
            If FontSize > MINFONTSIZE Then
                .FontSize = FontSize
            End If
        End With
    Next Cmd
End Sub

Public Function OpenTalkBox(speech As String, ParamArray strChoices()) As Integer
'NOTE: this function now requires that you fill the choices in order. No more of this
'OpenTalkBox("Yo!",,,"Hello Man",,"Bye") for cool spacing effects.
'It must now look like this: OpenTalkBox("Yo!","Hello man", "Bye")
Dim intChoices As Integer   'this is the number of choices in the ParamArray Choice
    intChoices = 0
Dim i   'this is the variant used in the for each loop.
Static intMaxEnabled As Integer 'the number that were enabled last time.
    
    'If frmButtons.Visible = True Then frmButtons.Hide   'NOTE:why do we do this?!? Shouldn't we just leave it alone?
    'If frmTalkBox.Visible = True Then frmTalkBox.Hide   'or is it for cosmetic reasons?

    frmTalkBox.lblSpeech.Caption = speech 'Display text in the textbox
    'determine how many of the params were filled.
    
    For Each i In strChoices
            intChoices = intChoices + 1 'if choice = 1
            With frmButtons.cmdChoice(intChoices - 1) 'then reference cmdChoice(0)
                .Caption = i
                .Visible = True
                .Enabled = True
            End With
    Next i
    'Set ClicktoEnd(i.e. whether or not you have to push buttons to get rid of the TalkBox)
    If intChoices = 0 Then  'if no choices passed, then no buttons
        ClickToEnd = True
        'set frmTalkbox's position
        frmTalkBox.Left = (Screen.Width / 2) - (frmTalkBox.Width / 2) 'CenterScreen
        frmTalkBox.Top = (Screen.Height / 2) - (frmTalkBox.Height / 2)
    Else                            'else we've got buttons, and the player can't just click to make us go away
        ClickToEnd = False
        'set frmTalkBox's position
        frmTalkBox.Top = (Screen.Height - (frmTalkBox.Height)) / 2
        frmTalkBox.Left = (Screen.Width - (frmTalkBox.Width + frmButtons.Width)) / 2
        'now disable the extraneous buttons.
        If intChoices < intMaxEnabled Then
            'here we have to make sure that there are at least as many buttons as there were last time.
            'if not, then we have to disable some of the buttons
            Dim x As CommandButton
            For i = intChoices To frmButtons.Controls.Count - 1 Step 1
                With frmButtons.cmdChoice(i)
                    .Caption = ""
                    .Enabled = False
                End With
            Next i
        End If
        intMaxEnabled = intChoices  'so we'll know how many buttons were enabled NEXT time.
    End If
    'Figure out how many rows of buttons to show.
    If (intChoices Mod NUMBUTTONROWS) = 1 Then    'if there are an odd number of buttons show the number of rows((buttons / 2) + 1).
    With frmButtons
        .Height = (((intChoices \ NUMBUTTONROWS) + 1) * BUTTONSPACE) 'note the use of int division.
        .cmdChoice(intChoices).Visible = False  'make sure that the odd button is INvisible to the user.
    End With
    Else    'hopefully 0, so an even number
        frmButtons.Height = ((intChoices \ NUMBUTTONROWS) * BUTTONSPACE)
    End If
    Choice = NONE 'set choice to a null value so that the forms can tell if its been filled yet.
    frmTalkBox.Show vbModal 'Show the TalkBox form, in modal mode, allowing the calling form to wait for an answer
    '(if ClickToEnd = False, then frmButtons is never shown)
    OpenTalkBox = Choice    'return a -1(NONE) if given no choice params.
End Function
