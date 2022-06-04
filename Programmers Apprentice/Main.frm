VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "The Programmer's Apprentice: Idea Demo"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About this whole idea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton cmdLesson3 
      Caption         =   "Lesson &3"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdLesson2 
      Caption         =   "Lesson &2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdLesson1 
      Caption         =   "Lesson &1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdIntro 
      Caption         =   "&Introduction"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Caption         =   "ogrammer's Apprenti"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   860
      TabIndex        =   2
      Top             =   350
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Caption         =   "R                             C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   300
      Width           =   3615
   End
   Begin VB.Label lblTitle 
      Caption         =   "P                               E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strName As String

Private Sub cmdAbout_Click()
    SetTalkBoxFont "Times New Roman", 14
    OpenTalkBox "This program is an idea to teach students Qbasic or some other programming language in the story setting of a RPG type setting. " + _
        "It takes place inside of a computer program that, unbeknownst to the owner is running this program in the background. Inside of this program there is a " + _
        "bustling world full of adventures and trade in a generally medieval [sic] setting."
    OpenTalkBox "In this world, programmers are one of the most powerful of the trades that can " + _
        "be learnt as they can directly(or almost directly) manipulate the game world around them. It is in this setting that the player finds himself. He starts(and stays" + _
        "for at least all of this program...I'm not sure about this part yet) as an apprentice who is apprenticed to a programmer for 7 years to learn the trade."
    OpenTalkBox "The player " + _
        "learns QBasic(or whatever) as the apprentice does. The advantage of this idea is both the interest generated by playing a game as well as possibly easier " + _
        "learning because of the ready visual analogies."

End Sub

Private Sub cmdIntro_Click()
    SetTalkBoxFont "Arial", 11
    TalkBox.OpenTalkBox "Hello, apprentice. You are welcome to my house. Since you'll be under my tutelage in one way or another for the next 7 years, let's " + _
        "start by exchanging names. Mine is Belgarath. What's yours?"
    strName = InputBox("Please enter your name. Press Start when finished. :)", "Name", "Chrono")
    OpenTalkBox "Well, " + strName + ", it's late already. Let's eat supper and we'll begin your training tomorrow morning."
    OpenTalkBox "Square® sleep music fills the air."
    OpenTalkBox "Next morning:"
    cmdIntro.Enabled = False
    cmdLesson1.Enabled = True
    cmdLesson2.Enabled = True
    cmdLesson3.Enabled = True
End Sub

Private Sub cmdLesson1_Click()
Dim strCode As String
Dim intStart As Integer, intEnd As Integer
Dim intPrint As Integer, strParam As String
    SetTalkBoxFont "Arial", 11
    OpenTalkBox "OK, " + strName + " for your first lesson, you will learn how to use the PRINT command. All this command does is project the words that you" + _
    " pass to it in the air in front of wherever the program is being run. In QBasic's case, this is always right in front of your Komputer, since it cannot be compiled " + _
    "to run away from a Komputer."
    OpenTalkBox "PRINT looks like this:  PRINT ""Whatever"". All you have remember is that what you want to print must be inside the double quotes. " + _
    "Now, say I need you to print the words: 'Transaction Accepted' every time a person bought something from a general store that had an automated cash register " + _
    "that uses a QBasic program. What would be the code for it?"
    strCode = InputBox("Belgarath hands you his Komputer." + vbCrLf + "You type:", "Code box")
    'now...let's dissect this code.
    
    intStart = InStr(LCase(strCode), "print")
    If intStart = 0 Then 'no print given...
        OpenTalkBox "The Komputer beeps. It says:" + vbCrLf + """Syntax error!""" + vbCrLf + vbCrLf + "Belgarath says: That means that you typed something " + _
            "completely unexpected that QBasic has NO idea what to do with...in this case, you really need to start completely over. Remember that PRINT looks like this" + _
            "       Print ""whatever"". You need to use print to display the words 'Transaction Accepted'"
            Exit Sub
    Else
        intPrint = intStart
        intStart = InStr(LCase(strCode), """")
        intEnd = InStr(intStart + 1, LCase(strCode), """")
        If intStart < intPrint Then 'it looks like : "whatever" PRINT
            OpenTalkBox "The Komputer beeps. It says:" + vbCrLf + """Syntax error!""" + vbCrLf + vbCrLf + "Belgarath says: That means that you typed something " + _
            "completely unexpected that QBasic has NO idea what to do with...in this case, you reversed the PRINT and the ""Transaction Accepted"". Your code " + _
            "should have looked like this: PRINT ""Transaction Accepted"", but instead it looked like this: " + vbCrLf + strCode
            'need to skip to end...for now just EXIT SUB
            Exit Sub
        End If

        If intEnd = 0 Then  'no end quotes given.
            strParam = Mid$(strCode, intStart + 1, Len(strCode))
        Else
            strParam = Mid$(strCode, intStart + 1, intEnd - intStart - 1)
        End If
        OpenTalkBox "In front of Belgarath's Komputer hover the words: " + vbCrLf + vbCrLf + strParam
        If intEnd = 0 Then  'no end quotes given.
            OpenTalkBox "Belgarath says: Notice that your code looked like this, though: " + vbCrLf + strCode + vbCrLf + vbCrLf + "You left off the last double quote. " + _
            "QBasic lets you do this, but it is NOT a good habit to get into! Most other languages are much less fault tolerant than QBasic. You wouldn't want your program " + _
            "to come to a halt in the middle of your customer's attacking a dragon!"
            Exit Sub
        End If

        If LCase(strParam) <> "transaction accepted" Then
            OpenTalkBox "Belgarath says: Very good! The only problem that you have is that you didn't follow my instructions: I told you to PRINT ""Transaction Accepted"", not whatever *you* wanted to.  " + _
            "If you're going to learn anything, you need to follow my instructions. Other than that, your code looks good."
        Else
            OpenTalkBox "Very good! Next to get you started I need to tell you about variables."
        End If
    End If
End Sub

Private Sub cmdLesson2_Click()
    OpenTalkBox "This lesson will explain the use of variables and their assignment, but will not explain their types, and other 'complex' stuff, although it is stuff you need" + _
    "need to know. This info will be in some sort of Reference that should be somewhat like a Help file, but packaged differently of course."
End Sub

Private Sub cmdLesson3_Click()
    OpenTalkBox "This lesson will use multiple choice when coded--this is because I don't know how to write a complex QBasic interpreter myself."
End Sub

Private Sub Form_Load()
LoadTalkBox
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadTalkBox
End Sub
