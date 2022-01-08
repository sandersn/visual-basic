VERSION 5.00
Begin VB.Form frmPalindrome 
   Caption         =   "Palindrome Detector"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Palindrome.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPalindrome 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "ARPANET"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblAnswer 
      AutoSize        =   -1  'True
      Caption         =   "ARPANET is not a palindrome."
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblPalindrome 
      AutoSize        =   -1  'True
      Caption         =   "Word:"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   435
   End
End
Attribute VB_Name = "frmPalindrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'REQUIRES ALL VARIABLES TO BE DECLARED BEFORE USE

'NOTICE THAT VB USES MEMBER STYLE ACCESS FOR MOST OF ITS
'PRIVATE MEMBERS, ALTHOUGH IN REALITY ACCESS IS OBTAINED
'THROUGH HIDDEN ACCESS FUNCTIONS.

Private Sub txtPalindrome_Change()
    If txtPalindrome.text = "" Then 'IF THE TEXT IS EMPTY, DON'T SHOW THE ANSWER
        lblAnswer.Visible = False
    Else
        lblAnswer.Caption = """" & txtPalindrome.text & """" & _
            IIf(IsPalindrome(txtPalindrome.text), " is a palindrome.", " is not a palindrome.")
        lblAnswer.Visible = True    'MAKE SURE ANSWER IS VISIBLE
        '(THE QUAD """" ARE VB'S WAY TO ALLOW " TO APPEAR IN THE OUTPUT STREAM)
    End If  'TEXT IS EMPTY
End Sub 'txtPalindrome_Change()
Private Function IsPalindrome(text As String) As Boolean
Dim i As Integer        'A SIMPLE COUNTER (CANNOT BE DECLARED INLINE IN VB)
Dim pal As Boolean    'THE FLAG TO DETERMINE WHETHER THE WORD IS A PALINDROME
Dim q As New Queue  'DECLARE THE QUEUE AND STACK TO BE USED(NEW KEYWORD
Dim st As New Stack  'OPERATES IDENTICALLY TO C/JAVA)

    For i = 1 To Len(text) Step 1
        st.Push Mid$(text, i, 1)    'STARTS AT 1 BECAUSE OF VB
        q.Insert Mid$(text, i, 1)   'STRING STRUCTURE
    Next i 'INSERT WORD INTO STACK/QUEUE
    
    pal = True  'START AT TRUE
    Do Until st.IsEmpty Or q.IsEmpty    'LOOP UNTIL ADTS ARE EMPTY
        If st.Pop <> q.Remove Then      'AND MAKE SURE THEY'RE EQUAL
            pal = False 'IF NOT, THEN FLAG AND QUIT
            Exit Do
        End If  'NOT EQUAL
    Loop    'END LOOP UNTIL EMPTY
    IsPalindrome = pal 'AND RETURN THE ANSWER
End Function    'IsPalindrome
