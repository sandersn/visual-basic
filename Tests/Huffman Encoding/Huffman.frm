VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHuffman 
   Caption         =   "Huffman Code Analysation and anything else you want done with them"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenHuffmanTree 
      Caption         =   "&Do not push this button"
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtResults 
      Height          =   5415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Huffman.frx":0000
      Top             =   0
      Width           =   7695
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "This area a toxic waste zone. Do not enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3015
      Left            =   9480
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblCChars 
      AutoSize        =   -1  'True
      Caption         =   "By Nathan Sanders"
      Height          =   195
      Left            =   7800
      TabIndex        =   3
      Top             =   1920
      Width           =   1380
   End
End
Attribute VB_Name = "frmHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intFileno As Integer
Dim strFilename As String
'Private Type Node
'    Parent As Node
'    lChild As Node
'    rChild As Node
'End Type
Private Type HuffmanVertex
    Char As String * 1
    Frequency As Long
'    pNode As Node
End Type
Dim hufSeq(0 To 255) As HuffmanVertex 'array of frequency values
Dim cChars As Integer 'keep track of how many different chars are actually used

'Huffman frequency analyzer that may be turned into a Huffman code generator that may be turned into a Huffman code
'cracker that may be turned into a Japanese Huffman code cracker...grandiose plans, ne?
'(also might add a ASCII/SJIS/EUC to Huffman convertor)
'Notes:
'this assumes ASCII encoding currently
'this could easily be turned into a simple frequency analyser at this point by adding 1.An option to determine length of sample
'(ie 2 or 3 chars in case the lucky chaps have the ability to use DTE) 2.Skip spaces option/skip returns in case they are hostile
'to that sort of thing

Private Sub cmdAnalyze_Click()
'**=undone
Dim i As Integer, j As Integer    'counter
Dim lMin As Long    'holds minimum value when figging LCD
Dim intLCD As Integer
Dim ch As String * 1    'char holder
Dim strMsg As String    'a composition string to display in txtResults
'heyheyhey here, the code goes

    'get the sheer totals of each char
    Do
        Get #intFileno, , ch
        hufSeq(Asc(ch)).Frequency = hufSeq(Asc(ch)).Frequency + 1
        hufSeq(Asc(ch)).Char = ch
        'DoEvents
    Loop Until EOF(intFileno)
    '**order results descending..the highest are at the low end of the array
    'bubble sort since I am too unlearned to do another
    For i = 0 To 255 Step 1
        For j = i To 255 Step 1
            If hufSeq(i).Frequency < hufSeq(j).Frequency Then
                Swap hufSeq(i), hufSeq(j)
            End If
        Next j
    Next i
    'print results (eliminate later maybe)
    For i = 0 To 255 Step 1
        If hufSeq(i).Frequency <> 0 Then
            strMsg = strMsg & Asc(hufSeq(i).Char) & "(" & hufSeq(i).Char & ") = " & hufSeq(i).Frequency & "   "
            If i Mod 2 = 0 Then
                strMsg = strMsg & vbCrLf
            End If
            cChars = cChars + 1
        End If
    Next i
    strMsg = "Total characters: " & LOF(intFileno) & vbCrLf & "Number of different characters: " & cChars & vbCrLf & vbCrLf & strMsg
    txtResults.Text = strMsg
End Sub

Private Sub cmdGenHuffmanTree_Click()
'assumes other two buttons have already been pressed or will crash. yay!
Dim treeroot As HufNode
    treeroot = Huffman(hufSeq, cChars)
End Sub

Private Sub cmdOpen_Click()
On Error GoTo OpenErr   'just quit if they hit cancel
    Close   'close any already open files
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    intFileno = FreeFile
    strFilename = CommonDialog1.filename
    Open strFilename For Binary As #intFileno
OpenErr:
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then txtResults.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close intFileno
End Sub
Private Sub Swap(huf1 As HuffmanVertex, huf2 As HuffmanVertex)
Dim hufTemp As HuffmanVertex
    hufTemp = huf2
    huf2 = huf1
    huf1 = hufTemp
End Sub
Private Function Huffman(f() As HuffmanVertex, n As Integer) As HufNode 'Node
'Private Function Huffman(start as Integer, n as Integer) As HufNode
'where n is the length of the sequence in the array
'and f is the start of the sequence in the array
Dim root As New HufNode
Dim i As Integer
If n = 2 Then
    root.Left = f(0)
    root.IsLeft = True
    root.Right = f(1)
    root.IsRight = True
    Set Huffman = root
    Exit Function
Else
Dim smaller(0 To n - 2) As HuffmanVertex
    'find two smallest values and their indices
Dim small1 As HuffmanVertex
Dim small2 As HuffmanVertex
Dim small1index As Integer
Dim small2index As Integer
Dim minval As Long
    minval = 2000000000   'we hope its less than this, anyway...
    For i = 0 To n - 1 Step 1 'find min
        If minval > f(i).Frequency And f(i).Frequency > 0 Then
            small1index = i
            minval = f(i).Frequency
        End If
    Next i
    minval = 2000000000   'we hope its less than this, anyway...
    For i = 0 To n - 1 Step 1 'find min
        If minval > f(i).Frequency And i <> small1index Then
            small2index = i
            minval = f(i).Frequency
        End If
    Next i
    small1 = f(small1index)
    small2 = f(small2index)
    For i = 0 To n - 3 Step 1
        If i = small1index Or i = small2index Then
            'skip
        Else
            If (small1.Frequency + small2.Frequency > f(i).Frequency) Then
                Dim temp As HuffmanVertex
                temp.Char = small1.Char + small2.Char
                temp.Frequency = small1.Frequency + small2.Frequency
                smaller(i) = temp
                smaller(i + 1) = f(i)
                i = i + 1
            Else
                smaller(i) = f(i)
            End If
        End If
    Next i
'    smaller(n - 2) = f(n - 2) + f(n - 1)
    Set root = Huffman(smaller, n - 1)
    'need a tree traversal to find a value which is equal to small1.Freq + small2.Freq
    Do
    Loop While True
    'rest of algorithm here
End If
End Function
'this part was mistakenly thought to be necessary, but it's not (! wow !)
'    'massage results to lowest common denominator
'    lMin = 2000000000   'we hope its less than this, anyway...
'    For i = 0 To 255 Step 1 'find min
'        If lMin > hufSeq(i).Frequency And hufSeq(i).Frequency > 0 Then lMin = hufSeq(i).Frequency
'    Next i
'    Debug.Print lMin
'    'see if the rest of the nums div by the min
'    intLCD = lMin
'    For i = 0 To 255 Step 1
'        If hufSeq(i).Frequency Mod lMin > 0 Then
'            'we found the min, yet somehow not a lowest common denominator!
'            intLCD = 1  'so the LCD becomes boring old 1
'        End If
'    Next i
'    Debug.Print intLCD
'    For i = 0 To 255 Step 1
'        hufSeq(i).Frequency = hufSeq(i).Frequency / intLCD
'    Next i
'

