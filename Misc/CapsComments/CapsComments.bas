Attribute VB_Name = "CapsComments"
Option Explicit
'note: this program only works for java files since it replaces the extension with .java and assumes a .java extension for the
'rename process
Private Const CPPBeginTag = "//"
Private Const CPPEndTag = vbCrLf
Private Const CPPEndTagUnix = 10  'the UNIX style endlines are only 1 byte kthx
Private Const CBeginTag = "/*"
Private Const CEndTag = "*/"


Sub Main()
Dim strFilename As String   'the file to work on
Dim strFileOnly As String   'just the *file* name not including directory structure
Dim strPathOnly As String   'just the path(probably short)
Dim intBeginFile As Integer 'where the filename begins of the total path\filename
Dim strOriginalFilename As String   'the file we started with
Dim intFileno As Integer    'file number for file so we can open it
Dim ch As String    'holds the two characters being analysed
Dim readText As String  'the text to be ucased when we reach the end of a comment
Dim bReading As Boolean 'inside a comment? t/f
Dim bSkipping As Boolean    'inside a ~~(which turns off capitilisation)? t/f
Dim endTag As String    'whether we're inside a C or C++ comment
Dim pAddr As Long   'current address we're reading from
Dim textAddr As Long    'address to put the ucased readText back into
'read filename from command line
strOriginalFilename = Trim$(Command)
ch = "  "
strFileOnly = Dir(strOriginalFilename, vbNormal)
intBeginFile = 0
Do While InStr(intBeginFile + 1, strOriginalFilename, "\") > 0
    intBeginFile = InStr(intBeginFile + 1, strOriginalFilename, "\")
Loop
strPathOnly = Left$(strOriginalFilename, intBeginFile)
If strFileOnly = "" Or strOriginalFilename = "" Then
    MsgBox "File not found. Retype filename"
    End
Else
'create a _final.java file to work on
strFilename = strPathOnly & Left$(strFileOnly, Len(strFileOnly) - 5) & "_final.java"
FileCopy strOriginalFilename, strFilename
'open file in binary mode
    intFileno = FreeFile
    pAddr = 1
    Open strFilename For Binary As #intFileno
    Do

        Get #intFileno, pAddr, ch

        If bReading Then
            'check to see if we need to skip--check for skip tag
            If ch = "~~" Then
                bSkipping = Not bSkipping
'store and start ignoring
                If bSkipping Then
'Ucase(text found)
                    readText = UCase$(readText)
'write back to file
                    If readText <> "" Then
                        Put #intFileno, textAddr + 2, Left$(readText, Len(readText) - 1)
                    End If
'reset readText for when we stop skipping again
                    readText = ""
                Else
                    textAddr = pAddr
                End If
            ElseIf bSkipping = False Then   'if we're not skipping, and did not find a skip tag
'read all text in until find a newl or */
                If ch = endTag Then
'Ucase(text found)
                    readText = UCase$(readText)
'write back to file
                    Put #intFileno, textAddr + 2, Left$(readText, Len(readText) - 1)
                    bReading = False
                    bSkipping = False
                ElseIf endTag = CPPEndTag And Left$(ch, 1) = Chr$(CPPEndTagUnix) Then
'Ucase(text found)
                    readText = UCase$(readText)
'write back to file
                    Put #intFileno, textAddr + 2, Left$(readText, Len(readText) - 1)
                    bReading = False
                    bSkipping = False
                Else
                        readText = readText & Right$(ch, 1)
                End If
            End If
        Else
'loop through each address of the file looking for // or /*
            If ch = CBeginTag Or ch = CPPBeginTag Then
                endTag = IIf(ch = CBeginTag, CEndTag, CPPEndTag)
                textAddr = pAddr
                readText = ""   'reinit readText
                bReading = True
            End If
        End If
        pAddr = pAddr + 1
    Loop Until EOF(intFileno)
    Close intFileno
    MsgBox "Process. Now be happy", vbAbortRetryIgnore + vbQuestion
End If

    




'input:filename
'output:confirmation to user/maybe show contents of updated file

End Sub
