Attribute VB_Name = "ChatterBox"
Option Explicit
'ChatterBox module. Author:Nathan Sanders
'(20/08/1998)
'Note that this module must have explore.bas in the same project since I need the name of the
'file to be opened. Opener from explore.bas also must be filled before calling LoadScript,
'RunScript, or InterpretScriptLine

'For help on the PeopleScript itself, see the Hi Mom!.scr file, and study InterpretScriptLine
'to see how it handles the different commands.
Private Script() As String
'Hi Mom!.map
'Hi Mom!.scr
Public Function LoadScript(ScriptNum As Integer) As Boolean
'this function loads a script in from a PeopleScript file(*.scr). You must pass it the number
'of the script to load(usually stored in a Thing.Desc variable of a Person). It returns False
'if it cannot find the script number in the jump table. This means that you MUST include a
'jumptable in your script files. The reason that I am using them is to speed up the read of
'the file(so I don't have to search for each 's', then check each 's' to see if the number
'is the script number.
Dim Temp As String
Dim ScriptOpen As String
Dim ScriptFileNo As Integer
Dim Count As Integer
Dim LineNumber As Long
Dim Found As Boolean
Dim Intro As String 'the intro that runstring shows for the first call to OpenTalkBox
    'init the counter variable
    Count = 0
    'open the file
    ScriptOpen = Left(Opener, Len(Opener) - 3) & "scr"
    ScriptFileNo = FreeFile
    Open ScriptFileNo For Input As ScriptFileNo
    'now loop through the jumptable until we find our 'script'.
    Do
        Line Input #ScriptFileNo, Temp
        If (CInt(Mid$(Temp, 1, 3))) = ScriptNum Then    'j000-100
            Found = True            '000 means the script number, 100 means the line number
            Exit Do                 'for the script.
        End If
        Count = Count + 1
    Loop While Left(Temp, 1) = "j"  'until we reach the end of the jump table
    
    If Found = False Then   'not in jump table, bad script-writer
        LoadScript = False  'so we bail out.
        Exit Function
    End If
    'otherwise, keep going
    
    'init the counter variable
    LineNumber = CInt(RIGHT(Temp, 3))   'this is the number given to us by the jumptable.
    For Count = Count To LineNumber Step 1
        Line Input #ScriptFileNo, Temp
    Next Count
    
    'make sure we get the intro to display to the user when we call opentalkbox the first time
    Intro = RIGHT(Temp, Len(Temp) - 4)  's000Hello!
    
    Count = 0   're-init count variable to 0 relative to the script array instead of the
    'script file.
    ReDim Preserve Script(2000) As String   'this is the current max number of lines per script.
    'we can change this if need be
    Do
        Line Input #ScriptFileNo, Script(Count)
        Count = Count + 1
    Loop Until Script(Count) = "}"
    ReDim Preserve Script(Count) As String  'now cut Script down to size so that we aren't
    'wasting memory.
    
    'we're done!
    LoadScript = True   'and we were successful at Loading the Script.
End Function

Public Sub RunScript()
'this sub actually runs the script and is designed to be called right after LoadScript. It
'actually handles all calls to InterpretScriptLine as well, but I am thinking of including a
'function called InterpretScript which would behave exactly as InterpretScriptLine with the
'exception that it would not use the Script() array but rather a string passed to it.
Dim Count As Integer    'this is the linecount with which we keep track of how complete the
'script is.
Dim i As Integer 'this is a simple counter for use in loops unrelated to updating the line
'position of the script.
Private ThreadLineno(0 To 5) As Integer 'these are bookmarks of the line numbers at which the threads
'start. They are declared Private because eventually I'm going to make them global. For now
'though, this is the most convenient place to stick them because I'm not actually using this code yet.
Dim HeadLineno(17) As Integer 'these are bookmarks of the headings(i.e. command buttons
'on opentalkbox)    'there are 3 per thread. if we have 6 current threads(which I am assuming
'throughout this module, then that means we have 18 of them.
Dim HeadText(0 To 17) As String
    Count = 0
        Do Until Script(Count) = "}"    'this loop gets all of the thread line numbers.
        's000You see a stupid looking mouse.
        '.
        '.
        '.
        '}
        If Script(Count) = "t" Then     'beginning of a thread.
            
            If Mid$(Script(Count), 4, 1) = "-" Then 'the scripter is specifying a range of
            'threads 't000-999 instead of just t000
                For i = 0 To 5 Step 1
                    'threads is a global array stored in some module but I don't think this one
                    If threads(i) >= CInt(Mid$(Script(Count), 1, 3)) And threads(i) <= CInt(Mid$(Script(Count), 5, 3)) Then
                    't000-999
                    'we have to be >= than the first number and <= than the second.
                        ThreadLineno(i) = Count 'a match!!
                    End If
                Next i
            End If
        Else    'the thread structure doesn't specify a range, just a single one...
            For i = 0 To 5 Step 1
                If threads(i) = CInt(Mid$(Script(Count), 1, 3)) Then
                    ThreadLineno(i) = Count 'a match!!
                End If
            Next i
        End If  'OK: now we have the line #'s of the three threads.
        Count = Count + 1   'make sure we increment the awful thing.
    Loop
    
    
Dim j As Integer    'another counter since we're already using i. This one keeps track
    j = 0           'of how many headings we've found.
    For i = 0 To 5 Step 1   'loop through all the threads to find all the headings
        'reset count to the start point of each thread.
        Count = ThreadLineno(i)
        Do Until Script(Count) = "]"    'end of thread identifier
            If Left(Script(Count), 1) = "h" Then    'beginning of heading; we must save it and
            'its text...    (hMiney the Mouse)
                HeadLineno(j) = Count   'set the bookmark
                HeadText(j) = RIGHT(Script(Count), Len(Script(Count)) - 1)  'get the heading
                'to show the user
                j = j + 1 'long hand for j++®   'and remember to tell the loop that we've
                'found another Heading.
            End If
            Count = Count + 1
        Loop
    Next i  'OK: we've found all the headings...
    'now what we want to do is loop endlessly until the user gets bored and clicks 'Bye'.
    'we do this by first calling OpenTalkBox, then processing the result and interpreting
    'the appropriate heading. (or quitting)
Dim Result As Integer
    Do
        'Ryan, you need to learn how to pass arrays!!
        Result = opentalkbox(Intro, HeadText(0), HeadText(1), HeadText(2), HeadText(3), _
        HeadText(4), HeadText(5), HeadText(6), HeadText(7), HeadText(8), HeadText(9), _
        HeadText(10), HeadText(11), HeadText(12), HeadText(13), HeadText(14), HeadText(15), _
        HeadText(16), HeadText(17), "Bye")
        
        If Result = 18 Then 'he pressed 'Bye' so we can quit.
            Exit Sub
        End If
        'now set Count to the correct heading line #
        Count = HeadLineno(Result)
        Do Until Script(Count) = ")"    'and loop until we're done with this heading
            If InterpretScriptLine(Count) = False Then  'this means that we should quit
                Result = 18 'usually results from the 'e'(end) command inside the script.
                Exit Do
            End If
        Loop
    Loop Until Result = 18  'this means 'Bye'
End Sub
Public Function InterpretScriptLine(ByRef Count As Integer) As Boolean
'this function actually interprets and executes each line of the Script as it is passed to it.
'Script() is the array with the script in it, and count is the value that tells us where
'we are in reading the current script. This is because this function is recursive and calls
'itself for interpreting Question('q') and haVe('v') instructions. Every time this function
'interpets a value it ups the count by one(I think; I haven't coded this yet).
'If the return value is False, it means that the script should stop executing. This is
'usually because the script has the End('e') command in it.
Dim OldThread As Integer    'this keeps track of the old thread which is changed in the "d"
'command
Dim NewThread As Integer    'do you really want me to tell you about this??
Dim i As Integer    'a simple counter
    Select Case Left(Script(Count), 1) 'figure out what command the all-powerful scripter
    'is commanding us to carry out.
        Case "c"    'a simple chatty OpenTalkBox
            opentalkbox RIGHT(Script(Count), Len(Script(Count) - 1))
        Case "d"    'change a thread
            'd000-999
            OldThread = CInt(Mid(Script(Count), 1, 3))
            NewThread = CInt(RIGHT(Script(Count), 3))
            For i = 0 To 5 Step 1
                If threads(i) = OldThread Then threads(i) = NewThread
            Next i
        Case "g"    'Give the player something (unimplemented)
        Case "k"    'taKe something away from the player(unimplemented)
        Case "p"    'Put something on the world map(unimplemented)
        Case "r"    'remove something from the world map(unimplemented)
        Case "e"    'end conversation
            InterpretScriptLine = False
            Exit Sub
        
        Case "v"    'if you haVe something. This requires that the scripter supply a yes block
        'and a no block.(unimplemented)
        Case "q"    'ask the player a Question. This requires that the scripter supply a yes block
        'and a no block.
        Dim Result As Integer
            Result = opentalkbox(RIGHT(Script(Count), Len(Script(Count)) - 1), "Yes", "No")
            If Result = 0 Then  'yes
                Do Until Script(Count) = "y"    'this allows for comments in between or
                    Count = Count + 1   'to put the no block first.
                Loop
                'now we're at the start of the yes block
                Do Until Script(Count) = "?"    'the signal for the end of a question.
                'but it may change later to ">" for consistency...
                    If InterpretScriptLine(Count) = False Then  'there is an "e" embedded
                        InterpretScriptLine = False         'somewhere in there.
                        Exit Sub
                    End If
                Loop
            Else    'probably no
                Do Until Script(Count) = "n"
                    Count = Count + 1
                Loop
                'now we're at the start of the no block
                Do Until Script(Count) = "?"
                    If InterpretScriptLine(Count) = False Then  'we found an "e" somewhere.
                        InterpretScriptLine = False
                        Exit Sub
                    End If
                Loop
            End If
    End Select
    Count = Count + 1   'make sure we increment the line count.
    InterpretScriptLine = True  'tell the user that it's OK to continue.
End Function

