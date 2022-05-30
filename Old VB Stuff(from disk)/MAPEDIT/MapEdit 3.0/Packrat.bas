Attribute VB_Name = "PackRat"
Option Explicit
'By Nathan Sanders
'1998 stuff
'(08/05) I have just typed in most of the code for this module(except PaintThings).
'Everything seems to look good(uh huh). However, I may have to change some ints to longs
'eventually. I just coded them as ints out of habit.
'(11/07) I have added the code for: MoveThing, MoveThingTo, TypeOfThing, and SaveThings.
'I also fixed the code for LoadThings. It was waay off in the way it was figuring out which
''Thing' to write where. Have not done PaintThings yet, but it shouldn't be too hard, as
'I don't have to do any clipping like PaintMap does. Also added constants for the three types
'of 'Things'. These serve to differentiate between the three
'(i.e. If TypeOfThing(Count) = PERSON Then...) and also to show the first element in arrays,
'ImageLists, and so on where each type starts.
'(i.e. imlType.ListImages(PERSON + Count).Picture = ... or something similar)
'(11/08) Have completed code for PaintThings. Note that I think that it is indeed possible
'to try to cheat and 'put' more than 10 objects in one screen because you could simply put
'the 'Thing' in one screen, but set its X,Y to be painted on a different screen. The only
'catch to this tachnique is that you would have to make SURE that it is impossible to load
'one screen into memory without the other. Otherwise you would have 'missing objects'.
'On the other hand, this could theoretically be beneficial if you wanted the player to look
'at a locked door, come one screen down and have several new objects appear. However, it
'would be much safer to simply create new objects triggered by some sort of script.
'(17/08) Added some movements constants and fixed a bug(off by 1 error) in the Load and Save
''Things' functions.
'*** 1999 ***
    '(27/08) Added a Clean function. This 'cleans' all the deleted Things so they cannot be restored later by RestoreThingsArray.
'By the way, I never documented the fact that I wrote a RestoreThings function that loops through all objects who have been previously
'deleted and restores them (using a random X value since it is the flag for deleted, and is therefore overwritten)
Public ObjOpener As String 'the filename of the 'Things' file.

'*** Constants ***
Public Const OBJ = 1 'these are constants for the ImageList and various other tests to tell where
Public Const PERSON = 33  'each type of 'Thing' begins. (Also used as identifiers that define what
Public Const MONSTER = 79 'type a 'Thing' is. (we'll need to update these every time we add
'something to the imagelist for the objects.
'movement constants:
Public Const STILL = 0
Public Const RANDOM = 1
Public Const FOLLOW = 2
Public Const SCRIPT_PATH = 3    '*** unsupported right now!! ***
Public Const ESCAPE = 4 'that is, (!FOLLOW)
Public Const SHIP = 5   'this means that the object can be Used to make the player disappear until he tries to move onto a non-sea tile.
'more VEHICLES here later.
'that's all the movement constants for now, folks!
Public Const OBJ_MAXTHINGSSCREEN As Integer = 10
Public Const OBJ_MAXTHINGSARRAY As Integer = MAP_NUMSCREENSX * MAP_NUMSCREENSY * OBJ_MAXTHINGSSCREEN
Public Const NOT_GIVEN As Integer = -99
Public Const NONE As Integer = -1
'*** End Constants ***
Type Thing  'I have changed this type from Object to Thing because there are far fewer
'naming conflicts as Object is a rather popular name inside VB.
    x As Integer
    Y As Integer
    Type As Integer 'defines 'Name', Picture, and Type of 'Thing'
    'decide exactly what we want for it. If the number is smaller than OBJ, then it is a Object. Then we
    'use Desc to look at the Description when the user wants to. If the number is bigger than
    'OBJ and smaller than MONSTER, the Thing is a Person. Then the Desc points to a Script.
    'If the number is bigger than MONSTER, the Thing is a Monster. Then Desc points to nothing yet!
    '(but might be useful to show which battle to start or something.)
    'Weight As Byte
    'Value as integer 'Do we want this? I don't know yet so am leaving it out.
    Desc As Integer 'in Objects, a pointer to a line of text in "mapname.obj"
    'in People a pointer to a Script in some Format.
    'in Monsters nothing definite, but maybe a pointer to the battle scenario they bring
    '(after all, no more random battles, right?)(actually I think no rand. bat. is bona ideo
    Movement As Byte    'see the movement constants approx. 1 screen above.
    Tag As String * 4 ' or int might even be able to instant switched out...
    'right now the first character is being used as the weight. I am doing a clunky Chr$/Asc conversion to work the string stuff
        'THIS Comment section is OLD. (i.e. it was speculation for something that wasn't working but is now, and is implemented
        'very differently. But it's fun to read :)
    'objects except by their type(for example when the player Looked at an object,
    'they would see 'pot' regardless of whether it was the Pot of Wonderfulness worth
    '2000 moneys, or the dreaded Pot of Plague!(Now that I think about it, maybe we DO
    'need a name. The problem is getting one short enough. Or on the other hand, we
    'could just link the name to the Type and have Type(123) = Picture("ShinyPot"),
    'Description("Pot of Wonderfulness") and Type(126) = Picture("GrungyPot"),
    'Description("Pot of Plague") Then if we wanted to have a tricky 'pot', we would
    'make another Type: Type(127) = Picture("ShinyPot"), Description("Pot o' Creeping Snails")
    '(or something like this)
        'END OLD
End Type
Public Things(0 To (OBJ_MAXTHINGSARRAY - 1)) As Thing 'here it is... starting at 10 objects per screen.
'I give up...I at first declared this (0 to 90), but my c influenced brain couldn't keep
'up with the kalculations. So here we are; (0 to 89)
Public Function IsThing(x As Integer, Y As Integer, Optional ByRef ArrayNum As Integer = NOT_GIVEN) As Boolean
'note that I added ByRef for emphasis only, because it is the default pass mode.
Dim Count As Integer

    For Count = 0 To (OBJ_MAXTHINGSARRAY - 1) Step 1
        If Things(Count).x = x And Things(Count).Y = Y Then
            IsThing = True
            If ArrayNum <> NOT_GIVEN Then ArrayNum = Count
            Exit Function
        End If
    Next
    IsThing = False
    ArrayNum = NONE   'this is extra...I just wanted to make double sure that if they
    'try to use this with no hit that it will mess them up.
End Function
Public Function IsThingExclude(x As Integer, Y As Integer, ExcludeNum As Integer, Optional ByRef ArrayNum As Integer = NOT_GIVEN) As Boolean
'This function sees if a 'Thing' is occupying the square of another 'Thing'(the excludenum)
'if one is, then we return the number that is occupying it.
Dim Count As Integer
    For Count = 0 To (OBJ_MAXTHINGSARRAY - 1) Step 1
        If Things(Count).x = x And Things(Count).Y = Y And Count <> ExcludeNum Then
            IsThingExclude = True
            If ArrayNum <> NOT_GIVEN Then ArrayNum = Count
            Exit Function
       End If
    Next
    IsThingExclude = False
    ArrayNum = NONE   'this is extra...I just wanted to make double sure that if they
    'try to use this with no hit that it will mess them up.

End Function

Public Function IsThingDesc(intDesc As Integer, Optional ObjScreenX As Integer = NOT_GIVEN, Optional ObjScreenY As Integer = NOT_GIVEN) As Integer
'this function takes the desc number of a 'thing' and checks to see if it is on the current screen, then if it is on the current
'object array. It then returns an integer which is the array index if found, and -1 if not
'ObjScreenX,Y are the screen to check first(must be evenly divisible by 10).(i.e. *the* border of that particular screen.
'note that these values should always be 0, 10, or 20, so there's not much room for variation. This means that you, gentle reader, must figure out for your
'self the screen you desire to search, and subtract *yourself* the number from the ScreenX,Y which the main module is graced with.
'If the desired object is not on that screen, this function checks the whole obj array.
Dim i As Integer
Dim intObjectScreen As Integer
        i = NONE    'extra init...not needed '~'
        'if U don't supply the screen co-ords, then the search will be heavily weighted toward the topleft screen.
        'but, hey, it's UR choice
        If ObjScreenX <> NOT_GIVEN Or ObjScreenY <> NOT_GIVEN Then
            intObjectScreen = ((ObjScreenY \ MAP_SCREENX) * MAP_ARRAYX) + ((ObjScreenX \ MAP_SCREENX) * OBJ_MAXTHINGSSCREEN)
            'note that this method of multiplication assumes that SelectX,Y are pointing at the person that is currently being talked to.
            For i = intObjectScreen To intObjectScreen + 9 Step 1 'loop through all the 'Things' on this particular screen.
                If Things(i).Desc = intDesc And Things(i).x <> NONE Then   'a match!!
                    IsThingDesc = i
                    Exit Function
                End If
            Next i
        End If
        'afterwards, just search the whole thing(that screen we might have searched originally is redundant)
        For i = 0 To (OBJ_MAXTHINGSARRAY - 1) Step 1 'the WHOLE object array.
            If Things(i).Desc = intDesc And Things(i).x <> NONE Then
                'a match!
                IsThingDesc = i
                Exit Function
            End If
        Next i

End Function
Public Function PutThing(SomeThing As Thing, ScreenX As Long, ScreenY As Long) As Boolean
'returns false if the screen is full
'Something is the 'Thing' to pass, ScreenX,Y are simply the current ScreenX,Y inside of the main module.
Dim Start As Integer 'the start number inside the array, figured out using the X, Y of
'the Thing passed to us.
Dim Count As Integer
    Start = (((SomeThing.Y - ScreenY) \ MAP_SCREENX) * MAP_ARRAYX) + (((SomeThing.x - ScreenX) \ MAP_SCREENX) * OBJ_MAXTHINGSSCREEN) 'yes, I know that all dividing and
    'multiplying looks krazy, but note that I am using Integer division. This means that
    'I am throwing away the remainder, then multiplying back out.(Although my grasp on the
    'reliability of this statement is still a little shaky)
    
    'for example:SomeThing.y = 42, SomeThing.x = 63;
    'screenx = 30, screeny = 50
    '(((SomeThing.y - ScreenY)\ 10) * 30) + (((SomeThing.X - ScreenX)\ 10) * 10)
    '(1 * 30) + (((SomeThing.X - ScreenX) \ 10) * 10)
    '(30) + (((SomeThing.X - ScreenX)\ 10) * 10)
    '(30) + (1 * 10)
    '30 + 10
    '40
'   |-----------------|
'   | 0-9 |10-19|20-29|
'   |0,0  |0,10 |0, 20|
'   |-----------------|
'            \/   right on the nose!!!
'   |30-39|40-49|50-59|
'   |10, 0|10,20|10,30|
'   |-----------------|
'   |60-69|70-79|80-89|
'   |20, 0|20,10|20,20|
'   |-----------------|
    '(I love examples)
    For Count = Start To (Start + 9) Step 1
        If Things(Count).x = NONE Then
            Things(Count) = SomeThing
            PutThing = True
            Exit Function
        End If
    Next Count
    PutThing = False
    SomeThing.x = NONE 'so you can check if you added it or not twice.
End Function
Public Function RemoveThing(x As Integer, Y As Integer) As Boolean
Dim Count As Integer
    If (IsThing(x, Y, Count)) Then
        Things(Count).x = NONE    'This is the flag inside code that this object is DEFUNCT,
        'deleted, outa here, blitzo, kapow, gone.
        RemoveThing = True
    Else
        RemoveThing = False 'there's nothing there to delete, defunctize, blitz, kapow, make
        'gone.
    End If
End Function
Public Sub RemoveThingArray(ArrayNum As Integer) 'we might need this if the programmer
'wants to remove things by referencing their number in the array rather than their X,Y
'(I think I'll write it just in case, so I'll have it to call...just in case)

'also if we want to add a tech or something called 'Rewind' it would be very easy do this.
'all we would have to do would be to set all 'Things' that have something besides -1 in Y,
'set their X to some convenient number on screen. For the player, this would have the
'beneficial effect of restoring some wonderful object that's worth 2000 moneys, and the bad
'effect of restoring the 22 enemies armed with 'Photon Tornados'.

    Things(ArrayNum).x = NONE
    'pretty simple, huh?
End Sub
Public Sub LoadThings(ObjFileno As Integer, x As Long, Y As Long)
'here's the sample Load function. Actually X, Y should be called screenx, screeny
'but I'm already using the name as a variable to hold the TRUE ScreenX, ScreenY numbers.
'(For example: 110, 50 are the numbers we're passed. We've got to convert them to 11, 5 to
'represent the screenx,y numbers.
Dim cX As Integer, cY As Integer, Count As Integer
Dim ScreenX As Integer, ScreenY As Integer
Dim ScreenMapSize As Integer
'x,y are called screenx, screeny inside Map Edit and look like this:(110, 50)
'therefore we need to divide them by 10 to figure out which 'screen' in the Things file
'that we need to point at. Example: (11, 5)
    ScreenX = x \ MAP_SCREENX    'prefiguring them into variables might yield a very small speed
    ScreenY = Y \ MAP_SCREENY    'gain...
    ScreenMapSize = MapXSize \ MAP_SCREENX
    For cY = 0 To (MAP_NUMSCREENSY - 1) Step 1
        For cX = 0 To (MAP_NUMSCREENSX - 1) Step 1  'for each screen...
            For Count = 0 To (OBJ_MAXTHINGSSCREEN - 1) Step 1   'get all of the Things for that screen
                'Get Fileno, (cy * (Map()'s XSize / screensize) * objsperscreen)) +
                '(X * objsjperscreen) + count
                Get #ObjFileno, ((ScreenY + cY) * (ScreenMapSize * MAP_SCREENX)) + (((ScreenX + cX) * OBJ_MAXTHINGSSCREEN) + Count) + 1, _
                Things((cY * MAP_NUMSCREENSX * OBJ_MAXTHINGSSCREEN) + (cX * OBJ_MAXTHINGSSCREEN) + Count)
            Next Count
        Next cX
    Next cY
    'Example time!
    'x = 110, y = 50, MapXSize = 200
'   screenx = 11, screeny = 5, ScreenMapSize = 20
'   Cy = 0, Cx = 0, Count = 0   'first iteration
'   get fileno (((11 + 0) * (20 * 10)) + ((11 + 0) * 10) + 0) + 1, things(0 * 30) + (0) + 0)
'   get fileno,((((11) * (200)) + ((11) * 10) + 0) + 1, things(0 + 0 + 0)
'   get fileno ((2200) + (110)) + 1, Things(0 + 0 + 0)
'   get fileno, ((2311), things(0)
'   'now all that remains is to see if 2310 is correct. But it looks right.
'   another example:
'   x = 20, y = 10, MapXSize = 30
'   screenx = 2, screeny = 1, screenmapsize = 3
'   cy = 0, cx = 0, count = 0 'first iteration
'   get fileno, (((1 + 0) * (3 * 10)) + (((2 + 0) * 10) + 0) + 1, things((0 * 30) + (0) + 0)
'   get fileno, (((1) * (30)) + (((2) * 10) + 0) + 1, things(0)
'   get fileno, (30 + 20) + 1, things(0)
'   get fileno, (51), things(0)
'   cy = 0, cx = 1, screenmapsize = 3
'   get fileno, (((1 + 0) * (30)) + (((3) * 10) + 0) + 1, things(((1) * 10) + 0)
'   get fileno, (30) + (31), things(10)
'   get fileno, 61, things(10)
End Sub
Public Sub SaveThings(ObjFileno As Integer, x As Long, Y As Long) 'comments same as above...
Dim cX As Integer, cY As Integer, Count As Integer
Dim ScreenX As Integer, ScreenY As Integer
Dim ScreenMapSize As Integer

    ScreenX = x \ MAP_SCREENX    'prefiguring them into variables might yield a very small speed
    ScreenY = Y \ MAP_SCREENY    'gain... (Notice int division)
    ScreenMapSize = MapXSize \ MAP_SCREENX
    For cY = 0 To (MAP_NUMSCREENSY - 1) Step 1
        For cX = 0 To (MAP_NUMSCREENSX - 1) Step 1  'for each screen...
            For Count = 0 To (OBJ_MAXTHINGSSCREEN - 1) Step 1   'get all of the Things for that screen
                'Put Fileno, (cy * (Map()'s XSize / screensize) * objsperscreen)) +
                '(X * objsjperscreen) + count
                Put #ObjFileno, ((ScreenY + cY) * (ScreenMapSize * MAP_SCREENX)) + (((ScreenX + cX) * MAP_SCREENX) + Count) + 1, _
                Things((cY * MAP_NUMSCREENSX * OBJ_MAXTHINGSSCREEN) + (cX * OBJ_MAXTHINGSSCREEN) + Count)
            Next Count
        Next cX
    Next cY

End Sub
Public Sub PaintThings(PicBox As PictureBox, Iml As ImageList, x As Integer, Y As Integer, ScreenX As Long, ScreenY As Long)
Dim cThings As Integer
Dim iTerrainVal As Integer
Dim TempX As Integer, TempY As Integer
Dim HighX As Integer, HighY As Integer, LowX As Integer, LowY As Integer
'try changing to map_screenx - 1 very soon.
    HighX = ScreenX + x + MAP_SCREENX
    LowX = ScreenX + x - 1
    HighY = ScreenY + Y + MAP_SCREENY
    LowY = ScreenY + Y - 1
    For cThings = 0 To (OBJ_MAXTHINGSARRAY - 1)
        If (Things(cThings).x < HighX) And (Things(cThings).x > LowX) And (Things(cThings).Y < HighY) And (Things(cThings).Y > LowY) Then
            Iml.ListImages(Things(cThings).Type).Draw PicBox.hDC, (Things(cThings).x - ScreenX - x) * MAP_TILEXSIZE, (Things(cThings).Y - ScreenY - Y) * MAP_TILEYSIZE, imlTransparent
        End If
    Next cThings

End Sub
Public Function TypeOfThing(ArrayNum As Integer) As Integer 'return numeric constant of
'type of thing
    If Things(ArrayNum).Type < PERSON Then 'it's an object
        TypeOfThing = OBJ
    ElseIf Things(ArrayNum).Type >= PERSON And Things(ArrayNum).Type < MONSTER Then
        TypeOfThing = PERSON
    ElseIf Things(ArrayNum).Type >= MONSTER Then
        TypeOfThing = MONSTER
    End If
End Function
Public Function MoveThing(ArrayNum As Integer, x As Integer, Y As Integer) As Boolean
'returns false if cannot move thing off screen.
Dim Temp As Thing
Dim ScreenX As Integer, ScreenY As Integer
    ScreenX = Things(ArrayNum).x \ MAP_SCREENX
    ScreenY = Things(ArrayNum).Y \ MAP_SCREENY
    Temp = Things(ArrayNum)
    Temp.x = x
    Temp.Y = Y
    If (Temp.x \ MAP_SCREENX) = ScreenX And (Temp.Y \ MAP_SCREENY) = ScreenY Then
    'see if we're still on'screen' if we aren't you must call RemoveThing,
    'then PutThing in the new screen(this error checking so that:
    '   1.we don't have 'Things' from one screen on another's,
    '   and 2. we have automatic bounds checking for randomly moving stuff.
    'this makes it very easy to simply generate a random movement for a Person.
    'then this function will return false if the random person tries to move
    'off the screen. (and it won't do anything). Presto! All we have to do
    'from there is make sure the random 'Person' doesn't walk through walls!
    '(for scripts, we'll have to make sure we do the Remove, Add thing when
    'the script hits the edge of a screen)
        Things(ArrayNum) = Temp
        MoveThing = True
    Else
        MoveThing = False
    End If
End Function
Public Function MoveThingTo(ArrayNum As Integer, Optional IncX As Long = 0, Optional IncY As Long = 0) As Boolean 'false if cannot move thing off
'screen
Dim Temp As Thing
Dim ScreenX As Integer, ScreenY As Integer
    ScreenX = Things(ArrayNum).x \ MAP_SCREENX
    ScreenY = Things(ArrayNum).Y \ MAP_SCREENY
    Temp = Things(ArrayNum)
    Temp.x = Temp.x + IncX
    Temp.Y = Temp.Y + IncY
    If (Temp.x \ MAP_SCREENX) = ScreenX And (Temp.Y \ MAP_SCREENY) = ScreenY Then
        MoveThingTo = True
        Things(ArrayNum) = Temp
    Else
        MoveThingTo = False
    End If
End Function
Public Sub RestoreThings(ScreenX As Long, Optional ScreenY As Long)
Dim i As Integer
    For i = 0 To OBJ_MAXTHINGSARRAY - 1 Step 1
        If Things(i).Y <> NONE And Things(i).x = NONE Then  'if it's(probably) an already deleted
        'thing...
            Select Case i   'figure out which screen it came from and put it back there.
                'column 1
                Case 0 To MAP_SCREENX - 1, _
                (MAP_NUMSCREENSX * MAP_SCREENX) To (MAP_NUMSCREENSX * MAP_SCREENX) + (MAP_SCREENX - 1), _
                (2 * (MAP_NUMSCREENSX * MAP_SCREENX)) To (2 * (MAP_NUMSCREENSX * MAP_SCREENX)) + (MAP_SCREENX - 1)
                    Things(i).x = ScreenX + Int(i Mod 9)
                'column 2
                Case MAP_SCREENX To (2 * MAP_SCREENX) - 1, _
                (MAP_NUMSCREENSX * MAP_SCREENX) + MAP_SCREENX To (MAP_NUMSCREENSX * MAP_SCREENX) + MAP_SCREENX + (MAP_SCREENX - 1), _
                (2 * (MAP_NUMSCREENSX * MAP_SCREENX)) + MAP_SCREENX To (2 * (MAP_NUMSCREENSX * MAP_SCREENX)) + MAP_SCREENX + (MAP_SCREENX - 1)
                    Things(i).x = ScreenX + MAP_SCREENX + Int(i Mod 9)
                'column 3
                Case 2 * MAP_SCREENX To (3 * MAP_SCREENX) - 1, _
                (MAP_NUMSCREENSX * MAP_SCREENX) + (2 * MAP_SCREENX) To (MAP_NUMSCREENSX * MAP_SCREENX) + (2 * MAP_SCREENX) + (MAP_SCREENX - 1), _
                (2 * (MAP_NUMSCREENSX * MAP_SCREENX)) + (2 * MAP_SCREENX) To (2 * (MAP_NUMSCREENSX * MAP_SCREENX)) + (2 * MAP_SCREENX) + (MAP_SCREENX - 1)
                    Things(i).x = ScreenX + (MAP_SCREENX * 2) + Int(i Mod 9)
            End Select
        End If
    Next i
End Sub

Public Sub CleanThingArray()
'this function 'cleans' the thing array--it takes all Things whose X is -1(the invisible flag) and Y is something else and
'resets the values of the whole Thing. That way the Restore no longer works and we've got rid of some nasty screen-too-full problems.
'Don't use this too much because it spoils people's fun :).
Dim i As Integer
    For i = 0 To (OBJ_MAXTHINGSARRAY - 1) Step 1
    With Things(i)
        If .x = NONE And .Y <> NONE Then
            'clean it!!
            .Desc = NONE
            .Movement = STILL   'still--the default
            .Type = NONE
            '.x is already -1 !
            .Y = NONE
        End If
    End With
    Next i
End Sub
