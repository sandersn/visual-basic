Attribute VB_Name = "Mapfile"
Option Explicit
'Nathan Sanders

'Current Subs:
'LoadMap(intFileNo as Integer)  'loads map from supplied file number to Map() Mapdata array
'SaveMap(intFileNo as Integer)  'saves Map() Mapdata array to supplied file number
'SaveMapPart(intFileNo As Integer, x As Integer, y As Long, dx As Integer, dy As Integer)
    'saves a part of the Map() specified by the extent(x,y - x+dx,x+dy) to supplied file number

'module started 28 Sep, 1999

'                               1999
    '28 Sep: Am writing type for MapData. VERY incomplete as Ryan didn't give me lots o' params and I don't know how he's going to use
'just height and color(??) besides, I don't know what a 'UDT' array is...unsigned data type or user data type...only Microsoft would call it
'that: struct is too easy and it's not a TLA.
'    Done with that. Module is basically done. I hope Ryan can keep up with my awesome coding speed :)
Private Type RGBColor      'this is private until it messes something up.
    Red As Integer
    Green As Integer
    Blue As Integer
End Type
Public Type MapData
    intHeight As Integer
    RGBColor As RGBColor
    'others to be defined later
End Type


Public Map() As MapData
Public MapX As Integer, MapY As Integer 'the map x,y(duh)

Public Sub LoadMap(intFileNo As Integer)    'may need the Fileno as param...DO need the Fileno as param!!(unless everything's global, ugh)
'intFileno is the file number of the map file(this may not be needed; in which I'll need a strFileName passed.
'opened as Random ^
'WARNING: This sub loads WHOLE file into memory, thus taking up MB upon MB of it!!
Dim cx As Integer, cy As Integer    'these R simple counters instead of i,j(which are more fun :()
Dim mapTemp As MapData
    Get #intFileNo, 1, mapTemp
    MapX = mapTemp.intHeight   'get x,y coords so we'll know how BIG this map is
    Get #intFileNo, , mapTemp   'for those of you who are in the dark about that 2a param, if you leave it out, it just gets the next struct in line
    MapY = mapTemp.intHeight
ReDim Map(1 To MapX, 1 To MapY) As MapData    'there it goes--crashhh
    For cy = 1 To MapY Step 1
        For cx = 1 To MapX Step 1
            Get #intFileNo, , Map(cx)(cy) 'man this is just too easy--maybe I should use the more complicated code from the MapEdit code ;)
            'the hard version of the second param is: (((cy) * x) + (cx)) + 2   heh heh heh don't try it unless this bugs up
        Next cx
    Next cy
    'we are done--woe betide those with 8 MB system RAM...
End Sub
Public Sub SaveMap(intFileNo As Integer)
'intFileno is the file number of the map file(this may not be needed; in which I'll need a strFileName passed.
'opened as Random ^
'WARNING: This sub saves WHOLE file
Dim cx As Integer, cy As Integer    'these R simple counters instead of i,j(which are more fun :()
    Seek #intFileNo, 2  'put position at 2, then start writing from there
    For cy = 1 To MapY Step 1
        For cx = 1 To MapX Step 1
            Put #intFileNo, , Map(cx)(cy) 'man this is just too easy--maybe I should use the more complicated code from the MapEdit code ;)
            'the hard version of the second param is: ((cy * mapX) + cx) + 2   heh heh heh don't try it unless this bugs up
        Next cx
    Next cy

End Sub

Public Sub SaveMapPart(intFileNo As Integer, x As Integer, y As Long, dx As Integer, dy As Integer)
'this sub just saves *part* of the map; x,y are start positions, dx,dy are the width and height; I will change the names if there is demand*
'intFileno is the file number of a file opened as Random
Dim cx As Integer, cy As Integer    'these R simple counters instead of i,j(which are more fun :()

    For cy = y To y + dy Step 1 'yohoho this should look familiar
        For cx = x To x + dx Step 1
            Put #intFileNo, ((cy * MapX) + cx) + 2, Map(cx)(cy) 'hope this works...
        Next cx
    Next cy
'* from someone that actually *codes* on this sub(and hasn't had calculus :)
'note: for those of you who are asleep, the +2 accounts for the two MapX,Y container structs at the
'beginning of the file
End Sub
