Attribute VB_Name = "Module1"
Option Explicit
Public Opener As String
Public MapXSize As Integer
Public MapYSize As Integer
Public Map(899) As Integer  'newsflash: when you declare an array as (900), VB gives you
'901 elements, 0 to 900!(contrary to popular belief that it gives you 1 to 900 or
'(like C) 0 to 899 when you say (900).

'Anyway, we weren't using the 900th element when it was declared as (900), so I
'it to (899)

'Some of my comments in this module look better when viewed in
'the default multi-procedure view.

Public Sub Paintmap(picViewport As PictureBox, imlTerrain As ImageList, Fileno As Integer, X As Integer, Y As Integer, TopX As Integer, TopY As Integer, ScreenX As Long, ScreenY As Long, CellX As Integer, CellY As Integer)
'MapXSize and MapYSize are already global(and I don't really want to change them back
'and add them as arguments to SaveMap and LoadMap
'------------------------------------------------------------
' Build a new map bitmap in memory (using windows.h calls),
' then BitBlt this new bitmap into the on-screen PictureBox.
'------------------------------------------------------------
'actually, we now draw the bitmap directly on the picture box. It increases flicker(on IE4 systems anyway) and speed
'and we don't use Win95 API anymore either
'Dim Dummy ' Place to stick the return value of BitBlt(do we have to have this? C Doesn't)
'oops, don't need this(no BitBlt, no return value)
Dim XIndex As Integer, YIndex As Integer
Dim TerrainVal As Integer
Dim TempX As Integer, TempY As Integer
Dim bSaved As Boolean

'First move the array and clip it to the edges.
    If TopX = 20 Then
        If ScreenX + 30 = MapXSize And bSaved = False Then  'we're at map edge!
            SaveMap Fileno, ScreenX, ScreenY    'save the array to disk but DO NOT move the array
            'over to next position because it would otherwise go off the edge.
            '(or reset the viewport)
            bSaved = True 'turn on a switch to make sure we don't repeatedly save to disk
            'when moving along the edge of the map(because we don't reset position when
            'moving along edge of map)
        ElseIf ScreenX + 30 = MapXSize And bSaved = True Then
            'do nothing at all(because continually saving degradates! performance.
        Else    'move along now
            TopX = 10   'reset the viewport to center of array
            SaveMap Fileno, ScreenX, ScreenY    'save changes of current position to disk
            ScreenX = ScreenX + 10  'move array over 10 cells to next pos.
            LoadMap Fileno, ScreenX, ScreenY    'load new position into array
            bSaved = False
        End If
    End If
    
    If TopY = 20 Then
        If ScreenY + 30 = MapYSize And bSaved = False Then  'same comments here...
            SaveMap Fileno, ScreenX, ScreenY
            bSaved = True
        ElseIf ScreenY + 30 = MapYSize And bSaved = True Then  'skip
        Else    'we're not at the edge of the screen, so business as usual
            TopY = 10
            SaveMap Fileno, ScreenX, ScreenY
            ScreenY = ScreenY + 10
            LoadMap Fileno, ScreenX, ScreenY
        End If
    End If
    'oops, forgot to add top, left checking(I was really tired last night)
    If TopX = 0 Then
        If ScreenX = 0 And bSaved = False Then 'we're at maps edge(world's end)
            SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
            bSaved = True
        ElseIf ScreenX = 0 And bSaved = True Then
        Else
            TopX = 10
            SaveMap Fileno, ScreenX, ScreenY
            ScreenX = ScreenX - 10
            LoadMap Fileno, ScreenX, ScreenY
            bSaved = False
        End If
    End If
    If TopY = 0 Then
        If ScreenY = 0 And bSaved = False Then 'we're at maps edge(world's end)
            SaveMap Fileno, ScreenX, ScreenY    'save to disk but DO NOT move the array
            bSaved = True
        ElseIf ScreenY = 0 And bSaved = True Then
        Else
            TopY = 10
            SaveMap Fileno, ScreenX, ScreenY
            ScreenY = ScreenY - 10
            LoadMap Fileno, ScreenX, ScreenY
            bSaved = False
        End If
    End If

    
'Now render the viewport
    For YIndex = 0 To 9 'change to a constant called YVIEWPORT
        For XIndex = 0 To 9 'change to a constant called XVIEWPORT
            TempX = X + XIndex 'figure our current position
            TempY = Y + YIndex
'               replace 30 with ARRAY_X_SIZE someday
            TerrainVal = Map((TempY * 30) + TempX) 'Look up the value in the array
            
            If TerrainVal > -1 Then 'blit a tile
                'Dummy = BitBlt(picViewport.hDC, XIndex * 32, YIndex * 32, 32, 32, picTerrain(TerrainVal).hDC, 0, 0, SRCCOPY)
                imlTerrain.ListImages(TerrainVal).Draw picViewport.hDC, XIndex * 32, YIndex * 32, 0
            ElseIf TerrainVal < 0 Then  'blit a space
                'Dummy = BitBlt(picViewport.hDC, XIndex * 32, YIndex * 32, 32, 32, picCanvas.hDC, 0, 0, SRCCOPY)
                imlTerrain.ListImages("blank").Draw picViewport.hDC, XIndex * 32, YIndex * 32, 0
            End If
        Next XIndex
    Next YIndex
End Sub

Public Sub DrawHighLight(PicBox As PictureBox, Left As Integer, Top As Integer, Width As Integer, Height As Integer)
    PicBox.Line (Left, Top)-(Left + Width, Top), vb3DHighlight
    PicBox.Line (Left + Width, Top)-(Left + Width, Top + Height), vb3DShadow
    PicBox.Line (Left + Width, Top + Height)-(Left, Top + Height), vb3DShadow
    PicBox.Line (Left, Top + Height)-(Left, Top), vb3DHighlight
End Sub
Public Sub DrawSelHighLight(PicBox As PictureBox, Left As Integer, Top As Integer, Width As Integer, Height As Integer)
    PicBox.Line (Left, Top)-(Left + Width, Top), vb3DShadow
    PicBox.Line (Left + Width, Top)-(Left + Width, Top + Height), vb3DHighlight
    PicBox.Line (Left + Width, Top + Height)-(Left, Top + Height), vb3DHighlight
    PicBox.Line (Left, Top + Height)-(Left, Top), vb3DShadow

End Sub
Public Sub LoadMap(Fileno As Integer, X As Long, Y As Long)
Dim XCount As Long, YCount As Long
    For YCount = 1 To 30 Step 1
        For XCount = 1 To 30 Step 1
            Get #Fileno, (((YCount + Y - 1) * MapXSize) + (XCount + X)) + 2, _
            Map(((YCount - 1) * 30) + (XCount - 1))
        Next XCount
    Next YCount
    'Note: the +2 accounts for that fact that I am storing my info
    'starting at the third position--the first two are MapXSize and
    'MapYSize
End Sub
Public Sub SaveMap(Fileno As Integer, X As Long, Y As Long)
Dim XCount As Long, YCount As Long
    For YCount = 1 To 30 Step 1
        For XCount = 1 To 30 Step 1
            Put #Fileno, (((YCount + Y - 1) * MapXSize) + (XCount + X)) + 2, _
            Map(((YCount - 1) * 30) + (XCount - 1))
        Next XCount
    Next YCount
End Sub

