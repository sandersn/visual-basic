Attribute VB_Name = "Explore"
Option Explicit
'(11/08) Changed PaintMap to only paint the map, not do edge clip logic as well. Now
'a sub inside frmMapEdit handles that called PaintPicViewPort. It also calls all other
'Paintxxx functions.
Public Opener As String
Public MapXSize As Integer
Public MapYSize As Integer
'Painting Declares
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
'OLD: this function declare and accompanying constant are unused currently, but may be used in the
'real game. NEW: well, I use these now!
Public Const SRCCOPY = &HCC0020

'*** Constants ***
Public Const MAP_BLANK As Integer = -1
Public Const MAP_SCREENX As Integer = 10
Public Const MAP_SCREENY As Integer = 10
Public Const MAP_NUMSCREENSX As Integer = 3
Public Const MAP_NUMSCREENSY As Integer = 3
Public Const MAP_ARRAYX As Integer = MAP_SCREENX * MAP_NUMSCREENSX
Public Const MAP_ARRAYY As Integer = MAP_SCREENY * MAP_NUMSCREENSY
Public Const MAP_ARRAYTOTAL As Integer = (MAP_ARRAYX * MAP_ARRAYY) - 1

Public Const MAP_TILEXSIZE As Integer = 32
Public Const MAP_TILEYSIZE As Integer = 32
'*** End Constants ***

Public Map(0 To MAP_ARRAYTOTAL) As Integer  'newsflash: when you declare an array as (900), VB gives you
'901 elements, 0 to 900!(contrary to popular belief that it gives you 1 to 900 or
'(like C) 0 to 899 when you say (900).

'Anyway, we weren't using the 900th element when it was declared as (900), so I changed
'it to (899)

'Some of my comments in this module look better when viewed in
'the default multi-procedure view.
Public Sub PaintMapFast(PicBox As PictureBox, Iml As ImageList, x As Integer, Y As Integer)
'This sub Paints the Map on the PicBox passed to it, using the Map() array, with images from the ImageList passed to
'it.    X and Y are called ScreenX, ScreenY inside all programs currently using this sub.
'NOTE: This version does NOT support blank tiles(i.e. -1). You can still use blank tiles in the Map Editor, but to
'be painted with this function, you must have explicitly drawn it.

'Dim Dummy ' Place to stick the return value of BitBlt(do we have to have this? C Doesn't)
'oops, don't need this(no BitBlt, no return value)
Dim XIndex As Integer, YIndex As Integer
Dim TerrainVal As Integer
Dim TempX As Integer, TempY As Integer

    
'ONLY render the viewport
    With Iml
    For YIndex = 0 To (MAP_SCREENY - 1) 'change to a constant called YVIEWPORT
        For XIndex = 0 To (MAP_SCREENX - 1) 'change to a constant called XVIEWPORT
            TempX = x + XIndex 'figure our current position
            TempY = Y + YIndex
            TerrainVal = Map((TempY * MAP_ARRAYX) + TempX) 'Look up the value in the array
            
            'Dummy = BitBlt(PicBox.hDC, XIndex * 32, YIndex * 32, 32, 32, .ListImages(TerrainVal).Picture, 0, 0, SRCCOPY)
            Iml.ListImages(TerrainVal).Draw PicBox.hDC, XIndex * MAP_TILEXSIZE, YIndex * MAP_TILEYSIZE, imlNormal
        Next XIndex
    Next YIndex
    End With

End Sub
Public Sub PaintMap(PicBox As PictureBox, Iml As ImageList, x As Integer, Y As Integer)
'This sub Paints the Map on the PicBox passed to it, using the Map() array, with images from the ImageList passed to
'it.    X and Y are called ScreenX, ScreenY inside all programs currently using this sub.
'*** Old ***
'------------------------------------------------------------
' Build a new map bitmap in memory (using windows.h calls),
' then BitBlt this new bitmap into the on-screen PictureBox.
'------------------------------------------------------------
'(actually this is from WAY back: it's some old code from Black Art of VB3 Programming!
'This sub has changed quite a bit since then...
'actually, we now draw the bitmap directly on the picture box. It increases flicker(on IE4 systems anyway) and speed
'and we don't use Win95 API anymore either

'Dim Dummy ' Place to stick the return value of BitBlt(do we have to have this? C Doesn't)
'oops, don't need this(no BitBlt, no return value)
'*** End Old ***
Dim XIndex As Integer, YIndex As Integer
Dim TerrainVal As Integer
Dim TempX As Integer, TempY As Integer

    
'ONLY render the viewport
    With Iml
    For YIndex = 0 To (MAP_SCREENY - 1) 'change to a constant called YVIEWPORT
        For XIndex = 0 To (MAP_SCREENX - 1) 'change to a constant called XVIEWPORT
            TempX = x + XIndex 'figure our current position
            TempY = Y + YIndex
            TerrainVal = Map((TempY * MAP_ARRAYX) + TempX) 'Look up the value in the array
            
            If TerrainVal > MAP_BLANK Then 'blit a tile
                'Dummy = BitBlt(PicBox.hDC, XIndex * 32, YIndex * 32, 32, 32, .ListImages(TerrainVal).Picture, 0, 0, SRCCOPY)
                Iml.ListImages(TerrainVal).Draw PicBox.hDC, XIndex * MAP_TILEXSIZE, YIndex * MAP_TILEYSIZE, imlNormal
            ElseIf TerrainVal < 0 Then  'blit a space
                'Dummy = BitBlt(PicBox.hDC, XIndex * 32, YIndex * 32, 32, 32, .ListImages("blank").Picture, 0, 0, SRCCOPY)
                Iml.ListImages("blank").Draw PicBox.hDC, XIndex * MAP_TILEXSIZE, YIndex * MAP_TILEYSIZE, imlNormal
            End If 'other style possibilities include imlTransparent,imlSelected, and imlFocus
        Next XIndex
    Next YIndex
    End With
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
Public Sub LoadMap(Fileno As Integer, x As Long, Y As Long)
Dim XCount As Long, YCount As Long
    For YCount = 1 To MAP_ARRAYY Step 1
        For XCount = 1 To MAP_ARRAYX Step 1
            Get #Fileno, (((YCount + Y - 1) * MapXSize) + (XCount + x)) + 2, _
            Map(((YCount - 1) * 30) + (XCount - 1))
        Next XCount
    Next YCount
    'Note: the +2 accounts for that fact that I am storing my info
    'starting at the third position--the first two are MapXSize and
    'MapYSize
End Sub
Public Sub SaveMap(Fileno As Integer, x As Long, Y As Long)
Dim XCount As Long, YCount As Long
    For YCount = 1 To MAP_ARRAYY Step 1
        For XCount = 1 To MAP_ARRAYX Step 1
            Put #Fileno, (((YCount + Y - 1) * MapXSize) + (XCount + x)) + 2, _
            Map(((YCount - 1) * 30) + (XCount - 1))
        Next XCount
    Next YCount
    'Note: the +2 accounts for that fact that I am storing my info
    'starting at the third position--the first two are MapXSize and
    'MapYSize
End Sub

