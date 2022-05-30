Option Explicit
'This is my proposed type to track all of the things for
'a character in a game.
Type tChar
    'This is to tell whether or not an enemy is active
    Active As Integer
    Health As Integer
    MaxHealth As Integer
    Sprite As tSprite 'This contains all of the information
    'concerning the work, image, and mask picboxed of the
    'character.
    Weapons As Integer 'A list of all the weapons collected
    Items As Integer 'A list of all collected items
    CurWeapon As Integer
    CurItem As Integer
    Name As String 'Used to identify what type of monster
    'is being animated.
    XDir As Integer 'We might use these for monsters, and
    YDir As Integer 'we might not. I'm not sure.
End Type

