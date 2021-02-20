Attribute VB_Name = "Animation"
Public Type CollisionRect
    Item As String * 6 ' The type of Collision `cls` or `attack`
    Rect As Rectangle ' The rectangle of Collision
    OffsetX As Integer ' The x distance of collision rect from the rectangle of Sprite
    OffsetY As Integer ' The x distance of collision rect from the rectangle of Sprite
End Type

Public Type Frame
    Id As Integer
    Delay As Integer
    Collisions() As CollisionRect
    Attacks() As CollisionRect
    hDC As Long
End Type

Public Type ActionItem
    Name As String
    Frames() As Frame
    Loop As Boolean
    Wait As Boolean
End Type


