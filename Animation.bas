Attribute VB_Name = "Animation"
Type CollisionRect
    Item As String * 6 'The type of Collision `cls` or `attack`
    Rect As Rectangle ' The rectangle of Collision
    OffsetX As Integer ' The x distance of collision rect from the rectangle of Sprite
    OffsetY As Integer ' The x distance of collision rect from the rectangle of Sprite
End Type

Type Frame
    Index As Integer
    Delay As Integer
    Collisions() As CollisionItem
    Attacks() As CollisionItem
    hDC As Long
End Type

Type ActionItem
    Name As String
    Frames() As Frame
    Loop As Boolean
    Wait As Boolean
End Type


