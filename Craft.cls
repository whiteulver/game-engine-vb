VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Craft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Const Registry As Integer = 1
Private Speed As Integer
Private Rect As Rectangle
Private GameEvent As DomainEvent
Private Actions(4) As Action
Private CurrentAction As Integer
Private CurrentFrame As Integer
Private Frames(14) As Frame ' Image has total of 14 sprites. Each frame holds the sprite position X,Y of the image
Private Const Height As Integer = 24 ' Height of single sprite image
Private Const Width As Integer = 32 ' Width of single sprite image
Private Alive As Boolean
Private Explosion As Sprite
Private Bullets As New Collection
Private BulletTick As Long
Private BulletIndex As Integer

Private Sub Class_Initialize()
    ' Define explosion sprite. When Craft is not Alive, delegate the Draw to Explosion sprite.
    Set Explosion = New Explosion
    Alive = True
    BulletTick = 0: BulletIndex = 0
    Set Rect = New Rectangle
    ' Initial position of Sprite, up left corner.
    Rect.Init 0, 0, Width, Height
    ' Initial Speed
    Speed = 2
    ' Current action is the Flying
    CurrentAction = 0
    
    ' Initialize Animation Actions
    Dim Skin As Integer ' The type of the craft. Available options are 0, 5, 10
    Skin = 0
    
    Dim Frames_1(0) As Integer
    Dim Frames_2(1) As Integer, Frames_3(1) As Integer
    Dim Frames_4(2) As Integer, Frames_5(2) As Integer
    
    ' Define the frame indexes for each action.
    Frames_1(0) = Skin
    Set Actions(0) = New Action
    
    ' Flying
    Actions(0).Init Frames_1, 5
    
    ' Flying Up Action
    Frames_2(0) = 1 + Skin
    Frames_2(1) = 2 + Skin
    Set Actions(1) = New Action
    Actions(1).Init Frames_2, 5
    
    ' Flying Down Action
    Frames_3(0) = 3 + Skin
    Frames_3(1) = 4 + Skin
    Set Actions(2) = New Action
    Actions(2).Init Frames_3, 5
    
    ' Restore Flying Up Action
    Frames_4(0) = 2 + Skin
    Frames_4(1) = 1 + Skin
    Frames_4(2) = Frames_1(0)
    Set Actions(3) = New Action
    Actions(3).Init Frames_4, 5
    
    ' Restore Flying Down Action
    Frames_5(0) = 4 + Skin
    Frames_5(1) = 3 + Skin
    Frames_5(2) = Frames_1(0)
    Set Actions(4) = New Action
    Actions(4).Init Frames_5, 5
    
    ' Set up the geometry of each frame to draw by computing
    ' the X,Y values of sprite images in the source image.
    Dim OffsetX As Integer, OffsetY As Integer
    OffsetX = 0
    OffsetY = 0
    
    Dim i As Integer
    For i = 0 To UBound(Frames)
        Set Frames(i) = New Frame
        If i = 5 Or i = 10 Then
            OffsetX = 0
            OffsetY = OffsetY + 1
        End If
        Frames(i).Init OffsetX * Width, OffsetY * Height, Width, Height, 0, 0, Width, 14, 5
        OffsetX = OffsetX + 1
    Next
End Sub

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    Dim dstX As Long
    Dim dstY As Long
    
    If Alive = False Then
        Explosion.Update Time, Delta
        BulletHandling Time, Delta
        Exit Sub
    End If
    
    If GameEvent.Name = "keydown" Then
        Dim KeyCodes() As Integer
        KeyCodes = GameEvent.getKeyCode
        
        If KeyCodes(0) = 37 Then ' Left
            dstX = dstX - Speed
            CurrentAction = 0
        End If
        
        If KeyCodes(1) = 39 Then ' Right
            dstX = dstX + Speed
            CurrentAction = 0
        End If
        
        If KeyCodes(2) = 1 Then ' Up
            dstY = dstY - Speed
            CurrentAction = 1
        End If
        
        If KeyCodes(3) = 2 Then ' Down
            dstY = dstY + Speed
            CurrentAction = 2
        End If
        
        If KeyCodes(4) = 88 Then ' Fire
            Shoot Time, Delta
        End If
        
        If KeyCodes(0) = 0 And KeyCodes(1) = 0 And KeyCodes(2) = 0 _
                And KeyCodes(3) = 0 Then
            If CurrentAction = 1 Or CurrentAction = 3 Then
                CurrentAction = 3
            ElseIf CurrentAction = 2 Or CurrentAction = 4 Then
                CurrentAction = 4
            Else
                CurrentAction = 0
            End If
            
        End If
    End If
    ' Reset all actions that are not running
    Dim i As Integer
    For i = 0 To UBound(Actions)
        If CurrentAction <> i Then
            Actions(i).Reset
        End If
    Next
    
    ' Move internal Rect
    Rect.Move dstX, dstY
    
    ' Check the screen boundaries and stop motion for the Craft
    If Rect.Left < 0 Then
        Rect.Left = 0
    End If
    
    If Rect.Left > GameEvent.getEngine().getBoundary().Width() - Width Then
        Rect.Left = GameEvent.getEngine().getBoundary().Width() - Width
    End If
    
    If Rect.Top < 0 Then
        Rect.Top = 0
    End If
    
    If Rect.Top > GameEvent.getEngine().getBoundary().Height() - Height Then
        Rect.Top = GameEvent.getEngine().getBoundary().Height() - Height
    End If
    
    ' Set the frame of the action as the current frame to draw
    CurrentFrame = Actions(CurrentAction).NextFrame
    Frames(CurrentFrame).CollisionRect.Align Rect
    
    BulletHandling Time, Delta
End Sub
Private Sub BulletHandling(ByVal Time As Long, Delta As Double)
    ' Update position of bullets
    Dim Bullet As Sprite
    For Each Bullet In Bullets
        Bullet.Update Time, Delta
    Next
    ' Remove bullets from array when not visible in screen
    RemoveBullets GameEvent.getEngine().getBoundary().Width
End Sub
Private Sub RemoveBullets(ByVal Width As Long)
    Dim index As Integer: index = 1
    Dim Bullet As Sprite
    For Each Bullet In Bullets
        If Bullet.Rect.Left > Width Or Bullet.IsAlive = False Then
            Bullets.Remove index
        End If
        index = index + 1
    Next
End Sub

Private Sub DrawBullets(ByRef Render As RenderingEngine)
    If IsEmpty(Bullets) Then
        Exit Sub
    End If
    Dim b As Integer, Bullet As Sprite
    For Each Bullet In Bullets
        Bullet.Draw Render
    Next
End Sub
Public Sub Sprite_Init()
End Sub

Public Property Get Sprite_Rect() As Rectangle
    If Alive = True Then
        Set Sprite_Rect = Frames(CurrentFrame).CollisionRect
    Else
        Set Sprite_Rect = Explosion.Rect
    End If
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    If Alive = True Then
        Render.Draw Registry, FrameRect(), Rect
    Else
        Explosion.Draw Render
    End If
    DrawBullets Render
End Sub

Private Function FrameRect() As Rectangle
    Set FrameRect = Frames(CurrentFrame).DrawRect
End Function

Public Sub Sprite_Destroy()
    If Alive = False Then
        Exit Sub
    End If
    Alive = False
    Explosion.Rect.Align Rect
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = Alive
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = Registry
End Property

Private Sub Shoot(ByVal Time As Long, Delta As Double)
    If Time - BulletTick < 300 Then
        Exit Sub
    End If
    BulletTick = Time

    Dim tmp As Sprite
    Set tmp = New CraftBullet
    Bullets.add tmp
    tmp.Rect.Init Rect.Left + 22, Rect.Top + 12, tmp.Rect.Width, tmp.Rect.Height
End Sub

Public Sub Sprite_Collide(ByRef Other As Sprite)
    If Rect.Collide(Other.Rect) And Other.IsAlive Then
        Sprite_Destroy
    End If
    Dim Bullet As Sprite
    For Each Bullet In Bullets
        If Bullet.Rect.Collide(Other.Rect) And Other.IsAlive Then
            Other.Destroy
            Bullet.Destroy
        End If
    Next
End Sub
