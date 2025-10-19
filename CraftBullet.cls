VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CraftBullet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Alive As Boolean
Private Const Speed = 4
Private RECT As Rectangle ' The rectal that defines the position of sprite in screen.
Private srcRECT As Rectangle ' The rectangle from bullet sprite in image file.
Private GameEvent As DomainEvent
Private pShouldRecycle As Boolean
Private Explosion As Sprite

Private Sub Class_Initialize()
    Set Explosion = New Explosion
    pShouldRecycle = False
    Alive = True
    Set RECT = New Rectangle
    Set srcRECT = New Rectangle
    RECT.Init 0, 0, 11, 6
    srcRECT.Init 0, 5, 11, 6
End Sub

Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    If Alive = False Then
        Explosion.Update Time, Delta
        Exit Sub
    End If
    RECT.Move Speed, 0
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    If Alive = True Then
        Render.Draw Sprite_Registry, srcRECT, RECT
    Else
        Explosion.Draw Render
    End If
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = RECT
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Public Sub Sprite_Destroy()
    If Alive = False Then
        Exit Sub
    End If
    Alive = False
    Explosion.RECT.Align RECT
    Explosion.Init
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = Alive
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = 4
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
End Sub
Public Sub Sprite_Collided(ByRef Other As Sprite)
    Sprite_Destroy
End Sub
Public Function Sprite_ShouldRecycle() As Boolean
    Sprite_ShouldRecycle = Not Alive And Explosion.ShouldRecycle
End Function
