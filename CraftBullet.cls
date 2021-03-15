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
Private Rect As Rectangle ' The rectal that defines the position of sprite in screen.
Private srcRect As Rectangle ' The rectangle from bullet sprite in image file.
Private GameEvent As DomainEvent

Private Sub Class_Initialize()
    Alive = True
    Set Rect = New Rectangle
    Set srcRect = New Rectangle
    Rect.Init 0, 0, 11, 6
    srcRect.Init 0, 5, 11, 6
End Sub

Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long)
    Rect.Move Speed, 0
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    Render.Draw Sprite_Registry, srcRect, Rect
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = Rect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Public Sub Sprite_Destroy()
    Alive = False
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = Alive
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = 4
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
End Sub
