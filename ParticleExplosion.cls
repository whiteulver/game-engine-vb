VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParticleExplosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Alive As Boolean
Private Bounds As Rectangle
Private mRegistry As Integer
Private pParticleLogic As ParticleLogic

Public Sub Class_Initialize()
    Alive = True
    Set Bounds = New Rectangle
End Sub

Public Sub Sprite_Init()
    Set pParticleLogic = New ParticleLogic
    pParticleLogic.Init 0.02, 0.4, 20, Bounds, 40
End Sub

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    pParticleLogic.Update Time, Delta
    Alive = pParticleLogic.Alive
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    pParticleLogic.Draw Render
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = Bounds
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    ' Explosion does not handle external events
End Sub

Public Sub Sprite_Destroy()
    pParticleLogic.Destroy
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = Alive
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = mRegistry
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
    ' Explosions don't collide
End Sub
Public Sub Sprite_Collided(ByRef Other As Sprite)
End Sub
Public Function Sprite_ShouldRecycle() As Boolean
    Sprite_ShouldRecycle = Not Alive
End Function
