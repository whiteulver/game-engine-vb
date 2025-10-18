VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayingState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements GameState

Private Render As RenderingEngine
Private GameEvents() As DomainEvent
Private GameEvent As DomainEvent
Private cftSprite As Sprite
Private bgnSprite As Sprite
Private alienSprite As Sprite
Private asteroidSprite As Sprite

Public Sub GameState_Init(ByRef RenderIn As RenderingEngine)

    Set Render = RenderIn
    
    ' Load Background
    Set bgnSprite = New Background
    Render.RegisterSprite bgnSprite.Registry, App.Path & "\resources\fantasy-512-x-256_full.bmp", App.Path & "\resources\fantasy-512-x-256_full.bmp"

    ' Load Craft
    Set cftSprite = New Craft
    Render.RegisterSprite cftSprite.Registry, App.Path & "\resources\pxplayer.bmp", App.Path & "\resources\pxplayer_m.bmp"
    
    'Load Asteroid
    Set asteroidSprite = New Asteroid
    Render.RegisterSprite asteroidSprite.Registry, App.Path & "\resources\ProjectX_SE_level1_sprites.bmp", App.Path & "\resources\ProjectX_SE_level1_sprites_mask.bmp"
    
    ' Register explosion image
    Render.RegisterSprite 3, App.Path & "\resources\ProjectX_SE_sprites1.bmp", App.Path & "\resources\ProjectX_SE_sprites1_mask.bmp"
    
    ' Register bullet image
    Render.RegisterSprite 4, App.Path & "\resources\ProjectX_SE_sprites2.bmp", App.Path & "\resources\ProjectX_SE_sprites2_mask.bmp"
End Sub

Public Function GameState_State() As GameState
    If cftSprite.ShouldRecycle Then
        Dim tmpState As GameState
        Set tmpState = New PlayingState
        tmpState.Init Render
        Set GameState_State = tmpState
    Else
        Set GameState_State = Me
    End If
End Function

Public Sub GameState_Update(Time As Long, Delta As Double)
    bgnSprite.Update Time, Delta
    cftSprite.Listen GameEvent
    cftSprite.Update Time, Delta
    asteroidSprite.Update Time, Delta
    ' Check collision of Craft and Asteroids
    CheckCollide cftSprite, asteroidSprite
End Sub

Public Sub GameState_Draw()
    bgnSprite.Draw Render
    asteroidSprite.Draw Render
    cftSprite.Draw Render
End Sub
Public Sub GameState_Listen(GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Private Sub CheckCollide(ByRef Source As Sprite, ByRef Other As Sprite)
    Source.Collide Other
End Sub
