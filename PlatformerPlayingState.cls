VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformerPlayingState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements GameState

Private Render As RenderingEngine
Private LevelSprite As Sprite

Public Sub GameState_Init(ByRef RenderIn As RenderingEngine)
    Set Render = RenderIn
    
    ' Create Level sprite (Level handles path internally)
    Set LevelSprite = New Level
    
    ' Register tiles.bmp sprite sheet (registry 6) before creating Level
    Render.RegisterSprite LevelSprite.Registry, App.Path & "\resources\tiles.bmp", App.Path & "\resources\tiles.bmp"
    
    LevelSprite.Init
End Sub

Public Function GameState_State() As GameState
    Set GameState_State = Me
End Function

Public Sub GameState_Update(Time As Long, Delta As Double)
    If Not LevelSprite Is Nothing Then
        LevelSprite.Update Time, Delta
    End If
End Sub

Public Sub GameState_Draw()
    If Not LevelSprite Is Nothing Then
        LevelSprite.Draw Render
    End If
End Sub

Public Sub GameState_Listen(GameEvent As DomainEvent)
    If Not LevelSprite Is Nothing Then
        LevelSprite.Listen GameEvent
    End If
End Sub

