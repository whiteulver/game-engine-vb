VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DemoState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements GameState

Private Render As RenderingEngine
Private GameEvents() As DomainEvent
Private GameEvent As DomainEvent
Private cftSprite As Sprite
Private bgnSprite As Sprite


Public Sub GameState_Initialize(RenderIn As RenderingEngine)

    Set Render = RenderIn
    
    ' Load Background
    Set bgnSprite = New Background
    Dim bgnBank As New BackgroundSprite
    Render.createSpriteContext bgnBank
    bgnSprite.Initialize bgnBank

    ' Load Craft
    Set cftSprite = New Craft
    Dim cftBank As New CraftSprite
    Render.createSpriteContext cftBank
    cftSprite.Initialize cftBank

End Sub


Public Function GameState_State() As GameState
    Set GameState_State = Me
End Function


Public Sub GameState_Update(Time As Long)
    cftSprite.Listen GameEvent
    cftSprite.Update Time
End Sub

Public Sub GameState_Draw()
    Render.Draw bgnSprite
    Render.Draw cftSprite
End Sub
Public Sub GameState_Listen(GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub
