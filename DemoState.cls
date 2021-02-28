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


Public Sub GameState_Init(ByRef RenderIn As RenderingEngine)

    Set Render = RenderIn
    
    ' Load Background
    Set bgnSprite = New Background
    Dim bgnBank As SpriteBank
    Set bgnBank = New BackgroundSprite
    Render.RegisterSprite 0, bgnBank.getFilename, bgnBank.getMaskFilename
    bgnSprite.Init bgnBank

    ' Load Craft
    Set cftSprite = New Craft
    Dim cftBank As SpriteBank
    Set cftBank = New CraftSprite
    Render.RegisterSprite 1, cftBank.getFilename, cftBank.getMaskFilename
    cftSprite.Init cftBank
    
End Sub


Public Function GameState_State() As GameState
    Set GameState_State = Me
End Function


Public Sub GameState_Update(Time As Long)
    cftSprite.Listen GameEvent
    cftSprite.Update Time
End Sub

Public Sub GameState_Draw()
    Render.DrawSprite 0, bgnSprite.Bank.Rect(), bgnSprite.Rect
    Render.DrawSprite 1, cftSprite.Bank.Rect(), cftSprite.Rect
End Sub
Public Sub GameState_Listen(GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub
