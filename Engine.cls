VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IsRunning As Boolean
Private Const FrameRate = 16
Private Sprites() As SpriteBank
Private Rendering As RenderingEngine
Private Keyboard As KeyboardInput
Private DeviceContextHandler As Long
Private Boundary As Rectangle
Private State As GameState

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Initialize(RenderIn As RenderingEngine, BoundaryIn As Rectangle, StateIn As GameState)
    Set Rendering = RenderIn
    Set Boundary = BoundaryIn
    Set State = StateIn
    State.Initialize Rendering
    Rendering.Initialize Boundary

    Set Keyboard = New KeyboardInput
End Sub

Public Sub run(source As Form)
    ' Game loop tutorial at https://gameprogrammingpatterns.com/game-loop.html
    Dim T1 As Long, T2 As Long, Lag As Long, Elapsed As Long
    IsRunning = True
    T1 = GetTickCount
    Lag = 0
    Do
        DoEvents
        T2 = GetTickCount
        Elapsed = T2 - T1
        T1 = T2
        Lag = Lag + Elapsed
        State.Listen KeyboardEvent
        Do
            DoEvents
            Lag = Lag - FrameRate
        Loop Until Lag < FrameRate Or IsRunning = False
        State.Update T2
        ' Draw
        Rendering.Cls
        State.Draw
        Rendering.DrawToScreen source.hDC, Boundary.Width(), Boundary.Height()
        source.Refresh
        Set State = State.State
    Loop Until IsRunning = False
    Rendering.ClearObjects (source.hDC)
End Sub
Public Sub terminate()
    IsRunning = False
End Sub

Public Function getRenderingEngine() As RenderingEngine
    Set getRenderingEngine = Rendering
End Function

Public Function getBoundary() As Rectangle
    Set getBoundary = Boundary
End Function

Private Function KeyboardEvent() As DomainEvent
    Dim KeyEvent As DomainEvent
    Set KeyEvent = New DomainEvent
    
    KeyEvent.setKeyCode Keyboard.getKeyCodes
    
    KeyEvent.Initialize "keydown", Me
    Set KeyboardEvent = KeyEvent
End Function
