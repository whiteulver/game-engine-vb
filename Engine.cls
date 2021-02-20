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
Private Const FrameRate = 15
Private Sprites() As SpriteBank
Private Rendering As RenderingEngine
Private Keyboard As KeyboardInput
Private DeviceContextHandler As Long
Private Boundary As Rectangle
Private State As GameState

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub initialize(RenderIn As RenderingEngine, BoundaryIn As Rectangle, StateIn As GameState)
    Set Rendering = RenderIn
    Set Boundary = BoundaryIn
    Set State = StateIn
    State.initialize Rendering
    Rendering.initialize Boundary

    Set Keyboard = New KeyboardInput
End Sub

Public Sub run(source As Form)
    Dim T1 As Long, T2 As Long
    
    IsRunning = True
    T1 = GetTickCount
    Do
        DoEvents
        T2 = GetTickCount
        Debug.Print Elapsed
        If (T2 - T1) >= FrameRate Then
            Rendering.Cls
            State.Listen KeyboardEvent
            State.Update T2
            State.Draw
            Rendering.DrawToScreen source.hdc, Boundary.Width(), Boundary.Height()
            source.Refresh
            Set State = State.State
            T1 = GetTickCount
        End If
    Loop Until IsRunning = False
    Rendering.ClearObjects (source.hdc)
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
    
    KeyEvent.initialize "keydown", Me
    Set KeyboardEvent = KeyEvent
End Function
