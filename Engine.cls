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
Private Const MsDelay = 16
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
    Dim T1 As Long, T2 As Long, FrameCount As Long, T3 As Long
    FrameCount = 0
    IsRunning = True
    T1 = GetTickCount
    'T3 = GetTickCount
    Do
        DoEvents
        T2 = GetTickCount
        If T2 - T1 >= MsDelay Then
            State.Listen KeyboardEvent
            State.Update T2
            ' Draw
            Rendering.Cls
            State.Draw
            Rendering.DrawToScreen source.hDC, Boundary.Width(), Boundary.Height()
            source.Refresh
            Set State = State.State
            'FrameCount = FrameCount + 1
        End If
        'If (GetTickCount - T3) = 1000 Then
        '    Debug.Print FrameCount
        '    FrameCount = 0
        '    T3 = GetTickCount
        'End If
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
