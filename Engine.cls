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
Option Explicit

Private IsRunning As Boolean
Private Const MsDelay As Long = 14
Private Rendering As RenderingEngine
Private Keyboard As KeyboardInput
Private DeviceContextHandler As Long
Private Boundary As Rectangle
Private State As GameState

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Sub Init(ByRef RenderIn As RenderingEngine, ByVal BoundaryIn As Rectangle, ByVal StateIn As GameState)
    Set Rendering = RenderIn
    Set Boundary = BoundaryIn
    Set State = StateIn
    Randomize Timer
    State.Init Rendering
    Set Keyboard = New KeyboardInput
End Sub

Public Sub run()
    Dim FrameCount As Long, EndTime As Long, Delta As Double
    Dim TStart As Currency, TNow As Currency
    Dim Frequency As Currency
    Dim ElapsedMs As Double
    FrameCount = 0
    IsRunning = True
    EndTime = GetTickCount
    QueryPerformanceFrequency Frequency
    QueryPerformanceCounter TStart
    Do
        DoEvents
        State.Listen KeyboardEvent
        QueryPerformanceCounter TNow
        ElapsedMs = (TNow - TStart) / Frequency * 1000   ' convert to ms
        State.Update TNow / Frequency * 1000, ElapsedMs
        ' Draw
        Rendering.Cls
        State.Draw
        Rendering.DrawToScreen
        Set State = State.State
        FrameCount = FrameCount + 1
        If (ElapsedMs) >= 1000 Then
            ' Debug.Print FrameCount
            FrameCount = 0
            QueryPerformanceCounter TStart  ' reset start time
        End If
        Sleep MsDelay
    Loop Until IsRunning = False
    Rendering.ClearObjects
End Sub
Public Sub StopEngine()
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
    
    KeyEvent.Init "keydown", Me
    Set KeyboardEvent = KeyEvent
End Function
