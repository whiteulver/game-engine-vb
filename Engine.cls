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
Private EventBucket() As DomainEvent
Private Dispatcher As EventDispatcher
Private Keyboard As KeyboardInput
Private DeviceContextHandler As Long
Private Boundary As Rectangle

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub initialize(RenderIn As RenderingEngine, DispatcherIn As EventDispatcher, BoundaryIn As Rectangle)
    Set Rendering = RenderIn
    Set Boundary = BoundaryIn
    Rendering.initialize Boundary
    
    Set Dispatcher = DispatcherIn
    Set Keyboard = New KeyboardInput
End Sub

Public Sub run(source As Form)
    Dim T1 As Long, T2 As Long
    
    IsRunning = True
    T1 = GetTickCount
    Do
        DoEvents
        T2 = GetTickCount
        If (T2 - T1) >= FrameRate Then
            KeyboardInput
            Rendering.Cls 640, 240
            ' source.Cls
            raiseIdleEvent
            DispatchEvents
            Rendering.DrawToScreen source.hdc, 640, 240
            source.Refresh
            T1 = GetTickCount
        End If
    Loop Until IsRunning = False
    Rendering.ClearObjects (source.hdc)
End Sub

Public Sub terminate()
    IsRunning = False
End Sub

Public Sub raise(e As DomainEvent)
    Dim index As Long
    
    If IsEmpty(EventBucket) Then
        index = 0
    Else
        index = UBound(EventBucket) + 1
    End If
    ReDim Preserve EventBucket(index)
    
    Set EventBucket(index) = e
End Sub

Private Sub DispatchEvents()
    If Not IsEmpty(EventBucket) Then
        Dim i As Integer
        For i = 0 To UBound(EventBucket)
            Dispatcher.dispatch EventBucket(i)
        Next
        Erase EventBucket
    End If
End Sub

Public Function getRenderingEngine() As RenderingEngine
    Set getRenderingEngine = Rendering
End Function

Public Function getBoundary() As Rectangle
    Set getBoundary = Boundary
End Function

Private Sub raiseIdleEvent()
    Dim IdleEvent As New DomainEvent
    
    IdleEvent.initialize "idle", Me
    raise IdleEvent
End Sub

Private Sub KeyboardInput()
    Dim KeyEvent As DomainEvent
    Set KeyEvent = New DomainEvent
    
    KeyEvent.setKeyCode Keyboard.getKeyCodes
    
    KeyEvent.initialize "keydown", Me
    Me.raise KeyEvent
End Sub
