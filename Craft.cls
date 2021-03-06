VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Craft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Sprite

Private Registry As Integer
Private Speed As Integer
Private pRect As Rectangle
Private GameEvent As DomainEvent
Private Actions(5) As Action
Private CurrentAction As Integer
Private CurrentFrame As Integer
Private Frames(14) As Rectangle ' Image has total of 14 sprites. Each frame holds the sprite position X,Y of the image
Private ExplodeFrames(16) As Rectangle
Private Const Height As Integer = 24 ' Height of single sprite image
Private Const Width As Integer = 32 ' Width of single sprite image
Private Alive As Boolean

Public Sub Sprite_Update(ByVal Time As Long)
    Dim dstX As Long
    Dim dstY As Long
    
    If Alive = False Then
        CurrentAction = 5
        CurrentFrame = Actions(CurrentAction).NextFrame
        Exit Sub
    End If
    
    If GameEvent.Name = "keydown" Then
        Dim KeyCodes() As Integer
        KeyCodes = GameEvent.getKeyCode
        
        If KeyCodes(0) = 37 Then ' Left
            dstX = dstX - Speed
            CurrentAction = 0
        End If
        
        If KeyCodes(1) = 39 Then ' Right
            dstX = dstX + Speed
            CurrentAction = 0
        End If
        
        If KeyCodes(2) = 1 Then ' Up
            dstY = dstY - Speed
            CurrentAction = 1
        End If
        
        If KeyCodes(3) = 2 Then ' Down
            dstY = dstY + Speed
            CurrentAction = 2
        End If
        
        If KeyCodes(0) = 0 And KeyCodes(1) = 0 And KeyCodes(2) = 0 _
                And KeyCodes(3) = 0 Then
            If CurrentAction = 1 Or CurrentAction = 3 Then
                CurrentAction = 3
            ElseIf CurrentAction = 2 Or CurrentAction = 4 Then
                CurrentAction = 4
            Else
                CurrentAction = 0
            End If
            
        End If
    End If
    ' Reset all actions that are not running
    For i = 0 To UBound(Actions)
        If CurrentAction <> i Then
            Actions(i).Reset
        End If
    Next
    
    ' Move internal Rect
    pRect.Move dstX, dstY
    
    ' Check the screen boundaries and stop motion for the Craft
    If pRect.Left < 0 Then
        pRect.Left = 0
    End If
    
    If pRect.Left > GameEvent.getEngine().getBoundary().Width() - Width Then
        pRect.Left = GameEvent.getEngine().getBoundary().Width() - Width
    End If
    
    If pRect.Top < 0 Then
        pRect.Top = 0
    End If
    
    If pRect.Top > GameEvent.getEngine().getBoundary().Height() - Height Then
        pRect.Top = GameEvent.getEngine().getBoundary().Height() - Height
    End If
    
    ' Set the frame of the action as the current frame to draw
    CurrentFrame = Actions(CurrentAction).NextFrame
End Sub

Private Sub Class_Initialize()
    Registry = 1
    Alive = True
    Set pRect = New Rectangle
    pRect.Init 0, 0, Width, Height
    Speed = 2
    CurrentAction = 0
    
    ' Initialize Animation Actions
    Dim Skin As Integer ' The type of the craft. Available options are 0, 5, 10
    Skin = 0
    Dim Frames_1(0) As Integer
    Dim Frames_2(1) As Integer, Frames_3(1) As Integer
    Dim Frames_4(2) As Integer, Frames_5(2) As Integer
    Dim Frames_6(15) As Integer
    
    Frames_1(0) = Skin
    Set Actions(0) = New Action
    Actions(0).Init Frames_1, 5
    
    ' Go Up Action
    ' Frames_2(0) = Frames_1(0)
    Frames_2(0) = 1 + Skin
    Frames_2(1) = 2 + Skin
    Set Actions(1) = New Action
    Actions(1).Init Frames_2, 5
    
    ' Go Down Action
    ' Frames_3(0) = Frames_1(0)
    Frames_3(0) = 3 + Skin
    Frames_3(1) = 4 + Skin
    Set Actions(2) = New Action
    Actions(2).Init Frames_3, 5
    
    ' Restore Up Action
    Frames_4(0) = 2 + Skin
    Frames_4(1) = 1 + Skin
    Frames_4(2) = Frames_1(0)
    Set Actions(3) = New Action
    Actions(3).Init Frames_4, 5
    
    ' Restore Down Action
    Frames_5(0) = 4 + Skin
    Frames_5(1) = 3 + Skin
    Frames_5(2) = Frames_1(0)
    Set Actions(4) = New Action
    Actions(4).Init Frames_5, 5
    
    ' Set up the geometry of each frame by computing
    ' the X,Y values of sprite images in the source image.
    Dim OffsetX As Integer, OffsetY As Integer
    OffsetX = 0
    OffsetY = 0
    
    For i = 0 To UBound(Frames)
        Set Frames(i) = New Rectangle
        If i = 5 Or i = 10 Then
            OffsetX = 0
            OffsetY = OffsetY + 1
        End If
        Frames(i).Init OffsetX * Width, OffsetY * Height, Width, Height
        OffsetX = OffsetX + 1
    Next
    
    Frames_6(0) = 1
    Frames_6(1) = 2
    Frames_6(2) = 3
    Frames_6(3) = 4
    Frames_6(4) = 5
    Frames_6(5) = 6
    Frames_6(6) = 7
    Frames_6(7) = 8
    Frames_6(8) = 9
    Frames_6(9) = 0
    Frames_6(10) = 11
    Frames_6(11) = 12
    Frames_6(12) = 13
    Frames_6(13) = 14
    Frames_6(14) = 15
    Frames_6(15) = 16
    Set Actions(5) = New Action
    Actions(5).Init Frames_6, 5
    For i = 0 To UBound(ExplodeFrames)
        Set ExplodeFrames(i) = New Rectangle
        If i < 10 Then
            ExplodeFrames(i).Init 32 * i, 0, 32, 32
        Else
            ExplodeFrames(i).Init 32 * (i - 10), 32, 32, 32
        End If
    Next
End Sub

Public Sub Sprite_Init()
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = pRect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    Render.Draw Registry, FrameRect(), pRect
End Sub

Private Function FrameRect() As Rectangle
    If Alive = True Then
        Set FrameRect = Frames(CurrentFrame)
    Else
        Set FrameRect = ExplodeFrames(CurrentFrame)
    End If
End Function

Public Sub Sprite_Destroy()
    If Alive = False Then
        Exit Sub
    End If
    Alive = False
    CurrentFrame = 0
    CurrentAction = 5
    Registry = 3
    pRect.Init pRect.Left, pRect.Top, 32, 32
End Sub
Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = Registry
End Property
