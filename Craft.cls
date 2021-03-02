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

Private Speed As Integer
Private pRect As Rectangle
Private GameEvent As DomainEvent
Private Actions(4) As Action
Private ActionIndex As Integer
Private CurrentFrame As Integer
Private Frames(14) As Frame ' Image has total of 14 sprites
Private Const Height As Integer = 24 ' Height of single sprite image
Private Const Width As Integer = 32 ' Width of single sprite image

Private Type Frame
    X As Long
    Y As Long
End Type
Public Sub Sprite_Update(Time As Long)
    Dim dstX As Long
    Dim dstY As Long
    
    If GameEvent.Name = "keydown" Then
        Dim KeyCodes() As Integer
        KeyCodes = GameEvent.getKeyCode
        
        If KeyCodes(0) = 37 Then ' Left
            dstX = dstX - Speed
            ActionIndex = 0
        End If
        
        If KeyCodes(1) = 39 Then ' Right
            dstX = dstX + Speed
            ActionIndex = 0
        End If
        
        If KeyCodes(2) = 1 Then ' Up
            dstY = dstY - Speed
            ActionIndex = 1
        End If
        
        If KeyCodes(3) = 2 Then ' Down
            dstY = dstY + Speed
            ActionIndex = 2
        End If
        
        If KeyCodes(0) = 0 And KeyCodes(1) = 0 And KeyCodes(2) = 0 _
                And KeyCodes(3) = 0 Then
            If ActionIndex = 1 Or ActionIndex = 3 Then
                ActionIndex = 3
            ElseIf ActionIndex = 2 Or ActionIndex = 4 Then
                ActionIndex = 4
            Else
                ActionIndex = 0
            End If
            
        End If
    End If
    ' Reset all actions that are not running
    For i = 0 To UBound(Actions)
        If ActionIndex <> i Then
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
    CurrentFrame = Actions(ActionIndex).NextFrame
End Sub

Private Sub Class_Initialize()
    Set pRect = New Rectangle
    pRect.Init 0, 0, Width, Height
    Speed = 2
    ActionIndex = 0
    
    ' Initialize Animation Actions
    Dim Skin As Integer ' The type of the craft. Available options are 0, 5, 10
    Skin = 0
    Dim Frames_1(0) As Integer
    Dim Frames_2(2) As Integer, Frames_3(2) As Integer
    Dim Frames_4(2) As Integer, Frames_5(2) As Integer
    
    Frames_1(0) = Skin
    Set Actions(0) = New Action
    Actions(0).Init Frames_1, 5
    
    ' Go Up Action
    Frames_2(0) = Frames_1(0)
    Frames_2(1) = 1 + Skin
    Frames_2(2) = 2 + Skin
    Set Actions(1) = New Action
    Actions(1).Init Frames_2, 5
    
    ' Go Down Action
    Frames_3(0) = Frames_1(0)
    Frames_3(1) = 3 + Skin
    Frames_3(2) = 4 + Skin
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
    
    
    ' Set up geometry of each frame
    Dim OffsetX As Integer, OffsetY As Integer
    OffsetX = 0
    OffsetY = 0
    
    For i = 0 To UBound(Frames)
        Dim tmpFrame As Frame
        If i = 5 Or i = 10 Then
            OffsetX = 0
            OffsetY = OffsetY + 1
        End If
        tmpFrame.X = OffsetX * Width
        tmpFrame.Y = OffsetY * Height
        Frames(i) = tmpFrame
        OffsetX = OffsetX + 1
    Next
End Sub

Public Sub Sprite_Init()
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = pRect
End Property

Public Sub Sprite_Listen(GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Public Sub Sprite_Draw(Render As RenderingEngine)
    Render.Draw 1, FrameRect(), pRect
End Sub

Private Function FrameRect() As Rectangle
    Set FrameRect = New Rectangle
    FrameRect.Init Frames(CurrentFrame).X, Frames(CurrentFrame).Y, Width, Height
End Function
