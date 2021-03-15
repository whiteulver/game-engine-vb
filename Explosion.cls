VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Explosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Const Registry = 3
Private ExplodeFrames(16) As Frame
Private Actions(0) As Action
Private CurrentAction As Integer
Private CurrentFrame As Integer
Private Const Height As Integer = 32 ' Height of single sprite image
Private Const Width As Integer = 32 ' Width of single sprite image
Private pRect As Rectangle

Private Sub Class_Initialize()
    CurrentAction = 0
    Set pRect = New Rectangle
    pRect.Init 0, 0, Width, Height
    
    Dim Frames_6(15) As Integer
    ' Set up explosion animation
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
    Set Actions(0) = New Action
    Actions(0).Init Frames_6, 3
    Dim i As Integer, Delay(16) As Integer
    Delay(0) = 1
    Delay(1) = 1
    Delay(2) = 1
    Delay(3) = 1
    Delay(4) = 1
    Delay(5) = 1
    Delay(6) = 1
    Delay(7) = 1
    Delay(8) = 1
    Delay(9) = 1
    Delay(10) = 1
    Delay(11) = 1
    Delay(12) = 1
    Delay(13) = 1
    Delay(14) = 1
    Delay(15) = 1
    Delay(16) = 1
    For i = 0 To UBound(ExplodeFrames)
        Set ExplodeFrames(i) = New Frame
        If i < 10 Then
            ExplodeFrames(i).Init 32 * i, 0, 32, 32, 0, 0, 32, 32, Delay(i)
        Else
            ExplodeFrames(i).Init 32 * (i - 10), 32, 32, 32, 0, 0, 32, 32, Delay(i)
        End If
    Next
End Sub

Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long)
    ' Set the frame of the action as the current frame to draw
    CurrentFrame = Actions(CurrentAction).NextFrame
    ExplodeFrames(CurrentFrame).CollisionRect.Align pRect
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    Render.Draw Registry, ExplodeFrames(CurrentFrame).DrawRect, pRect
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = pRect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
End Sub

Public Sub Sprite_Destroy()
End Sub
Public Function Sprite_IsAlive() As Boolean
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = Registry
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
End Sub
