VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Background"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Sprite

Private Const Width As Integer = 512
Private Const Height As Integer = 256
Private Const Registry As Integer = 0
Private ScrollStep As Long
Private pRect As Rectangle
Private GameEvent As DomainEvent
Private ScrollX As Long
Private ViewRectLeft As New Rectangle
Private ViewRectRight As New Rectangle
Private SourceRectLeft As New Rectangle
Private SourceRectRight As New Rectangle

Private Sub Class_Initialize()
    ScrollStep = 1
    ScrollX = 0
    Set pRect = New Rectangle
    pRect.Init 0, 0, Width, Height
End Sub

Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long)
    ScrollX = ScrollX + ScrollStep
    If ScrollX > pRect.Width Then
        ScrollX = 0
    End If
    
    ViewRectLeft.Init 0, 0, Width - ScrollX, Height
    ViewRectRight.Init Width - ScrollX, 0, ScrollX, Height
    SourceRectLeft.Init ScrollX, 0, ViewRectLeft.Width, Height
    SourceRectRight.Init 0, 0, ViewRectRight.Width, Height
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = pRect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    Render.Draw 0, SourceRectLeft, ViewRectLeft
    Render.Draw 0, SourceRectRight, ViewRectRight
End Sub

Public Sub Sprite_Destroy()
End Sub

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = Registry
End Property
