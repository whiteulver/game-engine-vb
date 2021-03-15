VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Alien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Registry As Integer
Private Speed As Integer
Private Const Height As Integer = 23 ' Height of single sprite image
Private Const Width As Integer = 20 ' Width of single sprite image
Private Rect As Rectangle
Private srcRect As Rectangle

Private Sub Class_Initialize()
    Registry = 2
    Set Rect = New Rectangle
    Rect.Init 320, 50, Width, Height
    Set srcRect = New Rectangle
    srcRect.Init 134, 39, Width, Height ' The source region from image ProjectX_SE_level1_sprites.bmp
    Speed = 2
End Sub

' 134x39
Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long)
    Rect.Move -Speed, 0
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    Render.Draw 2, srcRect, Rect
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Sprite_Rect = Rect
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
