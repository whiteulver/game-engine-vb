VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Asteroid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Sprite

' 145, 130 30x36
Private Speed As Integer
Private Const Height As Integer = 36 ' Height of single sprite image
Private Const Width As Integer = 30 ' Width of single sprite image
Private Rect As Rectangle
Private srcRect As Rectangle
Private Type AsteroidType
    srcX As Long
    srcY As Long
    Width As Long
    Height As Long
End Type
Private Variants(1) As AsteroidType

Private Sub Class_Initialize()
    Variants(0).srcX = 145
    Variants(0).srcY = 130
    Variants(0).Width = 30
    Variants(0).Height = 36
    
    Variants(1).srcX = 146
    Variants(1).srcY = 170
    Variants(1).Width = 30
    Variants(1).Height = 29
    
    Dim Skin As Integer
    Skin = Int((2 * Rnd) + 1)
    Skin = Skin - 1
    
    Set Rect = New Rectangle
    Rect.Init 320, 50, Variants(Skin).Width, Variants(Skin).Height
    Set srcRect = New Rectangle
    srcRect.Init Variants(Skin).srcX, Variants(Skin).srcY, Variants(Skin).Width, Variants(Skin).Height
    Speed = 2
End Sub

Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long)
    Rect.Move -Speed, 0
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    Render.Draw 2, srcRect, Rect
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = Rect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
End Sub

Public Sub Sprite_Destroy()
End Sub

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = 2
End Property
