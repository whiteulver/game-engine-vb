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
Option Explicit
Implements Sprite

Private Speed As Integer
Private Rect As Rectangle ' The rectal that defines the position of sprite in screen.
Private srcRect As Rectangle ' The rectangle from asteroid sprite in image file.
Private Type AsteroidType
    srcX As Long
    srcY As Long
    Width As Long
    Height As Long
End Type
Private Variants(1) As AsteroidType
Private Alive As Boolean
Private Explosion As Sprite

Private Sub Class_Initialize()
    Alive = True
    Set Explosion = New Explosion
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

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    If Alive = False Then
        Explosion.Update Time, Delta
        Exit Sub
    End If
    Rect.Move -Speed, 0
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    If Alive = True Then
        Render.Draw 2, srcRect, Rect
    Else
        Explosion.Draw Render
    End If
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = Rect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
End Sub

Public Sub Sprite_Destroy()
    If Alive = False Then
        Exit Sub
    End If
    Alive = False
    Explosion.Rect.Align Rect
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = Alive
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = 2
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
End Sub
