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
Private RECT As Rectangle ' The rectal that defines the position of sprite in screen.
Private srcRECT As Rectangle ' The rectangle from Asteroid sprite in image file.
Private Type AsteroidType
    srcX As Long
    srcY As Long
    Width As Long
    Height As Long
End Type
Private Variants(1) As AsteroidType
Private Alive As Boolean
Private Life As Integer
Private Explosion As Sprite
Private Skin As Integer


Private Sub Class_Initialize()
    Alive = True
    Life = 2
    ' Randomize speed slightly
    Speed = 2 + Int(Rnd() * 2) ' Speed between 2-3

    With Variants(0)
        .srcX = 145
        .srcY = 130
        .Width = 30
        .Height = 36
    End With
    
    With Variants(1)
        .srcX = 146
        .srcY = 170
        .Width = 30
        .Height = 29
    End With
    Skin = Int(Rnd() * 2)

    Set RECT = New Rectangle
    RECT.Init 320, 50, Variants(Skin).Width, Variants(Skin).Height
    Set srcRECT = New Rectangle
    srcRECT.Init Variants(Skin).srcX, Variants(Skin).srcY, Variants(Skin).Width, Variants(Skin).Height
End Sub

Public Sub Sprite_Init()
End Sub

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    If Alive = False Then
        Explosion.Update Time, Delta
        Exit Sub
    End If
    RECT.Move -Speed, 0
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    If Alive = True Then
        Render.Draw 2, srcRECT, RECT
    Else
        Explosion.Draw Render
    End If
End Sub

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = RECT
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
End Sub

Public Sub Sprite_Destroy()
    If Alive = False Then
        Exit Sub
    End If
    Alive = False
    Set Explosion = New Explosion
    Explosion.RECT.Align RECT
    Explosion.Init
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = Alive
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = 2
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
End Sub
Public Sub Sprite_Collided(ByRef Other As Sprite)
    If Other.IsAlive = False Then Exit Sub:
    Life = Life - 1
    If Life <= 0 Then Sprite_Destroy:
End Sub
Public Function Sprite_ShouldRecycle() As Boolean
    If Not Alive And Not (Explosion Is Nothing) Then
        Sprite_ShouldRecycle = Explosion.ShouldRecycle
        Exit Function
    End If
    Sprite_ShouldRecycle = False

End Function
