VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Collision As Rectangle ' The rectangle to use when calculating collision detection
Private Draw As Rectangle ' The rectangle to use when bitblit the sprite from source image

Public Sub Init(srcX As Long, srcY As Long, srcWidth As Long, srcHeight As Long, screenX As Long, screenY As Long, Width As Long, Height As Long)
    Set Draw = New Rectangle
    Set Collision = New Rectangle
    Collision.Init screenX, screenY, Width, Height
    Draw.Init srcX, srcY, srcWidth, srcHeight
End Sub

Public Property Get CollisionRect() As Rectangle
    Set CollisionRect = Collision
End Property

Public Property Get DrawRect() As Rectangle
    Set DrawRect = Draw
End Property

