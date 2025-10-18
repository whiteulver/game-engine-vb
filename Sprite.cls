VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Init()
End Sub
' Updates the state of sprite. Including position etc. The position of sprite can be accessed from Rect property
Public Sub Update(ByVal Time As Long, Delta As Double)
End Sub

Public Sub Draw(ByRef Render As RenderingEngine)
End Sub
' Returns a Rectangle that represents the sprite position in screen
Public Property Get Rect() As Rectangle
End Property

Public Sub Listen(ByRef GameEventIn As DomainEvent)
End Sub

Public Sub Destroy()
End Sub

Public Function IsAlive() As Boolean
End Function

Public Property Get Registry() As Integer
End Property

Public Sub Collide(ByRef Other As Sprite)
End Sub

Public Sub Collided(ByRef Other As Sprite)
End Sub

Public Function ShouldRecycle() As Boolean
End Function
