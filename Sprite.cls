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
Public Sub Init()
End Sub

Public Sub Update(ByVal Time As Long)
End Sub

Public Sub Draw(ByRef Render As RenderingEngine)
End Sub

Public Property Get Rect() As Rectangle
End Property

Public Sub Listen(ByRef GameEventIn As DomainEvent)
End Sub

Public Sub Destroy()
End Sub

Public Property Get Registry() As Integer
End Property

