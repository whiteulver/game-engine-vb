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
Public Sub Initialize(ByRef Bank As SpriteBank)
End Sub

Public Sub Update(Time As Long)
End Sub

Public Property Get Bank() As SpriteBank
End Property

Friend Property Let Bank(ByRef BankIn As SpriteBank)
End Property

Friend Property Let Rect(ByRef RectIn As Rectangle)
End Property

Public Property Get Rect() As Rectangle
End Property

Public Sub Listen(GameEventIn As DomainEvent)
End Sub
