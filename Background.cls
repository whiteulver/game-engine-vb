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
Private pBank As SpriteBank
Private pRect As Rectangle
Private GameEvent As DomainEvent

Public Sub Sprite_Init(ByRef Bank As SpriteBank)
    Set pBank = Bank
    Set pRect = New Rectangle
    pRect.Init 0, 0, pBank.Width, pBank.Height
End Sub

Public Sub Sprite_Update(Time As Long)
End Sub

Public Property Get Sprite_Bank() As SpriteBank
    Set Sprite_Bank = pBank
End Property

Friend Property Let Sprite_Bank(ByRef BankIn As SpriteBank)
    pBank = BankIn
End Property

Friend Property Let Sprite_Rect(ByRef RectIn As Rectangle)
    pRect = RectIn
End Property

Public Property Get Sprite_Rect() As Rectangle
    Set Sprite_Rect = pRect
End Property

Public Sub Sprite_Listen(GameEventIn As DomainEvent)
    Set GameEvent = GameEventIn
End Sub
