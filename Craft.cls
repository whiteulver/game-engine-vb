VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Craft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Sprite
Private Const Step As Integer = 2
Private pBank As SpriteBank
Private pRect As Rectangle
Private GameEvent As DomainEvent

Public Sub Sprite_Update(Time As Long)
    Dim dstX As Long
    Dim dstY As Long
    
    If GameEvent.Name = "keydown" Then
        Dim KeyCodes() As Integer
        KeyCodes = GameEvent.getKeyCode
        
        For Each KeyCode In KeyCodes
        
            If KeyCode = 39 Then
                dstX = dstX + Step
            End If
            
            If KeyCode = 37 Then
                dstX = dstX - Step
            End If
            
            If KeyCode = 1 Then
                dstY = dstY - Step
            End If
            
            If KeyCode = 2 Then
                dstY = dstY + Step
            End If
        
        Next KeyCode

    End If
    
    ' Move internal Rect
    pRect.Move dstX, dstY
    
    ' Check boundaries of the Rect
    If pRect.Left < 0 Then
        pRect.Left = 0
    End If
    
    If pRect.Left > GameEvent.getEngine().getBoundary().Width() - pBank.Width Then
        pRect.Left = GameEvent.getEngine().getBoundary().Width() - pBank.Width
    End If
    
    If pRect.Top < 0 Then
        pRect.Top = 0
    End If
    
    If pRect.Top > GameEvent.getEngine().getBoundary().Height() - pBank.Height Then
        pRect.Top = GameEvent.getEngine().getBoundary().Height() - pBank.Height
    End If
    
End Sub

Public Sub Sprite_Initialize(ByRef Bank As SpriteBank)
    Set pBank = Bank
    Set pRect = New Rectangle
    pRect.Initialize 0, 0, pBank.Width, pBank.Height
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
