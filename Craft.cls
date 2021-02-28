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
Private Actions(4) As Action
Private ActionIndex As Integer

Public Sub Sprite_Update(Time As Long)
    Dim dstX As Long
    Dim dstY As Long
    
    If GameEvent.Name = "keydown" Then
        Dim KeyCodes() As Integer
        KeyCodes = GameEvent.getKeyCode
        'ActionIndex = 0
        If KeyCodes(0) = 37 Then ' Left
            dstX = dstX - Step
            ActionIndex = 0
        End If
        
        If KeyCodes(1) = 39 Then ' Right
            dstX = dstX + Step
            ActionIndex = 0
        End If
        
        If KeyCodes(2) = 1 Then ' Up
            dstY = dstY - Step
            ActionIndex = 1
        End If
        
        If KeyCodes(3) = 2 Then ' Down
            dstY = dstY + Step
            ActionIndex = 2
        End If
        
        If KeyCodes(0) = 0 And KeyCodes(1) = 0 And KeyCodes(2) = 0 And KeyCodes(3) = 0 Then
            If ActionIndex = 1 Or ActionIndex = 3 Then
                ActionIndex = 3
            ElseIf ActionIndex = 2 Or ActionIndex = 4 Then
                ActionIndex = 4
            Else
                ActionIndex = 0
            End If
            
        End If
    End If
    For i = 0 To UBound(Actions)
        If ActionIndex <> i Then
            Actions(i).Reset
        End If
    Next
    
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
    
    pBank.Bank = Actions(ActionIndex).NextFrame
End Sub

Public Sub Sprite_Init(ByRef Bank As SpriteBank)
    Set pBank = Bank
    Set pRect = New Rectangle
    pRect.Init 0, 0, pBank.Width, pBank.Height
    pBank.Bank = 0
    ActionIndex = 0
    
    
    ' Initialize Animation Actions
    Dim Frames_1(0) As Integer
    Dim Frames_2(2) As Integer, Frames_3(2) As Integer
    Dim Frames_4(2) As Integer, Frames_5(2) As Integer
    
    Frames_1(0) = 0
    Set Actions(0) = New Action
    Actions(0).Init Frames_1, 10
    
    ' Go Up Action
    Frames_2(0) = 0
    Frames_2(1) = 1
    Frames_2(2) = 2
    Set Actions(1) = New Action
    Actions(1).Init Frames_2, 10
    
    ' Go Down Action
    Frames_3(0) = 0
    Frames_3(1) = 3
    Frames_3(2) = 4
    Set Actions(2) = New Action
    Actions(2).Init Frames_3, 10
    
    ' Restore Up Action
    Frames_4(0) = 2
    Frames_4(1) = 1
    Frames_4(2) = 0
    Set Actions(3) = New Action
    Actions(3).Init Frames_4, 5
    
    ' Restore Down Action
    Frames_5(0) = 4
    Frames_5(1) = 3
    Frames_5(2) = 0
    Set Actions(4) = New Action
    Actions(4).Init Frames_5, 5
    
    
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
