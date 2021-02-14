VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoveListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Listener
Private sprite As SpriteBank
Private dstX As Long
Private dstY As Long
Private Const Step As Integer = 5

Public Sub initialize(spriteIn As SpriteBank)
    Set sprite = spriteIn
    dstX = (512 / 2) - (sprite.Width() / 2)
    dstY = 200 - sprite.Height

End Sub

Public Sub Listener_execute(e As DomainEvent)
    If e.getName = "keydown" Then
        Dim KeyCodes() As Integer
        KeyCodes = e.getKeyCode
        
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
        
        If dstX < 0 Then
            dstX = 0
        End If
        
        If dstX > e.getEngine().getBoundary().Width() - sprite.Width Then
            dstX = e.getEngine().getBoundary().Width() - sprite.Width
        End If
        
        If dstY < 0 Then
            dstY = 0
        End If
        
        If dstY > e.getEngine().getBoundary().Height() - sprite.Height Then
            dstY = e.getEngine().getBoundary().Height() - sprite.Height
        End If

    End If
    
    e.getEngine.getRenderingEngine.draw sprite, dstX, dstY
End Sub

