VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackgroundListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Listener
Private Bank As SpriteBank


Public Sub Initialize(spriteIn As SpriteBank)
    Set Bank = spriteIn
End Sub

Public Sub Listener_execute(e As DomainEvent)
    e.getEngine.getRenderingEngine.draw Bank, 0, 0
End Sub
