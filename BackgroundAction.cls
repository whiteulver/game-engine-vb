VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackgroundAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements action
Private sprite As SpriteBank

Public Sub initialize(spriteIn As SpriteBank)
    Set sprite = spriteIn
End Sub

Public Sub Action_run(e As DomainEvent)
    'sprite.nextBank
    e.getEngine.getRenderingEngine.draw sprite, 0, 0
End Sub
