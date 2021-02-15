VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RenderingEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Initialize(BoundaryIn As Rectangle)
End Sub

Public Sub createSpriteContext(Bank As SpriteBank)
End Sub

Public Sub Draw(spr As sprite)
End Sub

Public Sub DrawToScreen(screen As Long, Width As Long, Height As Long)
End Sub

Public Sub Cls()
End Sub

Public Sub ClearObjects(ByRef hdc As Long)
End Sub
