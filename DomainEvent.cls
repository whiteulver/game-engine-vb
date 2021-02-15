VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DomainEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private GameEngine As Engine
Private KeyCode() As Integer

Public Sub Initialize(EventName As String, e As Engine)
    pName = EventName
    Set GameEngine = e
End Sub

Public Property Get Name() As String
    Name = pName
End Property

Public Function getEngine() As Engine
    Set getEngine = GameEngine
End Function

Public Sub setKeyCode(KeyCodeIn() As Integer)
    KeyCode = KeyCodeIn
End Sub

Public Function getKeyCode() As Integer()
    getKeyCode = KeyCode
End Function
