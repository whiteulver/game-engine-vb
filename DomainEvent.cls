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
Private name As String
Private GameEngine As Engine
Private KeyCode() As Integer

Public Sub initialize(EventName As String, e As Engine)
    name = EventName
    Set GameEngine = e
End Sub

Public Function getName() As String
    getName = name
End Function

Public Function getEngine() As Engine
    Set getEngine = GameEngine
End Function

Public Sub setKeyCode(KeyCodeIn() As Integer)
    KeyCode = KeyCodeIn
End Sub

Public Function getKeyCode() As Integer()
    getKeyCode = KeyCode
End Function
