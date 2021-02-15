VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private listeners() As Listener
Private listenersName() As String

Public Sub dispatch(e As DomainEvent)
    If Not IsEmpty(listeners) Then
        Dim i As Long
        For i = 0 To UBound(listeners)
            If listenersName(i) = e.Name Then
                listeners(i).execute e
            End If
        Next
    End If
End Sub

Public Sub addListener(EventName As String, l As Listener)
    Dim index As Integer
    If IsEmpty(listeners) Then
        index = 0
    Else
        index = UBound(listeners) + 1
    End If
    
    ReDim Preserve listeners(index)
    ReDim Preserve listenersName(index)
    
    Set listeners(index) = l
    listenersName(index) = EventName
End Sub
