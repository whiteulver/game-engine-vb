VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Name As String
Private element() As DomainEvent

Public Sub Init(nameIn As String)
    Name = nameIn
End Sub
Public Sub add(e As DomainEvent)
    Dim index As Integer
    If IsEmpty(element) Then
        index = 0
    Else
        index = UBound(element) + 1
    End If
    
    ReDim Preserve element(index)
    Set element(index) = e
End Sub
