VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpriteBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function getFilename() As String
End Function

Public Function getMaskFilename() As String
End Function

Public Sub setContext(context As Long, maskContext As Long)
End Sub

Public Function getContext() As Long
End Function
Public Function getMaskContext() As Long
End Function
Public Function X() As Long
End Function

Public Function Y() As Long
End Function

Public Function Width() As Long
End Function

Public Function Height() As Long
End Function

Public Sub nextBank()
End Sub

