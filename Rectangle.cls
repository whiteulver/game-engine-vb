VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Left As Integer
Private Top As Integer
Private Height As Integer
Private Width As Integer

Public Sub initialize(LeftIn As Integer, TopIn As Integer, WidthIn As Integer, HeightIn As Integer)
    Set Left = LeftIn
    Set Top = TopIn
    Set Width = WidthIn
    Set Height = HeightIn
End Sub

Public Function getTop() As Integer
    getTop = Top
End Function

Public Function getLeft() As Integer
    getLeft = Left
End Function

Public Function getWidth() As Integer
    getWidth = Width
End Function

Public Function getHeight() As Integer
    getHeight = Height
End Function
