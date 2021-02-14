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
Private pLeft As Integer
Private pTop As Integer
Private pHeight As Integer
Private pWidth As Integer

Public Sub Initialize(LeftIn As Integer, TopIn As Integer, WidthIn As Integer, HeightIn As Integer)
    pLeft = LeftIn
    pTop = TopIn
    pWidth = WidthIn
    pHeight = HeightIn
End Sub

Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Get Width() As Integer
    Width = pWidth
End Property

Public Property Get Height() As Integer
    Height = pHeight
End Property

Public Sub Move(LeftIn As Integer, TopIn As Integer)
    pLeft = LeftIn
    pTop = TopIn
End Sub
