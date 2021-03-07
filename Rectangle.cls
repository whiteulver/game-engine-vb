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
Option Explicit

Private pLeft As Long
Private pTop As Long
Private pHeight As Long
Private pWidth As Long

Public Sub Init(LeftIn As Long, TopIn As Long, WidthIn As Long, HeightIn As Long)
    pLeft = LeftIn
    pTop = TopIn
    pWidth = WidthIn
    pHeight = HeightIn
End Sub

Public Property Get Top() As Long
    Top = pTop
End Property

Public Property Get Left() As Long
    Left = pLeft
End Property

Public Property Get Width() As Long
    Width = pWidth
End Property

Public Property Get Height() As Long
    Height = pHeight
End Property

Public Property Let Top(TopIn As Long)
    pTop = TopIn
End Property

Public Property Let Left(LeftIn As Long)
    pLeft = LeftIn
End Property

Public Sub Move(LeftIn As Long, TopIn As Long)
    pLeft = pLeft + LeftIn
    pTop = pTop + TopIn
End Sub

Public Function Collide(Other As Rectangle) As Boolean
    Collide = False
    If pLeft < Other.Left + Other.Width And _
            pLeft + pWidth > Other.Left And _
            pTop < Other.Top + Other.Height And _
            pTop + pHeight > Other.Top Then
        Collide = True
    End If
End Function
' Align the position of this rectnge according to X, Y from Source rectangle
Public Sub Align(Source As Rectangle)
    pLeft = Source.Left
    pTop = Source.Top
End Sub
