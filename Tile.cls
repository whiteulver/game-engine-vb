VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDestRect As Rectangle
Private pSrcRect As Rectangle
Private pType As String

Public Property Get DestRect() As Rectangle
    Set DestRect = pDestRect
End Property

Public Property Set DestRect(ByVal RHS As Rectangle)
    Set pDestRect = RHS
End Property

Public Property Get SrcRect() As Rectangle
    Set SrcRect = pSrcRect
End Property

Public Property Set SrcRect(ByVal RHS As Rectangle)
    Set pSrcRect = RHS
End Property

Public Property Get TileType() As String
    TileType = pType
End Property

Public Property Let TileType(ByVal RHS As String)
    pType = RHS
End Property

