VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextFileLevelCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements LevelCreator

Private Grid() As Tile
Public Sub LevelCreator_Init(ByVal PathIn As String)
    Dim LineText As String
    Dim Lines As New Collection
    Open PathIn For Input As #1
        Do Until EOF(1)
            Line Input #1, LineText
            Lines.add LineText
        Loop
    Close #1
    If Lines.Count = 0 Then Exit Sub
    
    Dim Cols As Long
    Cols = Len(Lines(1))
    ReDim Grid(1 To Lines.Count, 1 To Cols)
     
    Dim Row As Long, Col As Long
    Dim X As Long, Y As Long
    For Row = 1 To Lines.Count
        LineText = Lines(Row)
        Y = (Row - 1) * 16
        For Col = 1 To Cols
            Dim t As Tile
            Set t = New Tile
            t.TileType = Mid$(LineText, Col, 1)
            X = (Col - 1) * 16
            
            Dim DestRect As Rectangle
            Set DestRect = New Rectangle
            DestRect.Init X, Y, 16, 16
            Set t.DestRect = DestRect
            
            ' Populate srcRect for solid tiles
            If t.TileType = "X" Then
                Dim srcRect As Rectangle
                Set srcRect = New Rectangle
                srcRect.Init 0, 0, 16, 16
                Set t.SrcRect = srcRect
            End If
            
            Set Grid(Row, Col) = t
        Next Col
    Next Row
End Sub

' Public method to get the Grid array (classes can be returned directly)
Public Function LevelCreator_GetGrid() As Tile()
    LevelCreator_GetGrid = Grid
End Function
