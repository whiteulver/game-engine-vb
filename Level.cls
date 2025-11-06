VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Grid() As Tile
Private LevelCreator As LevelCreator
Private Const Registry As Integer = 6
Private pRect As Rectangle
Private LevelWidth As Long
Private LevelHeight As Long
Private LevelPath As String

Private Sub Class_Initialize()
    LevelPath = App.Path & "\Resources\Level1.txt"
    Set pRect = New Rectangle
End Sub

Public Sub Sprite_Init()
    ' Create and initialize LevelCreator
    Set LevelCreator = New TextFileLevelCreator
    LevelCreator.Init LevelPath
    
    ' Get Grid from LevelCreator
    Grid = LevelCreator.GetGrid()
    
    ' Calculate level dimensions from grid
    Dim Rows As Long, Cols As Long
    On Error Resume Next
    Rows = UBound(Grid, 1)
    Cols = UBound(Grid, 2)
    On Error GoTo 0
    
    If Rows > 0 And Cols > 0 Then
        LevelWidth = Cols * 16
        LevelHeight = Rows * 16
    Else
        LevelWidth = 0
        LevelHeight = 0
    End If
    
    ' Initialize pRect with level dimensions
    pRect.Init 0, 0, LevelWidth, LevelHeight
End Sub

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    ' Empty for now, may add scrolling later
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    ' Draw all tiles by looping through the grid
    If LevelWidth = 0 Or LevelHeight = 0 Then Exit Sub
    
    ' Check if Grid is valid
    Dim Rows As Long, Cols As Long
    On Error Resume Next
    Rows = UBound(Grid, 1)
    Cols = UBound(Grid, 2)
    On Error GoTo 0
    If Rows <= 0 Or Cols <= 0 Then Exit Sub
    
    Dim Row As Long, Col As Long
    Dim t As Tile
    
    ' Loop through Grid array and draw each tile
    For Row = 1 To Rows
        For Col = 1 To Cols
            Set t = Grid(Row, Col)
            ' Only render non-empty tiles
            If t.TileType = "X" And Not t.SrcRect Is Nothing And Not t.DestRect Is Nothing Then
                ' Draw tile using the renderer
                Render.Draw Registry, t.SrcRect, t.DestRect
            End If
        Next Col
    Next Row
End Sub

Public Property Get Sprite_Rectangle() As Rectangle
    Set Sprite_Rectangle = pRect
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    ' Empty for now
End Sub

Public Sub Sprite_Destroy()
    ' Cleanup if needed
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = True
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = Registry
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
    ' Optimized collision: only check tiles near the actor
    Dim otherRect As Rectangle
    Set otherRect = Other.Rectangle
    
    ' Calculate which grid cells the actor overlaps
    Dim minCellX As Long, maxCellX As Long
    Dim minCellY As Long, maxCellY As Long
    
    ' Convert rectangle coordinates to grid indices (1-based)
    minCellX = (otherRect.Left \ 16) + 1
    maxCellX = ((otherRect.Left + otherRect.Width - 1) \ 16) + 1
    minCellY = (otherRect.Top \ 16) + 1
    maxCellY = ((otherRect.Top + otherRect.Height - 1) \ 16) + 1
    
    ' Clamp to valid grid bounds
    If minCellX < 1 Then minCellX = 1
    If minCellY < 1 Then minCellY = 1
    
    ' Get grid bounds
    Dim Rows As Long, Cols As Long
    On Error Resume Next
    Rows = UBound(Grid, 1)
    Cols = UBound(Grid, 2)
    On Error GoTo 0
    
    If Rows <= 0 Or Cols <= 0 Then Exit Sub
    
    If maxCellX > Cols Then maxCellX = Cols
    If maxCellY > Rows Then maxCellY = Rows
    
    ' Check only tiles within the overlapping grid cells
    Dim Row As Long, Col As Long
    Dim t As Tile
    
    For Row = minCellY To maxCellY
        For Col = minCellX To maxCellX
            Set t = Grid(Row, Col)
            ' Only check solid tiles
            If t.TileType = "X" And Not t.DestRect Is Nothing Then
                ' Check if Other.Rectangle actually collides with the tile's destRect
                If otherRect.Collide(t.DestRect) Then
                    ' Collision detected
                    Other.Collided Me
                    Exit Sub ' Exit on first collision
                End If
            End If
        Next Col
    Next Row
End Sub

Public Sub Sprite_Collided(ByRef Other As Sprite)
    ' Level doesn't react to collisions
End Sub

Public Function Sprite_ShouldRecycle() As Boolean
    Sprite_ShouldRecycle = False
End Function

