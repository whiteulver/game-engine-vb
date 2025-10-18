VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsteroidBatchSpawner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Sprite

Private Asteroids() As Sprite
Private MaxAsteroids As Integer
Private LastSpawnTime As Long
Private SpawnInterval As Long ' milliseconds
Private BatchSize As Integer
Private ScreenWidth As Long
Private ScreenHeight As Long

Private Sub Class_Initialize()
    MaxAsteroids = 100 ' Maximum asteroids in memory
    SpawnInterval = 1000
    BatchSize = 4 ' Number of asteroids per batch
    ScreenWidth = 320 ' Adjust based on your screen size
    ScreenHeight = 256 ' Adjust based on your screen size
    LastSpawnTime = 0
    ReDim Asteroids(MaxAsteroids - 1)
End Sub

Public Sub Sprite_Init()
    ' Reset spawn timer
    LastSpawnTime = 0
End Sub

Public Sub Sprite_Update(ByVal Time As Long, Delta As Double)
    ' Check if it's time to spawn a new batch
    If Time - LastSpawnTime >= SpawnInterval Then
        SpawnBatch Time
        LastSpawnTime = Time
    End If
    
    ' Update all active asteroids
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            Asteroids(i).Update Time, Delta
            ' Check if Asteroid is off-screen or should recycle
            If (Asteroids(i).Rect.Left + Asteroids(i).Rect.Width < 0) _
                Or Asteroids(i).ShouldRecycle() Then
                Set Asteroids(i) = Nothing
            End If
        End If
    Next i
End Sub

Public Sub Sprite_Draw(ByRef Render As RenderingEngine)
    ' Draw all active asteroids
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            Asteroids(i).Draw Render
        End If
    Next i
End Sub

Public Property Get Sprite_Rect() As Rectangle
    ' Return a dummy rectangle since this is a manager
    Set Sprite_Rect = New Rectangle
    Sprite_Rect.Init 0, 0, 0, 0
End Property

Public Sub Sprite_Listen(ByRef GameEventIn As DomainEvent)
    ' Pass events to all active asteroids
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            Asteroids(i).Listen GameEventIn
        End If
    Next i
End Sub

Public Sub Sprite_Destroy()
    ' Destroy all asteroids
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            If Asteroids(i).IsAlive Then Asteroids(i).Destroy
        End If
    Next i
End Sub

Public Function Sprite_IsAlive() As Boolean
    Sprite_IsAlive = True
End Function

Public Property Get Sprite_Registry() As Integer
    Sprite_Registry = 2
End Property

Public Sub Sprite_Collide(ByRef Other As Sprite)
    ' Check collisions with all active asteroids
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            If Asteroids(i).IsAlive Then Asteroids(i).Collide Other
        End If
    Next i
End Sub

Public Sub Sprite_Collided(ByRef Other As Sprite)
    ' Handle collision with all active asteroids
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            If Asteroids(i).IsAlive Then Asteroids(i).Collided Other
        End If
    Next i
End Sub

Public Function Sprite_ShouldRecycle() As Boolean
    Sprite_ShouldRecycle = False ' Manager never recycles
End Function

Private Sub SpawnBatch(ByVal Time As Long)
    Dim spawned As Integer
    Dim i As Integer
    
    spawned = 0
    For i = 0 To MaxAsteroids - 1
        If spawned >= BatchSize Then Exit For
        If Asteroids(i) Is Nothing Then
            Dim randomY As Integer
            randomY = Int(Rnd() * (ScreenHeight - 50)) + 10

            Dim hAsteroid As Sprite
            Set hAsteroid = New Asteroid
            hAsteroid.Rect.Left = ScreenWidth + 10
            hAsteroid.Rect.Top = randomY
            Set Asteroids(i) = hAsteroid

            spawned = spawned + 1
        End If
    Next i
End Sub

' Method to get all active asteroids for collision checking
Public Function GetActiveAsteroids() As Collection
    Set GetActiveAsteroids = New Collection
    Dim i As Integer
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            If Asteroids(i).IsAlive Then
                GetActiveAsteroids.Add Asteroids(i)
            End If
        End If
    Next i
End Function

Private Function OverlapsExisting(ByRef candidate As Sprite, ByRef placed As Collection) As Boolean
    Dim i As Integer
    ' Check against already active asteroids
    For i = 0 To MaxAsteroids - 1
        If Not Asteroids(i) Is Nothing Then
            If Asteroids(i).IsAlive Then
                If candidate.Rect.Collide(Asteroids(i).Rect) Then
                    OverlapsExisting = True
                    Exit Function
                End If
            End If
        End If
    Next i
    ' Check against asteroids placed in this batch
    If Not placed Is Nothing Then
        Dim a As Sprite
        For Each a In placed
            If candidate.Rect.Collide(a.Rect) Then
                OverlapsExisting = True
                Exit Function
            End If
        Next a
    End If
    OverlapsExisting = False
End Function
