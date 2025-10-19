VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParticleLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PARTICLE_ATLAS_REGISTRY As Integer = 5
Private Const TILE_SIZE As Long = 32

Private Type Particle
    X As Single
    Y As Single
    Radius As Single
    DirX As Single
    DirY As Single
    Speed As Single
    Life As Single
    Color As Long
    BaseColor As Long
    Row As Long
    Col As Long
End Type
Private Particles() As Particle
Private pAlive As Boolean
Private pParticleCount As Long
Private pRadiusGrow As Single
Public Sub Class_Initialize()
End Sub
Public Property Get Alive() As Boolean
    Alive = pAlive
End Property

Public Sub Init(ByVal Radius As Single, ByVal RadiusGrow As Single, _
    ByVal Life As Single, ByVal Bounds As Rectangle, _
    ByVal ParticleCount As Integer)
    
    Dim i As Long
    Dim cx As Single, cy As Single
    pParticleCount = ParticleCount
    pRadiusGrow = RadiusGrow
    
    cx = Bounds.Left + Bounds.Width / 2
    cy = Bounds.Top + Bounds.Height / 2
    
    ReDim Particles(ParticleCount - 1)
    
    For i = 0 To ParticleCount - 1
        With Particles(i)
            ' position: center of bounds
            .X = cx
            .Y = cy
            
            ' color: orange with slight variation
            Dim baseR As Integer, baseG As Integer, baseB As Integer
            baseR = 255
            baseG = 140 + Int(Rnd * 40)
            baseB = Int(Rnd * 30)
            .BaseColor = RGB(baseR, baseG, baseB)
            .Color = .BaseColor
            
            ' direction: random normalized vector
            Dim dx As Single, dy As Single, mag As Single
            dx = (Rnd * 2!) - 1!
            dy = (Rnd * 2!) - 1!
            mag = Sqr(dx * dx + dy * dy)
            If mag > 0 Then
                .DirX = dx / mag
                .DirY = dy / mag
            Else
                .DirX = 1!
                .DirY = 0!
            End If
            
            .Speed = 2!
            
            ' radius: start small
            .Radius = Radius
            
            ' life: start high, will be managed by screen bounds
            .Life = Life + Int(Rnd * 10)
        End With
    Next
End Sub
Public Sub Update(ByVal Time As Long, Delta As Double)
    
    Dim i As Long, active As Boolean
   
    active = False
    
    For i = 0 To pParticleCount - 1
        With Particles(i)
            If .Life > 0 Then
                ' simple linear movement
                .X = .X + .DirX * .Speed
                .Y = .Y + .DirY * .Speed

                ' grow radius steadily
                .Radius = .Radius + pRadiusGrow '0.4
                
                ' advance life and compute atlas indices
                .Life = .Life - 1
                Dim p As Single
                p = (.Radius - 0.02) / 8!
                If p < 0! Then p = 0!
                If p > 1! Then p = 1!
                .Col = CLng(p * 7!)
                .Row = MapRadiusToRow(.Radius)
                
                active = True
            End If
        End With
    Next
    
    pAlive = active
End Sub

Public Sub Draw(ByRef Render As RenderingEngine)
    Dim i As Long
    Dim srcRect As Rectangle
    Dim destRect As Rectangle
    Dim sx As Long, sy As Long, sw As Long, sh As Long

    For i = 0 To pParticleCount - 1
        With Particles(i)
            If .Life > 0 Then
                ' Use full 32x32 tile
                sx = .Col * TILE_SIZE
                sy = .Row * TILE_SIZE
                sw = TILE_SIZE
                sh = TILE_SIZE
                Set srcRect = New Rectangle
                srcRect.Init sx, sy, sw, sh

                Set destRect = New Rectangle
                destRect.Init CLng(.X) - (sw \ 2), CLng(.Y) - (sh \ 2), sw, sh

                Render.Draw PARTICLE_ATLAS_REGISTRY, srcRect, destRect
            End If
        End With
    Next
End Sub

Private Function MapRadiusToRow(ByVal r As Single) As Long
    If r < 3! Then
        MapRadiusToRow = 0
        Exit Function
    End If
    If r < 5! Then
        MapRadiusToRow = 1
        Exit Function
    End If
    If r < 7! Then
        MapRadiusToRow = 2
        Exit Function
    End If
    If r < 10! Then
        MapRadiusToRow = 3
        Exit Function
    End If
    If r < 14! Then
        MapRadiusToRow = 4
        Exit Function
    End If
    MapRadiusToRow = 5
End Function

Public Sub Destroy()
    Erase Particles
End Sub
