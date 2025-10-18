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
                
                ' fade color to white over time
                .Life = .Life - 1
                Dim p As Single
                p = 1! - (.Life / 40!)
                If p < 0! Then p = 0!
                If p > 1! Then p = 1!
                
                Dim r0 As Long, g0 As Long, b0 As Long
                Dim r As Long, g As Long, b As Long
                r0 = (.BaseColor And &HFF&)
                g0 = ((.BaseColor \ &H100&) And &HFF&)
                b0 = ((.BaseColor \ &H10000) And &HFF&)
                r = r0 + CLng((255 - r0) * p)
                g = g0 + CLng((255 - g0) * p)
                b = b0 + CLng((255 - b0) * p)
                .Color = RGB(r, g, b)
                
                active = True
            End If
        End With
    Next
    
    pAlive = active
End Sub

Public Sub Draw(ByRef Render As RenderingEngine)
    Dim i As Long
    For i = 0 To pParticleCount - 1
        With Particles(i)
            If .Life > 0 Then
                Render.DrawCircle _
                    CLng(.X), CLng(.Y), CLng(.Radius), .Color
            End If
        End With
    Next
End Sub

Public Sub Destroy()
    Erase Particles
End Sub
