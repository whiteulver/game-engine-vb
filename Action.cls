VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Tick As Long
Private Frames() As Integer
Private Delay As Integer
Private FrameIndex As Integer

Public Sub Init(FramesIn() As Integer, DelayIn As Integer)
    Frames = FramesIn
    Delay = DelayIn
End Sub

Public Function NextFrame() As Integer
    If UBound(Frames) = 0 Then
        FrameIndex = 0
    Else
        If Tick > Delay Then
            Tick = 0
            FrameIndex = FrameIndex + 1
            If FrameIndex > UBound(Frames) Then
                FrameIndex = UBound(Frames)
            End If
        End If
        Tick = Tick + 1
    End If
    NextFrame = Frames(FrameIndex)
End Function
Public Sub Reset()
    FrameIndex = 0
End Sub

Private Sub Class_Initialize()
    Tick = 0
    FrameIndex = 0
End Sub
