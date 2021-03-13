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
Option Explicit

' The Action class iterates through the frame indexes, according to delay
' and provides the next frame index to the sprite so it can display it.
Private Tick As Long ' Keep the ticks that have passed since the begining of the action. Reset when tick reaches the Delay.
Private Frames() As Integer ' An array of the frame indexes
Private Delay As Integer ' The Delay in ticks between the frames.
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
    Tick = 0
    FrameIndex = 0
End Sub

Private Sub Class_Initialize()
    Tick = 0
    FrameIndex = 0
End Sub
