VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CraftSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements SpriteBank

Private Bank As Integer
Private Height As Integer
Private Width As Integer
Private MaxBanks As Integer
Private Frames(14) As Frame

Private Type Frame
    OffsetX As Long
    OffsetY As Long
End Type

Private Sub Class_Initialize()
    Dim OffsetX As Integer, OffsetY As Integer
    Bank = 0
    OffsetX = 0
    OffsetY = 0
    Height = 24
    Width = 32
    MaxBanks = 14
    
    For i = 0 To MaxBanks
        Dim tmpFrame As Frame
        If i = 5 Or i = 10 Then
            OffsetX = 0
            OffsetY = OffsetY + 1
        End If
        tmpFrame.OffsetX = OffsetX * Width
        tmpFrame.OffsetY = OffsetY * Height
        Frames(i) = tmpFrame
        OffsetX = OffsetX + 1
    Next
End Sub

Public Function SpriteBank_getFilename() As String
    SpriteBank_getFilename = App.Path & "\resources\pxplayer.bmp"
End Function

Public Function SpriteBank_getMaskFilename() As String
    SpriteBank_getMaskFilename = App.Path & "\resources\pxplayer_m.bmp"
End Function

Public Function SpriteBank_getMaskContext() As Long
    SpriteBank_getMaskContext = SpriteMaskContext
End Function

Public Function SpriteBank_X() As Long
    SpriteBank_X = Frames(Bank).OffsetX
End Function

Public Function SpriteBank_Y() As Long
    SpriteBank_Y = Frames(Bank).OffsetY
End Function

Public Function SpriteBank_Width() As Long
    SpriteBank_Width = Width
End Function

Public Function SpriteBank_Height() As Long
    SpriteBank_Height = Height
End Function

Public Property Let SpriteBank_Bank(ByVal index As Integer)
    Bank = index
End Property

Public Property Get SpriteBank_Rect() As Rectangle
    Set SpriteBank_Rect = New Rectangle
    SpriteBank_Rect.Init SpriteBank_X(), SpriteBank_Y(), SpriteBank_Width(), SpriteBank_Height()
End Property
