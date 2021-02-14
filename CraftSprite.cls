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

Private SpriteContext As Long
Private SpriteMaskContext As Long
Private Bank As Integer
Private Height As Integer
Private Width As Integer
Private MaxBanks As Integer

Public Function SpriteBank_getFilename() As String
    SpriteBank_getFilename = App.Path & "\resources\pxplayer.bmp"
End Function

Public Function SpriteBank_getMaskFilename() As String
    SpriteBank_getMaskFilename = App.Path & "\resources\pxplayer_m.bmp"
End Function

Public Sub SpriteBank_setContext(ByRef context As Long, ByRef maskContext As Long)
    SpriteContext = context
    SpriteMaskContext = maskContext
    Bank = 0
    Height = 19
    Width = 30
    MaxBanks = 0
End Sub

Public Function SpriteBank_getContext() As Long
    SpriteBank_getContext = SpriteContext
End Function

Public Function SpriteBank_getMaskContext() As Long
    SpriteBank_getMaskContext = SpriteMaskContext
End Function

Public Function SpriteBank_X() As Long
    SpriteBank_X = 0
End Function

Public Function SpriteBank_Y() As Long
    SpriteBank_Y = Bank * Height
End Function

Public Function SpriteBank_Width() As Long
    SpriteBank_Width = Width
End Function

Public Function SpriteBank_Height() As Long
    SpriteBank_Height = Height
End Function


