VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackgroundSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements SpriteBank

Private Bank As Integer
Private Height As Long
Private Width As Long
Private MaxBanks As Integer
Private Rect As Rectangle

Private Sub Class_Initialize()
    Bank = 0
    Height = 256
    Width = 512
    MaxBanks = 8
End Sub

Public Function SpriteBank_getFilename() As String
    SpriteBank_getFilename = App.Path & "\resources\fantasy-512-x-256_full.bmp"
End Function

Public Function SpriteBank_getMaskFilename() As String
    SpriteBank_getMaskFilename = App.Path & "\resources\fantasy-512-x-256_full.bmp"
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

Public Property Let SpriteBank_Bank(ByVal index As Integer)
    Bank = index
End Property

Public Property Get SpriteBank_Rect() As Rectangle
    Set SpriteBank_Rect = New Rectangle
    SpriteBank_Rect.Init SpriteBank_X(), SpriteBank_Y(), SpriteBank_Width(), SpriteBank_Height()
End Property
