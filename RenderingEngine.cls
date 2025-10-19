VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RenderingEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Init(ByRef hWnd As Long, ByRef BoundaryIn As Rectangle)
End Sub

Public Sub DrawToScreen()
End Sub

Public Sub Cls()
End Sub

Public Sub ClearObjects()
End Sub

Public Sub RegisterSprite(ByVal Name As Integer, ByVal Filepath As String, Optional ByVal MaskFilepath As String = "", Optional ByVal UseColorKey As Boolean = False, Optional ByVal ColorKey As Long = 0)
End Sub

Public Sub RegisterSpriteFromContext(ByVal Name As Integer, ByVal ImgContext As Long, Optional ByVal MaskContext As Long = 0, Optional ByVal UseColorKey As Boolean = False, Optional ByVal ColorKey As Long = 0)
End Sub

' Draw a registered sprite by coping the srcRect of sprite image to the destRect in screen
Public Sub Draw(Name As Integer, srcRect As Rectangle, destRect As Rectangle)
End Sub

Public Sub DrawCircle(X As Long, Y As Long, Radius As Long, Color As Long)
End Sub

Public Sub BuildParticleAtlas(ByVal Name As Integer)
End Sub
