VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Engine"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GameEngine As Engine
Private Rendering As RenderingEngine
Private Boundary As Rectangle

Private Sub Form_GotFocus()
    Set GameEngine = New Engine
    Set Rendering = New GdiRenderingEngine
    Dim State As GameState
    Set State = New DemoState
    
    GameEngine.initialize Rendering, Boundary, State
    
    DoEvents
    
    GameEngine.run frmMain
    End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Form_Unload 0
    End If
End Sub

Private Sub Form_Load()
    Set Boundary = New Rectangle
    Boundary.initialize 0, 0, 512, 256
    Me.Height = Boundary.Height() * (Height / ScaleHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GameEngine.terminate
End Sub
