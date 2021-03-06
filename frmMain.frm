VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Engine"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
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
    Set State = New PlayingState
    
    Rendering.Init frmMain, Boundary
    GameEngine.Init Rendering, Boundary, State
    
    GameEngine.run
    End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Form_Unload 0
    End If
End Sub

Private Sub Form_Load()
    Set Boundary = New Rectangle
    Boundary.Init 0, 0, 320, 256
    Me.Height = Boundary.Height() * 2 * (Height / ScaleHeight)
    Me.Width = Boundary.Width * 2 * (Width / ScaleWidth)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GameEngine.StopEngine
End Sub
