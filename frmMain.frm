VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Engine"
   ClientHeight    =   7680
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GameEngine As Engine
Attribute GameEngine.VB_VarHelpID = -1
Private Rendering As RenderingEngine
Private Boundary As Rectangle
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Sub Form_GotFocus()
    Set GameEngine = New Engine
    Set Rendering = New GdiRenderingEngine
    Dim State As GameState
    Set State = New PlayingState
    
    Rendering.Init frmMain.hDC, Boundary
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

    ShowCursor False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCursor True
    GameEngine.StopEngine
End Sub
