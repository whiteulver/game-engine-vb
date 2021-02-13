VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
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
Private GameEventDispatcher As EventDispatcher
Private Boundary As Rectangle

Private Sub Form_GotFocus()
    Set GameEngine = New Engine
    Set Rendering = New GdiRenderingEngine
    Set GameEventDispatcher = New EventDispatcher
    
    initializeListeners GameEventDispatcher
    GameEngine.initialize Rendering, GameEventDispatcher, Boundary
    
    DoEvents
    
    GameEngine.run frmMain
End Sub

Private Sub Form_Load()
    Set Boundary = New Rectangle
    Boundary.initialize 0, 0, 512, 240
    Me.Height = Boundary.getHeight() * (Height / ScaleHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim TerminateEvent As DomainEvent
    Set TerminateEvent = New DomainEvent
    TerminateEvent.initialize "Terminate", GameEngine
    GameEngine.raise TerminateEvent
    GameEngine.terminate
    End
End Sub

Private Sub initializeListeners(Dispatcher As EventDispatcher)
    Dim bndListener As New BackgroundListener
    Dim bndSprite As New BackgroundSprite
    
    Rendering.createSpriteContext bndSprite
    bndListener.initialize bndSprite
    
    Dim mvListener As New MoveListener
    Dim cftSprite As New CraftSprite
    
    Rendering.createSpriteContext cftSprite
    mvListener.initialize cftSprite

    Dispatcher.addListener "idle", bndListener
    Dispatcher.addListener "idle", mvListener
    Dispatcher.addListener "keydown", mvListener
End Sub
