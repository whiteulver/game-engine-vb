VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Engine"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7680
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


Private Sub Form_GotFocus()
    Set GameEngine = New Engine
    Set Rendering = New GdiRenderingEngine
    Set GameEventDispatcher = New EventDispatcher
    
    initializeListeners GameEventDispatcher
    GameEngine.initialize Rendering, GameEventDispatcher
    
    DoEvents
    
    GameEngine.run frmMain
End Sub

Private Sub Form_Load()
    'Set GameEventDispatcher = New EventDispatcher
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
    Dim bndAction As New BackgroundAction
    Dim bndSprite As New BackgroundSprite
    
    Rendering.createSpriteContext bndSprite
    bndAction.initialize bndSprite
    bndListener.initialize bndAction
    
    Dim mvAction As New MoveAction
    Dim mvListener As New MoveListener
    Dim cftSprite As New CraftSprite
    
    Rendering.createSpriteContext cftSprite
    mvAction.initialize cftSprite
    mvListener.initialize mvAction

    Dispatcher.addListener "idle", bndListener
    Dispatcher.addListener "idle", mvListener
    Dispatcher.addListener "keydown", mvListener
End Sub
