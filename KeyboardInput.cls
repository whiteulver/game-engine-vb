VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyboardInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type KeyboardInput
    btnDown As Boolean
    btnRight As Boolean
    btnUp As Boolean
    btnLeft As Boolean
    btnA As Boolean
    btnB As Boolean
    btnC As Boolean
    btnD As Boolean
    btnE As Boolean
    btnF As Boolean
    btnG As Boolean
    btnH As Boolean
    btnI As Boolean
    btnJ As Boolean
    btnK As Boolean
    btnL As Boolean
    btnM As Boolean
    btnN As Boolean
    btnO As Boolean
    btnP As Boolean
    btnQ As Boolean
    btnR As Boolean
    btnS As Boolean
    btnT As Boolean
    btnU As Boolean
    btnV As Boolean
    btnW As Boolean
    btnX As Boolean
    btnY As Boolean
    btnZ As Boolean
    btnAlt As Boolean
    btnControl As Boolean
    btnReturn As Boolean
End Type

Private Keyboard As KeyboardInput

Private Sub GetKeyboardInput()

    With Keyboard
        .btnDown = GetAsyncKeyState(vbKeyDown)
        .btnUp = GetAsyncKeyState(vbKeyUp)
        .btnRight = GetAsyncKeyState(vbKeyRight)
        .btnLeft = GetAsyncKeyState(vbKeyLeft)
        .btnA = GetAsyncKeyState(vbKeyA)
        .btnB = GetAsyncKeyState(vbKeyB)
        .btnC = GetAsyncKeyState(vbKeyC)
        .btnD = GetAsyncKeyState(vbKeyD)
        .btnE = GetAsyncKeyState(vbKeyE)
        .btnF = GetAsyncKeyState(vbKeyF)
        .btnG = GetAsyncKeyState(vbKeyG)
        .btnH = GetAsyncKeyState(vbKeyH)
        .btnI = GetAsyncKeyState(vbKeyI)
        .btnJ = GetAsyncKeyState(vbKeyJ)
        .btnK = GetAsyncKeyState(vbKeyK)
        .btnL = GetAsyncKeyState(vbKeyL)
        .btnM = GetAsyncKeyState(vbKeyM)
        .btnN = GetAsyncKeyState(vbKeyN)
        .btnO = GetAsyncKeyState(vbKeyO)
        .btnP = GetAsyncKeyState(vbKeyP)
        .btnQ = GetAsyncKeyState(vbKeyQ)
        .btnR = GetAsyncKeyState(vbKeyR)
        .btnS = GetAsyncKeyState(vbKeyS)
        .btnT = GetAsyncKeyState(vbKeyT)
        .btnU = GetAsyncKeyState(vbKeyU)
        .btnV = GetAsyncKeyState(vbKeyV)
        .btnW = GetAsyncKeyState(vbKeyW)
        .btnX = GetAsyncKeyState(vbKeyX)
        .btnY = GetAsyncKeyState(vbKeyY)
        .btnZ = GetAsyncKeyState(vbKeyZ)
        .btnControl = GetAsyncKeyState(vbKeyControl)
        .btnReturn = GetAsyncKeyState(vbKeyReturn)
    End With
End Sub

Public Function getKeyCodes() As Integer()
    Dim KeyEvent As DomainEvent
    Dim KeyCode(3) As Integer
    KeyCode(0) = 0: KeyCode(1) = 0: KeyCode(2) = 0: KeyCode(3) = 0
    'KeyCode = 0
    
    GetKeyboardInput
    
    If Keyboard.btnLeft Then
        KeyCode(0) = 37
    End If
    
    If Keyboard.btnRight Then
        KeyCode(1) = 39
    End If
    
    If Keyboard.btnUp Then
        KeyCode(2) = 1
    End If
    
    If Keyboard.btnDown Then
        KeyCode(3) = 2
    End If
    
    getKeyCodes = KeyCode
End Function
