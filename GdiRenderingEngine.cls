VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GdiRenderingEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements RenderingEngine

Private DeviceContext As Long
Private EmptyDeviceContext As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Sub RenderingEngine_initialize()
    DeviceContext = CreateContext
    EmptyDeviceContext = CreateContext
    'ScreenDeviceContext = GetDC(0)
    'DeviceContext = CreateCompatibleDC(ScreenDeviceContext)
    'Bitmap = CreateCompatibleBitmap(ScreenDeviceContext, 640, 240)
    'SelectObject DeviceContext, Bitmap
    'DeleteObject Bitmap
End Sub
Private Function CreateContext() As Long
    Dim TmpDeviceContext As Long
    Dim Bitmap As Long
    TmpDeviceContext = CreateCompatibleDC(GetDC(0))
    Bitmap = CreateCompatibleBitmap(GetDC(0), 512, 256)
    SelectObject TmpDeviceContext, Bitmap
    CreateContext = TmpDeviceContext
    DeleteObject Bitmap
End Function
Public Sub RenderingEngine_createSpriteContext(sprite As SpriteBank)
    'On Error Resume Next
    Dim context As Long
    Dim maskContext As Long
    
    context = CreateCompatibleDC(GetDC(0))
    maskContext = CreateCompatibleDC(GetDC(0))
    SelectObject context, LoadPicture(sprite.getFilename)
    SelectObject maskContext, LoadPicture(sprite.getMaskFilename)
    
    sprite.setContext context, maskContext
End Sub

Public Sub RenderingEngine_draw(sprite As SpriteBank, dstX As Long, dstY As Long)
    BitBlt DeviceContext, dstX, dstY, sprite.getWidth, sprite.getHeight, sprite.getMaskContext(), sprite.getX(), sprite.getY(), vbSrcAnd
    BitBlt DeviceContext, dstX, dstY, sprite.getWidth, sprite.getHeight, sprite.getContext(), sprite.getX(), sprite.getY(), vbSrcPaint
    
    'ReleaseDC 0, ScreenDeviceContext
End Sub
Public Sub RenderingEngine_Cls(Width As Long, Height As Long)
    Dim Region As Long
    Dim TmpContext As Long
    
    TmpContext = CreateContext
    
    SetBkColor TmpContext, RGB(0, 0, 0)
    SetBkMode TmpContext, 1
    
    BeginPath TmpContext
        Rectangle TmpContext, 0, 0, 512, 256
    EndPath TmpContext
    Region = PathToRegion(TmpContext)
    
    brush = CreateSolidBrush(RGB(0, 0, 0))
    SelectObject TmpContext, brush
    FillRgn TmpContext, Region, brush
    
    BitBlt DeviceContext, 0, 0, Width, Height, TmpContext, 0, 0, vbSrcCopy
    DeleteDC TmpContext
    'DeleteDC DeviceContext
End Sub
Public Sub RenderingEngine_DrawToScreen(screen As Long, Width As Long, Height As Long)
    BitBlt screen, 0, 0, Width, Height, DeviceContext, 0, 0, vbSrcCopy
End Sub

Public Sub RenderingEngine_ClearObjects()
    DeleteDC DeviceContext
End Sub
