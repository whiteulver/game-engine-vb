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

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
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
Private Boundary As Rectangle
Private Contexts() As Long
Private ContextIndex As Integer


Public Sub RenderingEngine_initialize(BoundaryIn As Rectangle)
    Set Boundary = BoundaryIn
    ContenxtIndex = 0
    DeviceContext = CreateContext
 End Sub
Private Function CreateContext() As Long
    Dim Bitmap As Long
    Dim hdc As Long
    Dim BitmapHdc As Long
    
    hdc = GetDC(0)
    BitmapHdc = GetDC(0)
    CreateContext = CreateCompatibleDC(hdc)
    Bitmap = CreateCompatibleBitmap(BitmapHdc, Boundary.Width(), Boundary.Height())
    SelectObject CreateContext, Bitmap
    ' Cleanup
    DeleteObject Bitmap
    DeleteDC hdc
    DeleteDC BitmapHdc
End Function
Public Sub RenderingEngine_createSpriteContext(Bank As SpriteBank)
    'On Error Resume Next
    Dim context As Long
    Dim maskContext As Long
    Dim cHdc As Long
    Dim mHdc As Long
    
    cHdc = GetDC(0)
    mHdc = GetDC(0)
    
    context = CreateCompatibleDC(cHdc)
    maskContext = CreateCompatibleDC(mHdc)
    SelectObject context, LoadPicture(Bank.getFilename)
    SelectObject maskContext, LoadPicture(Bank.getMaskFilename)
    DeleteDC cHdc
    DeleteDC mHdc
    Bank.setContext context, maskContext
    AddContext context
    AddContext maskContext
    
End Sub

Public Sub RenderingEngine_Draw(spr As Sprite)
    BitBlt DeviceContext, spr.Rect.Left, spr.Rect.Top, spr.Rect.Width, spr.Rect.Height, spr.Bank.getMaskContext(), spr.Bank.X, spr.Bank.Y, vbSrcAnd
    BitBlt DeviceContext, spr.Rect.Left, spr.Rect.Top, spr.Rect.Width, spr.Rect.Height, spr.Bank.getContext(), spr.Bank.X, spr.Bank.Y, vbSrcPaint
    'BitBlt DeviceContext, dstX, dstY, spriteIn.Width, spriteIn.Height, spriteIn.getMaskContext(), spriteIn.X, spriteIn.Y, vbSrcAnd
    'BitBlt DeviceContext, dstX, dstY, spriteIn.Width, spriteIn.Height, spriteIn.getContext(), spriteIn.X, spriteIn.Y, vbSrcPaint
End Sub

Public Sub RenderingEngine_Cls()
    Dim Region As Long
    Dim TmpContext As Long
    
    TmpContext = CreateContext
    
    SetBkColor TmpContext, RGB(0, 0, 0)
    SetBkMode TmpContext, 1
    
    BeginPath TmpContext
        Rectangle TmpContext, 0, 0, Boundary.Width(), Boundary.Height()
    EndPath TmpContext
    Region = PathToRegion(TmpContext)
    
    brush = CreateSolidBrush(RGB(0, 0, 0))
    SelectObject TmpContext, brush
    FillRgn TmpContext, Region, brush
    DeleteObject brush
    DeleteObject Region
    
    BitBlt DeviceContext, 0, 0, Boundary.Width(), Boundary.Height(), TmpContext, 0, 0, vbSrcCopy
    DeleteDC TmpContext
End Sub

Public Sub RenderingEngine_DrawToScreen(screen As Long, Width As Long, Height As Long)
    BitBlt screen, 0, 0, Width, Height, DeviceContext, 0, 0, vbSrcCopy
End Sub

Public Sub RenderingEngine_ClearObjects(ByRef hdc As Long)
    Dim i As Long
    For i = 0 To UBound(Contexts)
        DeleteDC Contexts(i)
    Next
    DeleteDC DeviceContext
    ReleaseDC 0, hdc
End Sub

Private Sub AddContext(context As Long)
    ReDim Preserve Contexts(ContextIndex)
    Contexts(index) = context
    ContextIndex = ContextIndex + 1
End Sub
