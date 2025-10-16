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
Option Explicit
Implements RenderingEngine

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private phDC As Long
Private DeviceContext As Long
Private BlackBuffer As Long ' Black bitmap to erase contents in screen
Private Boundary As Rectangle
Private Contexts() As Long
Private ContextIndex As Long
Private RegisteredContexts() As Long
Private RegisteredMaskContexts() As Long

Public Sub RenderingEngine_Init(hDC As Long, BoundaryIn As Rectangle)
    phDC = hDC
    Set Boundary = BoundaryIn
    DeviceContext = CreateContext
    CreateBlackBuffer
 End Sub

Private Function CreateContext(Optional SizeIn As Rectangle) As Long
    Dim Bitmap As Long
    Dim hDC As Long
    Dim BitmaphDC As Long
    Dim Size As Rectangle
    
    If SizeIn Is Nothing Then
        Set Size = Boundary
    Else
        Set Size = SizeIn
    End If
    
    hDC = GetDC(0)
    BitmaphDC = GetDC(0)
    CreateContext = CreateCompatibleDC(hDC)
    Bitmap = CreateCompatibleBitmap(BitmaphDC, Size.Width(), Size.Height())
    SelectObject CreateContext, Bitmap
    ' Cleanup
    DeleteObject Bitmap
    DeleteDC hDC
    DeleteDC BitmaphDC
End Function

' Gets the filepath of an image and creates a hDC context
Private Function RenderingEngine_CreateImageContext(Filepath As String) As Long
    Dim context As Long, hDC As Long
    
    hDC = GetDC(0)
    context = CreateCompatibleDC(hDC)
    SelectObject context, LoadPicture(Filepath)
    DeleteDC hDC
    AddContext context
    RenderingEngine_CreateImageContext = context
End Function

Public Sub RenderingEngine_Cls()
    BitBlt DeviceContext, 0, 0, Boundary.Width(), Boundary.Height(), BlackBuffer, 0, 0, vbSrcCopy
End Sub
Private Sub CreateBlackBuffer()
    Dim Region As Long
    Dim tmpContext As Long
    Dim brush As Long
    
    tmpContext = CreateContext
    
    SetBkColor tmpContext, RGB(0, 0, 0)
    SetBkMode tmpContext, 1
    
    BeginPath tmpContext
        Rectangle tmpContext, 0, 0, Boundary.Width(), Boundary.Height()
    EndPath tmpContext
    Region = PathToRegion(tmpContext)
    
    brush = CreateSolidBrush(RGB(0, 0, 0))
    SelectObject tmpContext, brush
    FillRgn tmpContext, Region, brush
    DeleteObject brush
    DeleteObject Region

    BlackBuffer = tmpContext
    DeleteDC tmpContext
End Sub
Public Sub RenderingEngine_DrawToScreen()
    StretchBlt phDC, 0, 0, Boundary.Width * 2, Boundary.Height * 2, DeviceContext, 0, 0, Boundary.Width, Boundary.Height, vbSrcCopy
    'pForm.Refresh
End Sub

Public Sub RenderingEngine_ClearObjects()
    Dim i As Long
    For i = 0 To UBound(Contexts)
        DeleteDC Contexts(i)
    Next
    DeleteDC DeviceContext
    DeleteDC BlackBuffer
    ReleaseDC 0, phDC
End Sub

Private Sub AddContext(context As Long)
    ReDim Preserve Contexts(ContextIndex)
    Contexts(ContextIndex) = context
    ContextIndex = ContextIndex + 1
End Sub

Public Sub RenderingEngine_RegisterSprite(ByVal Name As Integer, ByVal Filepath As String, ByVal MaskFilepath As String)
    ReDim Preserve RegisteredContexts(Name)
    ReDim Preserve RegisteredMaskContexts(Name)

    RegisteredContexts(Name) = RenderingEngine_CreateImageContext(Filepath)
    RegisteredMaskContexts(Name) = RenderingEngine_CreateImageContext(MaskFilepath)
End Sub

Public Sub RenderingEngine_Draw(Name As Integer, srcRect As Rectangle, destRect As Rectangle)
    Dim mask As Long
    Dim img As Long
    
    mask = RegisteredMaskContexts(Name)
    img = RegisteredContexts(Name)
    
    BitBlt DeviceContext, destRect.Left, destRect.Top, srcRect.Width, srcRect.Height, mask, srcRect.Left, srcRect.Top, vbSrcAnd
    BitBlt DeviceContext, destRect.Left, destRect.Top, srcRect.Width, srcRect.Height, img, srcRect.Left, srcRect.Top, vbSrcPaint
End Sub
