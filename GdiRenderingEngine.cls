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

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Const NULL_PEN As Long = 8

Private phDC As Long
Private phWnd As Long
Private DeviceContext As Long
Private BlackBuffer As Long ' Black bitmap to erase contents in screen
Private Boundary As Rectangle
Private Contexts() As Long
Private ContextIndex As Long
Private RegisteredContexts() As Long
Private RegisteredMaskContexts() As Long

Public Sub RenderingEngine_Init(ByRef hWnd As Long, ByRef BoundaryIn As Rectangle)
    ' Treat incoming handle as hwnd for engine-agnostic Init
    phWnd = hWnd
    phDC = GetDC(hWnd)
    Set Boundary = BoundaryIn
    DeviceContext = CreateContext
    CreateBlackBuffer
 End Sub

Private Function CreateContext(Optional SizeIn As Rectangle) As Long
    Dim BITMAP As Long
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
    BITMAP = CreateCompatibleBitmap(BitmaphDC, Size.Width(), Size.Height())
    SelectObject CreateContext, BITMAP
    ' Cleanup
    DeleteObject BITMAP
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
End Sub

Public Sub RenderingEngine_ClearObjects()
    Dim i As Long
    For i = 0 To UBound(Contexts)
        DeleteDC Contexts(i)
    Next
    DeleteDC DeviceContext
    DeleteDC BlackBuffer
    If phWnd <> 0 And phDC <> 0 Then
        ReleaseDC phWnd, phDC
    End If
End Sub

Private Sub AddContext(context As Long)
    ReDim Preserve Contexts(ContextIndex)
    Contexts(ContextIndex) = context
    ContextIndex = ContextIndex + 1
End Sub

Public Sub RenderingEngine_RegisterSprite(ByVal Name As Integer, ByVal Filepath As String, Optional ByVal MaskFilepath As String = "", Optional ByVal UseColorKey As Boolean = False, Optional ByVal ColorKey As Long = 0)
    ReDim Preserve RegisteredContexts(Name)
    ReDim Preserve RegisteredMaskContexts(Name)

    RegisteredContexts(Name) = RenderingEngine_CreateImageContext(Filepath)
    If LenB(MaskFilepath) > 0 Then
        If StrComp(MaskFilepath, Filepath, vbTextCompare) = 0 Then
            RegisteredMaskContexts(Name) = 0 ' treat identical file as no mask
        Else
            RegisteredMaskContexts(Name) = RenderingEngine_CreateImageContext(MaskFilepath)
        End If
    Else
        RegisteredMaskContexts(Name) = 0
    End If
End Sub

Public Sub RenderingEngine_RegisterSpriteFromContext(ByVal Name As Integer, ByVal ImgContext As Long, Optional ByVal MaskContext As Long = 0, Optional ByVal UseColorKey As Boolean = False, Optional ByVal ColorKey As Long = 0)
    ReDim Preserve RegisteredContexts(Name)
    ReDim Preserve RegisteredMaskContexts(Name)

    RegisteredContexts(Name) = ImgContext
    RegisteredMaskContexts(Name) = MaskContext
End Sub

Public Sub RenderingEngine_Draw(Name As Integer, srcRect As Rectangle, destRect As Rectangle)
    Dim mask As Long
    Dim img As Long
    
    mask = RegisteredMaskContexts(Name)
    img = RegisteredContexts(Name)
    
    If mask <> 0 Then
        BitBlt DeviceContext, destRect.Left, destRect.Top, srcRect.Width, srcRect.Height, mask, srcRect.Left, srcRect.Top, vbSrcAnd
        BitBlt DeviceContext, destRect.Left, destRect.Top, srcRect.Width, srcRect.Height, img, srcRect.Left, srcRect.Top, vbSrcPaint
    Else
        ' No mask provided: regular copy
        BitBlt DeviceContext, destRect.Left, destRect.Top, srcRect.Width, srcRect.Height, img, srcRect.Left, srcRect.Top, vbSrcCopy
    End If
End Sub
Public Sub RenderingEngine_DrawCircle(X As Long, Y As Long, Radius As Long, Color As Long)
    Dim pen As Long
    Dim brush As Long
    Dim oldPen As Long
    Dim oldBrush As Long

    ' Create outline and fill
    pen = CreatePen(0, 1, Color)
    brush = CreateSolidBrush(Color)

    ' Select them into the DC, keeping the old ones
    oldPen = SelectObject(DeviceContext, pen)
    oldBrush = SelectObject(DeviceContext, brush)

    ' Draw circle centered at (X,Y)
    Ellipse DeviceContext, X - Radius, Y - Radius, X + Radius, Y + Radius

    ' Restore original pen/brush
    SelectObject DeviceContext, oldPen
    SelectObject DeviceContext, oldBrush

    ' Delete ours to avoid leaks
    DeleteObject pen
    DeleteObject brush
End Sub

Public Sub RenderingEngine_BuildParticleAtlas(ByVal Name As Integer)
    Dim atlasImgDC As Long, atlasMaskDC As Long
    Dim atlasImgBitmap As Long, atlasMaskBitmap As Long
    Dim hDC As Long
    Dim i As Long, j As Long
    Dim Radius As Long
    Dim tileX As Long, tileY As Long
    Dim centerX As Long, centerY As Long
    Dim Color As Long
    Dim brush As Long, oldBrush As Long
    
    ' Create atlas DCs (256x192 pixels)
    hDC = GetDC(0)
    atlasImgDC = CreateCompatibleDC(hDC)
    atlasMaskDC = CreateCompatibleDC(hDC)
    
    ' Create bitmaps for the atlas
    atlasImgBitmap = CreateCompatibleBitmap(hDC, 256, 192)
    atlasMaskBitmap = CreateCompatibleBitmap(hDC, 256, 192)
    
    ' Select bitmaps into DCs
    SelectObject atlasImgDC, atlasImgBitmap
    SelectObject atlasMaskDC, atlasMaskBitmap
    
    ' Clear img to black, mask to white
    SetBkColor atlasImgDC, RGB(0, 0, 0)
    SetBkColor atlasMaskDC, RGB(255, 255, 255)
    SetBkMode atlasImgDC, 1
    SetBkMode atlasMaskDC, 1
    
    ' Fill with black background (no border)
    Dim bgOldPen As Long
    brush = CreateSolidBrush(RGB(0, 0, 0))
    oldBrush = SelectObject(atlasImgDC, brush)
    bgOldPen = SelectObject(atlasImgDC, GetStockObject(NULL_PEN))
    Rectangle atlasImgDC, 0, 0, 256, 192
    SelectObject atlasImgDC, bgOldPen
    SelectObject atlasImgDC, oldBrush
    DeleteObject brush
    
    brush = CreateSolidBrush(RGB(255, 255, 255))
    oldBrush = SelectObject(atlasMaskDC, brush)
    bgOldPen = SelectObject(atlasMaskDC, GetStockObject(NULL_PEN))
    Rectangle atlasMaskDC, 0, 0, 256, 192
    SelectObject atlasMaskDC, bgOldPen
    SelectObject atlasMaskDC, oldBrush
    DeleteObject brush
    
    ' Create circles for each radius and color combination
    Dim nullPen As Long
    Dim oldPen As Long
    nullPen = GetStockObject(NULL_PEN)
    For i = 0 To 5 ' 6 rows (radii: 2, 4, 6, 8, 12, 16)
        Radius = Choose(i + 1, 2, 4, 6, 8, 12, 16)
        tileY = i * 32
        
        For j = 0 To 7 ' 8 columns (color steps)
            tileX = j * 32
            centerX = tileX + 16 ' Center of 32x32 tile
            centerY = tileY + 16
            
            ' Calculate color based on orange->white fade
            Color = CalculateParticleColor(j)
            
            ' Draw filled circle in image atlas (pad by 2px to avoid touching tile edges)
            Dim drawR As Long
            drawR = Radius - 2
            If drawR < 1 Then drawR = 1
            brush = CreateSolidBrush(Color)
            oldBrush = SelectObject(atlasImgDC, brush)
            oldPen = SelectObject(atlasImgDC, nullPen)
            Ellipse atlasImgDC, centerX - drawR, centerY - drawR, centerX + drawR, centerY + drawR
            SelectObject atlasImgDC, oldPen
            SelectObject atlasImgDC, oldBrush
            DeleteObject brush
            
            ' Draw black circle in mask atlas
            brush = CreateSolidBrush(RGB(0, 0, 0))
            oldBrush = SelectObject(atlasMaskDC, brush)
            oldPen = SelectObject(atlasMaskDC, nullPen)
            Ellipse atlasMaskDC, centerX - drawR, centerY - drawR, centerX + drawR, centerY + drawR
            SelectObject atlasMaskDC, oldPen
            SelectObject atlasMaskDC, oldBrush
            DeleteObject brush
        Next j
    Next i
    
    ' Register the atlas contexts
    RenderingEngine_RegisterSpriteFromContext Name, atlasImgDC, atlasMaskDC
    
    ' Add contexts to cleanup list
    AddContext atlasImgDC
    AddContext atlasMaskDC
    
    ' Debug: Verify atlas was created successfully
    Debug.Print "Particle Atlas created successfully for registry " & Name
    
    ' Cleanup
    DeleteObject atlasImgBitmap
    DeleteObject atlasMaskBitmap
    DeleteDC hDC
End Sub

Private Function CalculateParticleColor(ByVal step As Long) As Long
    Dim baseR As Long, baseG As Long, baseB As Long
    Dim r As Long, g As Long, b As Long
    Dim p As Single
    
    ' Base orange color
    baseR = 255: baseG = 160: baseB = 0
    
    ' Fade parameter (0 to 1)
    p = step / 7!
    
    ' Interpolate to white
    r = baseR + CLng((255 - baseR) * p)
    g = baseG + CLng((255 - baseG) * p)
    b = baseB + CLng((255 - baseB) * p)
    
    CalculateParticleColor = RGB(r, g, b)
End Function
