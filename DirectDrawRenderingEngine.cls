VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectDrawRenderingEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements RenderingEngine

' DirectX objects
Private DX As DirectX7
Private DDRAW As DirectDraw7
Private Primary As DirectDrawSurface7
Private BackSurface As DirectDrawSurface7
Private Clipper As DirectDrawClipper

' Surface management
Private RegisteredSurfaces() As DirectDrawSurface7
Private RegisteredMaskSurfaces() As DirectDrawSurface7
Private UseColorKeyFlag() As Boolean
Private SurfaceIndex As Long

' Rendering state
Private Boundary As Rectangle
Private phWnd As Long
Private destRect As RECT

' Helper declarations
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Const IMAGE_BITMAP As Long = 0
Private Const LR_LOADFROMFILE As Long = &H10
Private Const LR_CREATEDIBSECTION As Long = &H2000

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const NULL_PEN As Long = 8

' Helper to zero DDSURFACEDESC2
Private Sub ZeroDDSD(ByRef d As DDSURFACEDESC2)
    Dim tmp As DDSURFACEDESC2
    tmp.lSize = Len(tmp)
    tmp.lFlags = 0
    tmp.lHeight = 0
    tmp.lWidth = 0
    tmp.lPitch = 0
    tmp.lBackBufferCount = 0
    tmp.ddsCaps.lCaps = 0
    d = tmp
End Sub

' Helper to convert Rectangle to RECT
Private Sub RectangleToRECT(ByRef srcRect As Rectangle, ByRef r As RECT)
    r.Left = srcRect.Left
    r.Top = srcRect.Top
    r.Right = srcRect.Left + srcRect.Width
    r.Bottom = srcRect.Top + srcRect.Height
End Sub

' Clip a blit so it stays within the backbuffer bounds. Returns False if fully off-screen
Private Function ClipToBackbuffer(ByRef srcIn As RECT, ByVal destXIn As Long, ByVal destYIn As Long, ByVal backW As Long, ByVal backH As Long, ByRef srcOut As RECT, ByRef destXOut As Long, ByRef destYOut As Long) As Boolean
    Dim width As Long, height As Long
    Dim destX As Long, destY As Long
    Dim src As RECT
    Dim dx As Long, dy As Long
    Dim overW As Long, overH As Long

    src = srcIn
    destX = destXIn
    destY = destYIn

    width = src.Right - src.Left
    height = src.Bottom - src.Top

    ' Clip left
    If destX < 0 Then
        dx = -destX
        If dx >= width Then ClipToBackbuffer = False: Exit Function
        src.Left = src.Left + dx
        destX = 0
        width = width - dx
    End If

    ' Clip top
    If destY < 0 Then
        dy = -destY
        If dy >= height Then ClipToBackbuffer = False: Exit Function
        src.Top = src.Top + dy
        destY = 0
        height = height - dy
    End If

    ' Clip right
    If destX + width > backW Then
        overW = destX + width - backW
        If overW >= width Then ClipToBackbuffer = False: Exit Function
        width = width - overW
    End If
    src.Right = src.Left + width

    ' Clip bottom
    If destY + height > backH Then
        overH = destY + height - backH
        If overH >= height Then ClipToBackbuffer = False: Exit Function
        height = height - overH
    End If
    src.Bottom = src.Top + height

    ' Output
    srcOut = src
    destXOut = destX
    destYOut = destY
    ClipToBackbuffer = (width > 0 And height > 0)
End Function

Public Sub RenderingEngine_Init(ByRef hWnd As Long, ByRef BoundaryIn As Rectangle)
    On Error GoTo InitFailed
    
    ' Store references
    phWnd = hWnd
    Set Boundary = BoundaryIn
    
    ' Create DirectX objects
    Set DX = New DirectX7
    Set DDRAW = DX.DirectDrawCreate("")
    
    ' Set windowed mode
    DDRAW.SetCooperativeLevel phWnd, DDSCL_NORMAL
    
    ' Create primary surface
    Dim SurfDesc As DDSURFACEDESC2
    ZeroDDSD SurfDesc
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set Primary = DDRAW.CreateSurface(SurfDesc)
    
    ' Create backbuffer (system memory for GDI compatibility)
    ZeroDDSD SurfDesc
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    SurfDesc.lWidth = Boundary.Width
    SurfDesc.lHeight = Boundary.Height
    Set BackSurface = DDRAW.CreateSurface(SurfDesc)
    
    ' Create clipper and attach to form
    Set Clipper = DDRAW.CreateClipper(0)
    Clipper.SetHWnd phWnd
    Primary.SetClipper Clipper
    
    ' Calculate and store destination rect
    DX.GetWindowRect phWnd, destRect
    
    ' Initialize surface arrays
    ReDim RegisteredSurfaces(0)
    ReDim RegisteredMaskSurfaces(0)
    ReDim UseColorKeyFlag(0)
    SurfaceIndex = 0
    
    Exit Sub
    
InitFailed:
    MsgBox "DirectDraw initialization failed: " & Err.Description, vbCritical
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub RenderingEngine_DrawToScreen()
    Dim srcRect As RECT
    Dim fullRect As RECT
    
    ' Update destination rect (in case window moved)
    Dim client As RECT
    Dim pt As POINTAPI
    GetClientRect phWnd, client
    pt.X = 0: pt.Y = 0
    ClientToScreen phWnd, pt
    destRect.Left = pt.X
    destRect.Top = pt.Y
    destRect.Right = pt.X + (client.Right - client.Left)
    destRect.Bottom = pt.Y + (client.Bottom - client.Top)
    
    ' Source rect: entire backbuffer
    fullRect.Left = 0
    fullRect.Top = 0
    fullRect.Right = Boundary.Width
    fullRect.Bottom = Boundary.Height
    
    ' Blit backbuffer to primary surface (clipper handles form bounds)
    Primary.Blt destRect, BackSurface, fullRect, DDBLT_WAIT
End Sub

Public Sub RenderingEngine_Cls()
    Dim fullRect As RECT
    fullRect.Left = 0
    fullRect.Top = 0
    fullRect.Right = Boundary.Width
    fullRect.Bottom = Boundary.Height
    BackSurface.BltColorFill fullRect, RGB(0, 0, 0)
End Sub

Public Sub RenderingEngine_ClearObjects()
    On Error Resume Next
    
    ' Clean up surfaces
    Dim i As Long
    For i = 0 To UBound(RegisteredSurfaces)
        Set RegisteredSurfaces(i) = Nothing
        Set RegisteredMaskSurfaces(i) = Nothing
    Next
    
    ' Clean up main surfaces
    Set BackSurface = Nothing
    Set Primary = Nothing
    
    ' Clean up clipper
    If Not Clipper Is Nothing Then
        Clipper.SetHWnd 0
        Set Clipper = Nothing
    End If
    
    ' Clean up DirectDraw objects
    Set DDRAW = Nothing
    Set DX = Nothing
End Sub

Public Sub RenderingEngine_RegisterSprite(ByVal Name As Integer, ByVal Filepath As String, Optional ByVal MaskFilepath As String = "", Optional ByVal UseColorKey As Boolean = False, Optional ByVal ColorKey As Long = 0)
    ' Resize arrays to accommodate the new sprite
    If Name > UBound(RegisteredSurfaces) Then
        ReDim Preserve RegisteredSurfaces(Name)
        ReDim Preserve RegisteredMaskSurfaces(Name)
        ReDim Preserve UseColorKeyFlag(Name)
    End If
    
    ' Create surface for image (ignore mask for DirectDraw)
    Set RegisteredSurfaces(Name) = CreateSurfaceFromFile(Filepath)
    ' Do not create mask surface; colorkey will be used instead
    
    ' Apply color key only if requested
    If UseColorKey And Not RegisteredSurfaces(Name) Is Nothing Then
        Dim ddck As DDCOLORKEY
        ddck.low = ColorKey
        ddck.high = ColorKey
        RegisteredSurfaces(Name).SetColorKey DDCKEY_SRCBLT, ddck
        If Name > UBound(UseColorKeyFlag) Then ReDim Preserve UseColorKeyFlag(Name)
        UseColorKeyFlag(Name) = True
    End If
End Sub

Public Sub RenderingEngine_RegisterSpriteFromContext(ByVal Name As Integer, ByVal ImgContext As Long, Optional ByVal MaskContext As Long = 0, Optional ByVal UseColorKey As Boolean = False, Optional ByVal ColorKey As Long = 0)
    ' Resize arrays to accommodate the new sprite
    If Name > UBound(RegisteredSurfaces) Then
        ReDim Preserve RegisteredSurfaces(Name)
        ReDim Preserve RegisteredMaskSurfaces(Name)
    End If
    
    ' Create surfaces from GDI contexts
    Set RegisteredSurfaces(Name) = CreateSurfaceFromContext(ImgContext)
    ' Mask ignored in DirectDraw path
    
    ' Apply color key if requested
    If UseColorKey And Not RegisteredSurfaces(Name) Is Nothing Then
        Dim ddck2 As DDCOLORKEY
        ddck2.low = ColorKey
        ddck2.high = ColorKey
        RegisteredSurfaces(Name).SetColorKey DDCKEY_SRCBLT, ddck2
        If Name > UBound(UseColorKeyFlag) Then ReDim Preserve UseColorKeyFlag(Name)
        UseColorKeyFlag(Name) = True
    End If
End Sub

' Helper to create DirectDraw surface from bitmap file
Private Function CreateSurfaceFromFile(ByVal Filepath As String) As DirectDrawSurface7
    Dim SurfDesc As DDSURFACEDESC2
    Dim surface As DirectDrawSurface7
    Dim hDC As Long, surfaceDC As Long
    Dim hBmp As Long
    Dim bm As BITMAP
    
    ' Load bitmap from file as DIBSection to get real size
    hBmp = LoadImage(0, Filepath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    If hBmp = 0 Then
        Set CreateSurfaceFromFile = Nothing
        Exit Function
    End If
    GetObject hBmp, Len(bm), bm
    
    ' Create offscreen surface sized to bitmap
    ZeroDDSD SurfDesc
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    SurfDesc.lWidth = bm.bmWidth
    SurfDesc.lHeight = bm.bmHeight
    Set surface = DDRAW.CreateSurface(SurfDesc)
    
    ' Get DC for surface and copy bitmap
    surfaceDC = surface.GetDC()
    hDC = CreateCompatibleDC(surfaceDC)
    SelectObject hDC, hBmp
    BitBlt surfaceDC, 0, 0, bm.bmWidth, bm.bmHeight, hDC, 0, 0, vbSrcCopy
    surface.ReleaseDC surfaceDC
    
    ' Cleanup
    DeleteDC hDC
    DeleteObject hBmp
    
    Set CreateSurfaceFromFile = surface
End Function

' Helper to create DirectDraw surface from GDI context
Private Function CreateSurfaceFromContext(ByVal context As Long) As DirectDrawSurface7
    Dim SurfDesc As DDSURFACEDESC2
    Dim surface As DirectDrawSurface7
    Dim hDC As Long, surfaceDC As Long
    
    ' Create offscreen surface (assume 32x32 for now)
    ZeroDDSD SurfDesc
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    SurfDesc.lWidth = 32
    SurfDesc.lHeight = 32
    Set surface = DDRAW.CreateSurface(SurfDesc)
    
    ' Get DC for surface and copy from context
    surfaceDC = surface.GetDC()
    BitBlt surfaceDC, 0, 0, 32, 32, context, 0, 0, vbSrcCopy
    surface.ReleaseDC surfaceDC
    
    Set CreateSurfaceFromContext = surface
End Function

Public Sub RenderingEngine_Draw(Name As Integer, srcRect As Rectangle, destRect As Rectangle)
    Dim imgSurface As DirectDrawSurface7
    Dim srcR As RECT
    Dim destR As RECT
    Dim clippedSrc As RECT
    Dim dx As Long
    Dim dy As Long
    
    ' Get surface
    Set imgSurface = RegisteredSurfaces(Name)
    
    ' Convert Rectangle to RECT
    RectangleToRECT srcRect, srcR
    RectangleToRECT destRect, destR
    
    ' Use helper to clip to backbuffer; skip if fully outside
    If imgSurface Is Nothing Then Exit Sub
    If Not ClipToBackbuffer(srcR, destR.Left, destR.Top, Boundary.Width, Boundary.Height, clippedSrc, dx, dy) Then Exit Sub

    ' Color-keyed or regular blit depending on registration
    Dim flags As Long
    flags = DDBLTFAST_WAIT
    If Name <= UBound(UseColorKeyFlag) Then
        If UseColorKeyFlag(Name) Then flags = flags Or DDBLTFAST_SRCCOLORKEY
    End If
    BackSurface.BltFast dx, dy, imgSurface, clippedSrc, flags
End Sub

Public Sub RenderingEngine_DrawCircle(X As Long, Y As Long, Radius As Long, Color As Long)
    Dim surfaceDC As Long
    Dim pen As Long, brush As Long
    Dim oldPen As Long, oldBrush As Long
    
    ' Get DC for backbuffer
    surfaceDC = BackSurface.GetDC()
    
    ' Create pen and brush
    pen = CreatePen(0, 1, Color)
    brush = CreateSolidBrush(Color)
    
    ' Select them into the DC
    oldPen = SelectObject(surfaceDC, pen)
    oldBrush = SelectObject(surfaceDC, brush)
    
    ' Draw circle
    Ellipse surfaceDC, X - Radius, Y - Radius, X + Radius, Y + Radius
    
    ' Restore original objects
    SelectObject surfaceDC, oldPen
    SelectObject surfaceDC, oldBrush
    
    ' Cleanup
    DeleteObject pen
    DeleteObject brush
    BackSurface.ReleaseDC surfaceDC
End Sub

Public Sub RenderingEngine_BuildParticleAtlas(ByVal Name As Integer)
    Dim atlasImgSurface As DirectDrawSurface7, atlasMaskSurface As DirectDrawSurface7
    Dim SurfDesc As DDSURFACEDESC2
    Dim imgDC As Long, maskDC As Long
    Dim i As Long, j As Long
    Dim Radius As Long
    Dim tileX As Long, tileY As Long
    Dim centerX As Long, centerY As Long
    Dim Color As Long
    Dim brush As Long, oldBrush As Long
    Dim nullPen As Long, oldPen As Long
    
    ' Create atlas surfaces (256x192 pixels)
    ZeroDDSD SurfDesc
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    SurfDesc.lWidth = 256
    SurfDesc.lHeight = 192
    Set atlasImgSurface = DDRAW.CreateSurface(SurfDesc)
    Set atlasMaskSurface = DDRAW.CreateSurface(SurfDesc)
    
    ' Get DCs for both surfaces
    imgDC = atlasImgSurface.GetDC()
    maskDC = atlasMaskSurface.GetDC()
    
    ' Clear img to black, mask to white
    SetBkColor imgDC, RGB(0, 0, 0)
    SetBkColor maskDC, RGB(255, 255, 255)
    SetBkMode imgDC, 1
    SetBkMode maskDC, 1
    
    ' Fill with background colors
    brush = CreateSolidBrush(RGB(0, 0, 0))
    oldBrush = SelectObject(imgDC, brush)
    nullPen = SelectObject(imgDC, GetStockObject(NULL_PEN))
    Rectangle imgDC, 0, 0, 256, 192
    SelectObject imgDC, nullPen
    SelectObject imgDC, oldBrush
    DeleteObject brush
    
    brush = CreateSolidBrush(RGB(255, 255, 255))
    oldBrush = SelectObject(maskDC, brush)
    nullPen = SelectObject(maskDC, GetStockObject(NULL_PEN))
    Rectangle maskDC, 0, 0, 256, 192
    SelectObject maskDC, nullPen
    SelectObject maskDC, oldBrush
    DeleteObject brush
    
    ' Create circles for each radius and color combination
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
            oldBrush = SelectObject(imgDC, brush)
            oldPen = SelectObject(imgDC, nullPen)
            Ellipse imgDC, centerX - drawR, centerY - drawR, centerX + drawR, centerY + drawR
            SelectObject imgDC, oldPen
            SelectObject imgDC, oldBrush
            DeleteObject brush
            
            ' Draw black circle in mask atlas
            brush = CreateSolidBrush(RGB(0, 0, 0))
            oldBrush = SelectObject(maskDC, brush)
            oldPen = SelectObject(maskDC, nullPen)
            Ellipse maskDC, centerX - drawR, centerY - drawR, centerX + drawR, centerY + drawR
            SelectObject maskDC, oldPen
            SelectObject maskDC, oldBrush
            DeleteObject brush
        Next j
    Next i
    
    ' Release DCs
    atlasImgSurface.ReleaseDC imgDC
    atlasMaskSurface.ReleaseDC maskDC
    
    ' Register the atlas surfaces
    RenderingEngine_RegisterSpriteFromContext Name, 0, 0  ' We'll override with our surfaces
    If Name > UBound(RegisteredSurfaces) Then
        ReDim Preserve RegisteredSurfaces(Name)
        ReDim Preserve RegisteredMaskSurfaces(Name)
    End If
    Set RegisteredSurfaces(Name) = atlasImgSurface
    Set RegisteredMaskSurfaces(Name) = atlasMaskSurface
    
    Debug.Print "Particle Atlas created successfully for registry " & Name
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
