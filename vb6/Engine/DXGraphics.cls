VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtDXGSurfaceInfo
    surfPath As String
    surfColorKey As Long
    surf3DSprite As Boolean
    surfNotBitmap As Boolean
End Type

'function to cap framerate
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'type to describe a surface
Private Type DDSurface
    Alive As Boolean
    Surface As DirectDrawSurface7
    width As Long
    height As Long
    is3DSprite As Boolean
End Type

'directx objects
Private DX As New DirectX7
Private DD As DirectDraw7
Private D3D As Direct3D7
Private D3DDev As Direct3DDevice7

'screen width and height
Private screenWidth As Long
Private screenHeight As Long

'holds all the surfaces
Private Surfaces() As DDSurface
Private SurfaceData() As udtDXGSurfaceInfo

'the font object which holds current font data
Private oFontData As New StdFont

'the primary and backbuffer surfaces
Private Primary As DirectDrawSurface7
Private Backbuffer As DirectDrawSurface7

'holds the handle of the container form
Private main_hWnd As Long

'holds the current frame rate
Private zFPS As Double
Private bLagging As Boolean
'

Public Property Get Lagging() As Boolean
    Lagging = bLagging
End Property

Public Sub BeginScene(Optional frameCap As Long = 0)
    CalculateFPS frameCap
    D3DDev.BeginScene
End Sub
Public Sub EndScene()
    D3DDev.EndScene
    Primary.Flip Nothing, DDFLIP_WAIT
    DoEvents
    If Not GameIsActiveWindow Then
        Do Until GameIsActiveWindow
            DoEvents
        Loop
        ReloadAllSurfaces
    End If
End Sub
Private Function GameIsActiveWindow() As Boolean
    GameIsActiveWindow = (DD.TestCooperativeLevel = DD_OK)
End Function


'allow the project to see the framerate
Private Sub CalculateFPS(iCap As Long)
Dim diff As Single
Dim lSleep As Long
Static last As Long

    diff = (DX.TickCount - last)
    last = DX.TickCount
    If diff = 0 Then diff = 1
    zFPS = 1 / (diff / 1000)
    
    bLagging = False
    If iCap Then
        If zFPS > iCap Then
            lSleep = CLng(((1 / iCap) * 1000) - diff)
            If lSleep > 500 Or lSleep < 0 Then lSleep = 0
            Sleep lSleep
        ElseIf zFPS < iCap Then
            bLagging = True
        End If
    End If
    
End Sub
Public Property Get FPS() As Double
FPS = zFPS
End Property



Public Sub SetFont(sFont As String, iSize As Integer, bold As Boolean, italic As Boolean, underline As Boolean, strike As Boolean)
    With oFontData
        .Name = sFont
        .bold = bold
        .italic = italic
        .underline = underline
        .Strikethrough = strike
        .Size = iSize
    End With
    Backbuffer.SetFont oFontData
       
End Sub

Public Sub DrawText(sText As String, X As Long, Y As Long, forecolor As Long, Optional backcolor As Long = -1)
With Backbuffer
    If sText = "" Then Exit Sub
    .SetForeColor forecolor
    If backcolor >= 0 Then
        .SetFontTransparency False
        .SetFontBackColor backcolor
    Else
        .SetFontTransparency True
    End If
       .DrawText X, Y, sText, False
End With
End Sub

'
'  some properties to deal with surfaces
'
Public Property Get SurfaceCount() As Boolean
    SurfaceCount = UBound(Surfaces) + 1
End Property
Public Property Get SurfaceWidth(surfaceID As Long) As Long
    If surfaceID >= 0 And surfaceID <= UBound(Surfaces) Then
        SurfaceWidth = Surfaces(surfaceID).width
    End If
End Property
Public Property Get SurfaceHeight(surfaceID As Long) As Long
    If surfaceID >= 0 And surfaceID <= UBound(Surfaces) Then
        SurfaceHeight = Surfaces(surfaceID).height
    End If
End Property
Public Property Get SurfaceIs3D(surfaceID As Long) As Boolean
    If surfaceID >= 0 And surfaceID <= UBound(Surfaces) Then
        SurfaceIs3D = Surfaces(surfaceID).is3DSprite
    End If
End Property



Public Sub DrawRect(X As Long, Y As Long, width As Long, height As Long, Color As Long)
Dim r As RECT
    With r
        .Left = X
        .Top = Y
        .Right = X + width
        .Bottom = Y + height
    End With
    Backbuffer.BltColorFill r, Color
End Sub


'
' This sub performs various transformations on a d3d surface and blts it
' to the backbuffer or if the surface is a simple dd surface it just blts it to the backbuffer
'
Public Sub DrawSurface(surfaceID As Long, SrcX As Single, SrcY As Single, SrcWidth As Long, _
                       SrcHeight As Long, DestX As Single, DestY As Single, _
                       Optional ByVal DestWidth As Long, Optional ByVal DestHeight As Long, _
                       Optional ByVal angle As Single = 0, Optional ByVal Alpha As Single = 100, _
                       Optional ByVal r As Single = 100, Optional ByVal g As Single = 100, _
                       Optional ByVal b As Single = 100)
On Error GoTo errOut
Dim SrcRect As RECT
Dim DestRect As RECT
Dim TempVerts(3) As D3DTLVERTEX

    If DestWidth = 0 Then DestWidth = SrcWidth
    If DestHeight = 0 Then DestHeight = SrcHeight
    
    With SrcRect
        .Top = SrcY
        .Bottom = SrcY + SrcHeight
        .Left = SrcX
        .Right = SrcX + SrcWidth
    End With
        
    With DestRect
        .Top = DestY
        If .Top > screenHeight Then Exit Sub
        .Bottom = DestY + DestHeight
        If .Bottom < 0 Then Exit Sub
        .Left = DestX
        If .Left > screenWidth Then Exit Sub
        .Right = DestX + DestWidth
        If .Right < 0 Then Exit Sub
    End With
    
    If Surfaces(surfaceID).is3DSprite Then
    
        r = r / 100
        g = g / 100
        b = b / 100
        Alpha = Alpha / 100
    
        Setup3DGeometry TempVerts, surfaceID, SrcRect, DestRect, r, g, b, Alpha, angle
        With D3DDev
            .SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
            .SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
            .SetRenderState D3DRENDERSTATE_COLORKEYBLENDENABLE, True
            .SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
            .SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
            .SetRenderState D3DRENDERSTATE_TEXTUREFACTOR, DX.CreateColorRGBA(1, 1, 1, Alpha)
            .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTA_TFACTOR
            .SetTexture 0, Surfaces(surfaceID).Surface
            .SetTextureStageState 0, D3DTSS_MIPFILTER, 3
            .DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, TempVerts(0), 4, D3DDP_DEFAULT
            .SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
        End With
        
    Else
    
        If DestRect.Left < 0 Then
            SrcRect.Left = SrcRect.Left + Abs(DestRect.Left)
            DestRect.Left = 0
        End If
        If DestRect.Right > screenWidth Then
            SrcRect.Right = SrcRect.Right - (DestRect.Right - screenWidth)
            DestRect.Right = screenWidth
        End If
        If DestRect.Top < 0 Then
            SrcRect.Top = SrcRect.Top + Abs(DestRect.Top)
            DestRect.Top = 0
        End If
        If DestRect.Bottom > screenHeight Then
            SrcRect.Bottom = SrcRect.Bottom - (DestRect.Bottom - screenHeight)
            DestRect.Bottom = screenHeight
        End If
        Backbuffer.Blt DestRect, Surfaces(surfaceID).Surface, SrcRect, DDBLT_KEYSRC Or DDBLT_WAIT
        
    End If

errOut:
End Sub

'
'
'
'
'
' This function initilizes directdraw7 and direct3d7 and returns true if successfull
'
Public Function Initialize(hWnd As Long, width As Long, height As Long, depth As Long)
On Error GoTo errOut
Dim ddsdp As DDSURFACEDESC2
Dim caps As DDSCAPS2

    main_hWnd = hWnd
    screenWidth = width
    screenHeight = height

    Set DD = DX.DirectDrawCreate("")
    DD.SetCooperativeLevel main_hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT
    DD.SetDisplayMode width, height, depth, 0, DDSDM_DEFAULT
    
    With ddsdp
        .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_VIDEOMEMORY
        .lBackBufferCount = 1
    End With
    Set Primary = DD.CreateSurface(ddsdp)
        
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set Backbuffer = Primary.GetAttachedSurface(caps)
    
    Backbuffer.SetForeColor vbWhite
    
    Set D3D = DD.GetDirect3D
    Set D3DDev = D3D.CreateDevice("IID_IDirect3DHALDevice", Backbuffer)
    
    ReDim Surfaces(1)
    ReDim SurfaceData(1)

Initialize = True
errOut:
End Function
'

'
' creates a new surface from a file and returns the position of the new surface in the array
'
Public Function CreateSurface(ByVal sPath As String, ByVal colorKey As Long, ByVal b3DSprite As Boolean, Optional ByVal notBitmap As Boolean = False, Optional ByVal lReload As Long = -1) As Long
On Error GoTo errOut
Dim SDesc As DDSURFACEDESC2
Dim CKey As DDCOLORKEY
Dim cursurf As Long

    If lReload > 0 Then
        cursurf = lReload
        CreateSurface = -2
        With SurfaceData(cursurf)
            sPath = .surfPath
            colorKey = .surfColorKey
            b3DSprite = .surf3DSprite
            notBitmap = .surfNotBitmap
        End With
    Else
        If Surfaces(UBound(Surfaces)).Alive Then ReDim Preserve Surfaces(UBound(Surfaces) + 1)
        ReDim Preserve SurfaceData(UBound(Surfaces))
        cursurf = UBound(Surfaces)
        CreateSurface = cursurf
        With SurfaceData(cursurf)
            .surfPath = sPath
            .surfColorKey = colorKey
            .surf3DSprite = b3DSprite
            .surfNotBitmap = notBitmap
        End With
    End If
    
    If notBitmap Then CreateSurface_saveGIF sPath

    If b3DSprite Then
    
        With SDesc
            .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
            .ddsCaps.lCaps = DDSCAPS_TEXTURE
            .ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
            .ddckCKSrcBlt.high = colorKey
            .ddckCKSrcBlt.low = colorKey
        End With
        Set Surfaces(cursurf).Surface = DD.CreateSurfaceFromFile(sPath, SDesc)
        With Surfaces(cursurf)
            .width = SDesc.lWidth
            .height = SDesc.lHeight
            .is3DSprite = True
            .Alive = True
        End With
        
    Else
    
        SDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        SDesc.ddsCaps.lCaps = DDSCAPS_VIDEOMEMORY Or DDSCAPS_OFFSCREENPLAIN
        Set Surfaces(cursurf).Surface = DD.CreateSurfaceFromFile(sPath, SDesc)
        With Surfaces(cursurf)
            CKey.low = colorKey
            CKey.high = colorKey
            .Surface.SetColorKey DDCKEY_SRCBLT, CKey
            .Surface.SetForeColor vbBlack
            .width = SDesc.lWidth
            .height = SDesc.lHeight
            .is3DSprite = False
            .Alive = True
        End With
        
    End If
    
    If notBitmap Then Kill sPath

Exit Function
errOut:
With SurfaceData(cursurf)
.surfPath = sPath
.surfColorKey = colorKey
.surf3DSprite = b3DSprite
.surfNotBitmap = notBitmap
End With
With Surfaces(cursurf)
Set .Surface = Nothing
.height = 0
.width = 0
.is3DSprite = False
.Alive = False
End With
CreateSurface = -1
End Function
Private Sub CreateSurface_saveGIF(sPath As String)
    frmMain.picLoadGIF.Picture = LoadPicture(sPath)
    sPath = sPath & "z.bmp"
    SavePicture frmMain.picLoadGIF.Picture, sPath
End Sub
'

'
' this sub clears a specific surface from memory
'
Public Sub DestroySurface(ByVal lSurfID As Long)

    If lSurfID <= 0 Or lSurfID > UBound(Surfaces) Then Exit Sub

    With Surfaces(lSurfID)
        Set .Surface = Nothing
        .is3DSprite = False
        .width = 0
        .height = 0
        .Alive = False
    End With
    With SurfaceData(lSurfID)
        .surf3DSprite = False
        .surfColorKey = 0
        .surfNotBitmap = False
        .surfPath = ""
    End With
    
End Sub

'
'  this sub reloads all the that have already been loaded
'
Public Sub ReloadAllSurfaces()
Dim i As Long
    
    DD.RestoreAllSurfaces
    For i = 0 To UBound(Surfaces)
        If Surfaces(i).Alive Then CreateSurface "", 0, False, False, i
    Next i
    
End Sub

'
' this sub clears all surfaces from memory
'
Private Sub DestroySurfaces()
Dim i As Long

    For i = 0 To UBound(Surfaces)
        With Surfaces(i)
            Set .Surface = Nothing
            .is3DSprite = False
            .width = 0
            .height = 0
            .Alive = False
        End With
        With SurfaceData(i)
            .surf3DSprite = False
            .surfColorKey = 0
            .surfNotBitmap = False
            .surfPath = ""
        End With
    Next i
    Erase Surfaces
    
End Sub



'
'  This sub calculates the 3D geometry for a d3d sprite according to
' position, size, color, alpha and stores the 4 created vertexes in the array passed
'
Private Sub Setup3DGeometry(Verts() As D3DTLVERTEX, surfaceID As Long, Src As RECT, Dest As RECT, r As Single, g As Single, b As Single, a As Single, ByVal angle As Single)
Dim SurfW As Single
Dim SurfH As Single
Dim XCenter As Single
Dim YCenter As Single
Dim Radius As Single
Dim XCor As Single
Dim YCor As Single

    SurfW = Surfaces(surfaceID).width
    SurfH = Surfaces(surfaceID).height
    XCenter = Dest.Left + (Dest.Right - Dest.Left - 1) / 2
    YCenter = Dest.Top + (Dest.Bottom - Dest.Top - 1) / 2
    
    'change the angle from degrees to rads
    angle = (angle + 90) * 3.14159265358979 / 180
    
    'set up bottom left vertex
    XCor = XCenter + (Dest.Left - XCenter) * Sin(angle) + (Dest.Bottom - YCenter) * Cos(angle)
    YCor = YCenter + (Dest.Bottom - YCenter) * Sin(angle) - (Dest.Left - XCenter) * Cos(angle)
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, DX.CreateColorRGBA(r, g, b, a), 0, Src.Left / SurfW, (Src.Bottom + 1) / SurfH, Verts(0)
    
    'set up top left vertex
    XCor = XCenter + (Dest.Left - XCenter) * Sin(angle) + (Dest.Top - YCenter) * Cos(angle)
    YCor = YCenter + (Dest.Top - YCenter) * Sin(angle) - (Dest.Left - XCenter) * Cos(angle)
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, DX.CreateColorRGBA(r, g, b, a), 0, Src.Left / SurfW, Src.Top / SurfH, Verts(1)
    
    'set up bottom right vertex
    XCor = XCenter + (Dest.Right - XCenter) * Sin(angle) + (Dest.Bottom - YCenter) * Cos(angle)
    YCor = YCenter + (Dest.Bottom - YCenter) * Sin(angle) - (Dest.Right - XCenter) * Cos(angle)
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, DX.CreateColorRGBA(r, g, b, a), 0, (Src.Right + 1) / SurfW, (Src.Bottom + 1) / SurfH, Verts(2)
    
    'set up top right vertex
    XCor = XCenter + (Dest.Right - XCenter) * Sin(angle) + (Dest.Top - YCenter) * Cos(angle)
    YCor = YCenter + (Dest.Top - YCenter) * Sin(angle) - (Dest.Right - XCenter) * Cos(angle)
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, DX.CreateColorRGBA(r, g, b, a), 0, (Src.Right + 1) / SurfW, Src.Top / SurfH, Verts(3)

End Sub
'

'
'  restores original display mode, returns control to windows, and cleans up after DX
'
Private Sub Class_Terminate()
    DestroySurfaces
    DD.SetCooperativeLevel main_hWnd, DDSCL_NORMAL
    DD.RestoreDisplayMode
    Set oFontData = Nothing
    Set Backbuffer = Nothing
    Set Primary = Nothing
    Set D3DDev = Nothing
    Set D3D = Nothing
    Set DD = Nothing
    Set DX = Nothing
End Sub
