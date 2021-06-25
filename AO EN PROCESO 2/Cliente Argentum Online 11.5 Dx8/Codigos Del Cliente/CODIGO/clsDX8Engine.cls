VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim font_count As Long
Dim font_last As Long

Private font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Dim Texture As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double
Private particletimer As Single
Public engineBaseSpeed As Single
Private lFrameTimer As Long
Private lFrameLimiter As Long
Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop As Integer
Private MainViewLeft As Integer
Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte
Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer
Private Const GrhFogata As Integer = 1521

Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

'particulas----------
Dim base_tile_size As Integer

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Private Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    Speed As Single
    life_counter As Long
End Type
 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long
 
    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte
 
    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    Speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

 
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
'/particulas----------------------


Dim valoreBlur As Long
Dim dimeTex As Long
Dim tex As Direct3DTexture8
Dim D3DbackBuffer As Direct3DSurface8
Dim zTarget As Direct3DSurface8
Dim stencil As Direct3DSurface8
Dim superTex As Direct3DSurface8
Dim blur As Boolean
Public blur_factor As Byte

Dim bump_map_texture As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported As Boolean
Dim bump_map_powa As Boolean
'Sets a Grh animation to loop indefinitely.







#Const HARDCODED = False 'True ' == MÁS FPS ^^






Public Function change_day_effect(ByVal o_r As Long, ByVal og As Long, ByVal ob As Long, ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long) As Long
Dim r As Long
Dim g As Long
Dim B As Long
r = o_r
g = og
B = ob
    
    If o_r > r_ Then
        r = r - 1
    ElseIf o_r < r_ Then
        r = r + 1
    End If

    If og > g_ Then
        g = g - 1
    ElseIf og < r Then
        g = g + 1
    End If

    If ob > b_ Then
        B = B - 1
    ElseIf og < r Then
        B = B + 1
    End If
If r > 255 Then r = 255
If g > 255 Then g = 255
If B > 255 Then B = 255
day_r_old = r
day_g_old = g
day_b_old = B

If r < 0 Or g < 0 Or B < 0 Then r = 1: g = 1: B = 1
change_day_effect = ARGB(r, g, B, 255)

End Function

Public Sub setup_ambient()

'Noche 87, 61, 43
luz_dia(0).r = 87
luz_dia(0).g = 61
luz_dia(0).B = 43
luz_dia(1).r = 87
luz_dia(1).g = 61
luz_dia(1).B = 43
luz_dia(2).r = 87
luz_dia(2).g = 61
luz_dia(2).B = 43
luz_dia(3).r = 87
luz_dia(3).g = 61
luz_dia(3).B = 43
'4 am 124,117,91
luz_dia(4).r = 124
luz_dia(4).g = 127
luz_dia(4).B = 91
'5,6 am 143,137,135
luz_dia(5).r = 143
luz_dia(5).g = 137
luz_dia(5).B = 135
luz_dia(6).r = 143
luz_dia(6).g = 137
luz_dia(6).B = 135
'7 am 212,205,207
luz_dia(7).r = 212
luz_dia(7).g = 205
luz_dia(7).B = 207
luz_dia(8).r = 212
luz_dia(8).g = 205
luz_dia(8).B = 207
luz_dia(9).r = 212
luz_dia(9).g = 205
luz_dia(9).B = 207
luz_dia(10).r = 212
luz_dia(10).g = 205
luz_dia(10).B = 207
luz_dia(11).r = 212
luz_dia(11).g = 205
luz_dia(11).B = 207
luz_dia(12).r = 212
luz_dia(12).g = 205
luz_dia(12).B = 207
'Dia 255, 255, 255
luz_dia(12).r = 255
luz_dia(12).g = 255
luz_dia(12).B = 255
luz_dia(13).r = 255
luz_dia(13).g = 255
luz_dia(13).B = 255
'Medio Dia 255, 200, 255
luz_dia(14).r = 255
luz_dia(14).g = 250
luz_dia(14).B = 255
luz_dia(15).r = 255
luz_dia(15).g = 240
luz_dia(15).B = 255
luz_dia(16).r = 255
luz_dia(16).g = 230
luz_dia(16).B = 255
'17/18 0, 100, 255
luz_dia(17).r = 230
luz_dia(17).g = 230
luz_dia(17).B = 255
'18/19 0, 100, 255
luz_dia(18).r = 230
luz_dia(18).g = 230
luz_dia(18).B = 255
'19/20 156, 142, 83
luz_dia(19).r = 156
luz_dia(19).g = 142
luz_dia(19).B = 83
luz_dia(20).r = 87
luz_dia(20).g = 61
luz_dia(20).B = 43
luz_dia(21).r = 87
luz_dia(21).g = 61
luz_dia(21).B = 43
luz_dia(22).r = 87
luz_dia(22).g = 61
luz_dia(22).B = 43
luz_dia(23).r = 87
luz_dia(23).g = 61
luz_dia(23).B = 43
luz_dia(24).r = 87
luz_dia(24).g = 61
luz_dia(24).B = 43
            
End Sub


Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function



Public Sub Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False, Optional ByVal Nombre As Byte, Optional ByVal rojo As Byte, Optional ByVal verde As Byte, Optional ByVal azul As Byte)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + Height
    TextRect.Right = left + Width
        
    If shadow Then
        ShadowRect.top = top - 1
        ShadowRect.left = left - 2
        ShadowRect.bottom = (top + Height) - 1
        ShadowRect.Right = (left + Width) - 2
    End If
D3DX.DrawText font, color, Text, TextRect, format
End Sub
Public Sub Text_Render_Inventario(ByVal font As D3DXFont, ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
' editado por: juanjo
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + Height
    TextRect.Right = left + Width
        
    If shadow Then
        ShadowRect.top = top - 1
        ShadowRect.left = left - 2
        ShadowRect.bottom = (top + Height) - 1
        ShadowRect.Right = (left + Width) - 2
       'D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
        D3DX.DrawText font, D3DColorXRGB(255, 255, 255), Text, ShadowRect, format
    End If
    
    D3DX.DrawText font, color, Text, TextRect, format
    
End Sub
Public Sub Text_Render_ext(ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal Center As Boolean = False)
    If Center = True Then
        Call Text_Render_Inventario(font_list(1), Text, top, left, Width, Height, color, fa_center, shadow)
    Else
        Call Text_Render_Inventario(font_list(1), Text, top, left, Width, Height, color, DT_TOP Or DT_LEFT, shadow)
    End If
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
  MakeVector.X = X
  MakeVector.Y = Y
  MakeVector.Z = Z
End Function

Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo ErrHandler:

    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd
    End With
    DispMode.format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format
        End If
    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format
    End If
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
                                                            
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.012
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    ScrollPixelsPerFrameX = 9
    ScrollPixelsPerFrameY = 9
    
    UserPos.X = 50
    UserPos.Y = 50
    
    MinXBorder = XMinMapSize + (frmMain.renderer.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmMain.renderer.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmMain.renderer.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmMain.renderer.ScaleHeight / 64)

     Texto.Engine_Init_FontSettings
        Texto.Engine_Init_FontTextures
    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    'motion blur
    Set D3DbackBuffer = D3DDevice.GetRenderTarget
    Set zTarget = D3DDevice.GetDepthStencilSurface
    Set stencil = D3DDevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    Set tex = D3DX.CreateTexture(D3DDevice, dimeTex, dimeTex, 1, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
    Set superTex = tex.GetSurfaceLevel(0)
    blur_factor = 10
    'bump mapping
    
    Font_Create "Tahoma", 8, True, 0
    Font_Create "Morpheus", 20, True, 0
    'Font_Create "Verdana", 8, False, 0
    
bRunning = True
Exit Sub
ErrHandler:
Debug.Print "Error Number Returned: " & Err.Number
bRunning = False
End Sub

Public Sub Engine_Deinit()
Erase MapData
Erase charlist
Set D3DDevice = Nothing
Set D3D = Nothing
Set dX = Nothing
End
End Sub


Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Sub Engine_ActFPS()
    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
End Sub

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Public Sub Draw_GrhIndex(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False)
    Dim CurrentGrhIndex As Integer
    If Grh.grhindex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1
               
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
   
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
 
    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
 
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If
   
 
 
    Device_Box_Textured_Render CurrentGrhIndex, _
        X, Y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        MapData(map_x, map_y).light_value, _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        Alpha _
        , angle, Invert_x, Invert_y
'exits:
 
End Sub

Public Sub Engine_Blur_Set(ByVal bool As Boolean)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
blur = bool
If blur = True Then
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
Else
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
End If
End Sub

Public Sub Engine_Blur_Toggle()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
blur = Not blur
If blur = True Then
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
Else
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
End If
End Sub

Public Sub Render()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************

D3DDevice.BeginScene

If blur Then
    D3DDevice.SetRenderTarget superTex, stencil, 0
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    D3DDevice.SetRenderTarget D3DbackBuffer, zTarget, 0
Else
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
End If
    ShowNextFrame
    
    If blur Then
        D3DDevice.SetTexture 0, tex
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
        D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(blur_factor, 255, 255, 255)
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertList(0), Len(vertList(0))
    End If

    Text_Render font_list(1), Tiempo, 55, 10, 70, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True 'Thusing

    Text_Render font_list(1), "FPS: " & FPS, 25, 10, 60, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    engine.Engine_ActFPS
    
    
    Text_Render font_list(1), "Coord: " & UserMap & "," & UserPos.X & "," & UserPos.Y, 40, 10, 90, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    Call Dialogos.Render
    Call DialogosClanes.Draw
D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

lFrameLimiter = GetTickCount
FramesPerSecCounter = FramesPerSecCounter + 1
timerElapsedTime = GetElapsedTime()
timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
particletimer = timerElapsedTime * 0.05
End Sub



Sub ShowNextFrame()
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        engine.Engine_ActFPS
End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim minY                As Integer  'Start Y pos on current map
    Dim maxY                As Integer  'End Y pos on current map
    Dim minX                As Integer  'Start X pos on current map
    Dim maxX                As Integer  'End X pos on current map
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    

    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1 **********************************
            
#If HARDCODED = True Then
                    If MapData(X, Y).Graphic(1).Started = 1 Then
                        MapData(X, Y).Graphic(1).FrameCounter = MapData(X, Y).Graphic(1).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(1).grhindex).NumFrames / MapData(X, Y).Graphic(1).Speed)
                        If MapData(X, Y).Graphic(1).FrameCounter > GrhData(MapData(X, Y).Graphic(1).grhindex).NumFrames Then
                            MapData(X, Y).Graphic(1).FrameCounter = (MapData(X, Y).Graphic(1).FrameCounter Mod GrhData(MapData(X, Y).Graphic(1).grhindex).NumFrames) + 1
                            
                            If MapData(X, Y).Graphic(1).Loops <> -1 Then
                                If MapData(X, Y).Graphic(1).Loops > 0 Then
                                    MapData(X, Y).Graphic(1).Loops = MapData(X, Y).Graphic(1).Loops - 1
                                Else
                                    MapData(X, Y).Graphic(1).Started = 0
                                End If
                            End If
                        End If
                    End If

                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(1).grhindex).Frames(MapData(X, Y).Graphic(1).FrameCounter)

                Device_Box_Textured_Render CurrentGrhIndex, _
                    (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, _
                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    MapData(X, Y).light_value, _
                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    False _
                    , 0
#Else
                Call Draw_Grh(MapData(X, Y).Graphic(1), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        0, 1, , X, Y)
                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, , X, Y)
                End If
#End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
                'Layer 2 **********************************
    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************

                'Object Layer **********************************
                If .ObjGrh.grhindex <> 0 Then
                
                    Call Draw_Grh(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)

                End If
                '***********************************************
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
                End If
                '*************************************************
                
                
                
 'Layer 3 *****************************************
If .Graphic(3).grhindex = 735 Or .Graphic(3).grhindex >= 6994 And .Graphic(3).grhindex <= 7002 Then
   If Abs(UserPos.X - X) < 4 And (Abs(UserPos.Y - Y)) < 4 Then
       Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, True, X, Y)
   Else
       Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
   End If
 
Else
 
   If .Graphic(3).grhindex <> 0 Then
       Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
   End If
End If
 
'************************************************
                

            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5
                    
                    
                    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    

    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            With MapData(X, Y)
                '***********************************************
                If .particle_group > 0 Then
                
                Particle_Group_Render .particle_group, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY
                
                End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
                    
                    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                    
    If Not bTecho Then
        'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).grhindex Then
                    Call Draw_Grh(MapData(X, Y).Graphic(4), _
                        ScreenX * 32 + PixelOffsetX, _
                        ScreenY * 32 + PixelOffsetY, _
                        1, 1, , X, Y)
                End If
                '**********************************
                
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    '//Dibuja Mapa.
    If PuedeMapa = True Then
        Contador = Contador + 1
        If Contador <= 250 Then
            Call engine.Text_Render(font_list(2), NombreMapa, 45, 225, 400, 200, &HFFFFFFFF, DT_TOP Or DT_LEFT, False)
        Else
            Contador = 0
            PuedeMapa = False
        End If
    End If
    '//Dibuja Mapa.
    
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Function Geometry_Create_TLVertex2(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
'mz
Geometry_Create_TLVertex2.X = X
Geometry_Create_TLVertex2.Y = Y
Geometry_Create_TLVertex2.Z = Z
Geometry_Create_TLVertex2.rhw = rhw
Geometry_Create_TLVertex2.color = color
Geometry_Create_TLVertex2.Specular = Specular
Geometry_Create_TLVertex2.tu1 = tu1
Geometry_Create_TLVertex2.tv1 = tv1
Geometry_Create_TLVertex2.tu2 = tu2
Geometry_Create_TLVertex2.tv2 = tv2
End Function
Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    Dim auxr As RECT
   
    If angle <> 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left) / 2
        y_center = dest.top + (dest.bottom - dest.top) / 2
       
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
       
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
   
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
   
    auxr = src
   
    If angle < 0 Then
        src.left = auxr.Right
        src.Right = auxr.left
    End If
   
    If Invert_x Then
        src.left = auxr.Right
        src.Right = auxr.left
    End If
   
    If Invert_y Then
        src.top = auxr.bottom
        src.bottom = auxr.top
    End If
       
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
   
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / Textures_Width, src.top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
   
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
   
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If
 
End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    'If Not char_current_blind Then
    '    If (light_value(0) = 0) Then light_value(0) = 0
    '    If (light_value(1) = 0) Then light_value(1) = 0
    '    If (light_value(2) = 0) Then light_value(2) = 0
    '    If (light_value(3) = 0) Then light_value(3) = 0
    'Else
        light_value(0) = &HFFFFFFFF 'blind_color
        light_value(1) = &HFFFFFFFF 'blind_color
        light_value(2) = &HFFFFFFFF 'blind_color
        light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
   
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
   
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
   
    'If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = base_light
        If (light_value(1) = 0) Then light_value(1) = base_light
        If (light_value(2) = 0) Then light_value(2) = base_light
        If (light_value(3) = 0) Then light_value(3) = base_light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
       
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
       
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .left = dest_x
        .Right = dest_x + src_width
        .top = dest_y
    End With
   
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle, Invert_x, Invert_y
   
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
   
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
   
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
   
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
End Sub


Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
            If Not .invisible Then
                'Draw Body

                            If .Body.Walk(.Heading).grhindex Then _
                                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                                
                            If .Head.Head(.Heading).grhindex Then
                                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, , X, Y)
                                If .Casco.Head(.Heading).grhindex Then _
                                    Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, , X, Y)
                                If .Arma.WeaponWalk(.Heading).grhindex Then _
                                    Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                                If .Escudo.ShieldWalk(.Heading).grhindex Then _
                                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                            End If
                            
                    'Draw name over head
                    Dim colorclan As Long
                    If Nombres Then
                        If Len(.Nombre) > 0 Then
                            Pos = InStr(.Nombre, "<")
                            If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            If .priv = 0 Then
                            
                            If UserCounters.Alineacion = 1 Then
                                color = D3DColorXRGB(255, 0, 0)
                                colorclan = D3DColorXRGB(255, 0, 0)
                            ElseIf UserCounters.Alineacion = 2 Then
                                color = D3DColorXRGB(40, 40, 255)
                                colorclan = D3DColorXRGB(40, 40, 255)
                            ElseIf UserCounters.Alineacion = 0 Then
                                color = D3DColorXRGB(124, 124, 124)
                                colorclan = D3DColorXRGB(124, 124, 124)
                            End If
                            Else
                                color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).B)
                            End If

                            If .priv = 1 Then 'CONSE
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, "<Consejero>" & line, color, , DT_CENTER)

                            ElseIf .priv = 2 Then 'SEMI DIOS
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, "<Semi-Dios>" & line, color, , DT_CENTER)

                            ElseIf .priv = 3 Then 'DIOS
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, "<Dios>" & line, color, , DT_CENTER)

                            ElseIf .priv = 4 Then 'ADMIN
                            Call Texto.Engine_Text_Draw(PixelOffsetX + 16, PixelOffsetY + 45, "<Creador Del AO>" & line, color, , DT_CENTER)

                            End If
                            
 'Nick
                                                line = left$(.Nombre, Pos - 2)
                                                Call Texto.Engine_Text_Draw(PixelOffsetX + 15, PixelOffsetY + 30, line, color, , DT_CENTER)
                            
                                                'Clan
                                                line = mid$(.Nombre, Pos)
                                                Call Texto.Engine_Text_Draw(PixelOffsetX + 15, PixelOffsetY + 45, line, color, , DT_CENTER)
                                                End If
                    End If
                End If
            'End If

        ''Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X - 168, PixelOffsetY + .Body.HeadOffset.Y, CharIndex)
        
        'particulas para meditar
        Dim i As Integer
        If .particle_count > 0 Then
            For i = 1 To .particle_count
                If .particle_group(i) > 0 Then
                    Particle_Group_Render .particle_group(i), PixelOffsetX, PixelOffsetY
                End If
            Next i
        End If
        
        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.Fx, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, True)

            If .Fx.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Public Sub Start()

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 2) As Integer
Dim loopc As Integer


    day_r_old = luz_dia(Hour(Time)).r
    day_g_old = luz_dia(Hour(Time)).g
    day_b_old = luz_dia(Hour(Time)).B
    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
    
DoEvents
vertList(0).X = 0: vertList(0).Y = 0
vertList(1).X = 800: vertList(1).Y = 0
vertList(2).X = 0: vertList(2).Y = 600
vertList(3).X = 800: vertList(3).Y = 600
vertList(0).rhw = 1
vertList(1).rhw = 1
vertList(2).rhw = 1
vertList(3).rhw = 1
vertList(0).color = D3DColorXRGB(255, 255, 255)
vertList(1).color = D3DColorXRGB(255, 255, 255)
vertList(2).color = D3DColorXRGB(255, 255, 255)
vertList(3).color = D3DColorXRGB(255, 255, 255)
vertList(0).tu = 0
vertList(0).tv = 0
vertList(1).tu = 1
vertList(1).tv = 0
vertList(2).tu = 0
vertList(2).tv = 1
vertList(3).tu = 1
vertList(3).tv = 1
On Error Resume Next ' putos timers!!!!!
Do While prgRun

If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
    CheckKeys
    engine.Render
    engine.DrawInv
Else
    Sleep 10&
End If
DoEvents
    Rem Limitar FPS
'    While (GetTickCount - lFrameLimiter) < FramesPerSecCounter
'        Sleep 1
'    Wend
'    While GetTickCount - lFrameLimiter < 55
'       Sleep 5
'    Wend

        esttick = GetTickCount
        For loopc = 1 To UBound(timers)
            timers(loopc) = timers(loopc) + (esttick - ulttick)
            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If
            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If
        Next loopc
        ulttick = GetTickCount

        DoEvents
Loop
engine.Engine_Deinit

    EngineRun = False
    frmCargando.Show

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    

    
End
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = Fx
        
        If .FxIndex > 0 Then
            Call InitGrh(.Fx, FxData(Fx).Animacion)
        
            .Fx.Loops = Loops
        End If
    End With
End Sub

Public Sub Char_Move_by_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Public Sub Char_Move_by_Pos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then
    Call Dialogos.RemoveDialog(CharIndex)
    'Call Hits.RemoveHit(CharIndex)
    End If
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

'Public Sub FXMake( _
ByVal Index As Integer, _
ByVal map_y As Integer, _
ByVal map_x As Integer, _
ByVal tipo As fxs_type)
'    Dim datos(2) As Byte
'    ReDim Preserve fxs_list(0 To Index)
    
'    MapData(map_x, map_y).FXS_group = Index
'    fxs_list(Index).Set_fx tipo, 0, datos()
'End Sub

Private Function CreateColorVal(A As Integer, r As Integer, g As Integer, B As Integer) As D3DCOLORVALUE
    CreateColorVal.A = A
    CreateColorVal.r = r
    CreateColorVal.g = g
    CreateColorVal.B = B
End Function

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, g As Integer, B As Integer, A As Integer
    r = 127 * Vec.X + 128
    g = 127 * Vec.Y + 128
    B = 127 * Vec.Z + 128
    A = 255 * fHeight
    VectorToRGBA = D3DColorARGB(A, r, g, B)
End Function



Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addy As Byte
    Dim addx As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        


        Select Case Heading
            Case NORTH
                addy = -1
        
            Case EAST
                addx = 1
        
            Case SOUTH
                addy = 1
            
            Case WEST
                addx = -1
        End Select
        
        nX = map_x + addx
        nY = map_y + addy
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'menduz
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(0) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(1) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(2) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(3) = color
                
    'For ia = 0 To 6.5 Step 0.05
    '    For i = 1 To light_list(light_index).range / 2
    '        x = (Sin(ia) * i) + light_list(light_index).map_x
    '        y = (Cos(ia) * i) + light_list(light_index).map_y
    '        'Debug.Print x; y
    '        If InMapBounds(x, y) Then
    '            If i = (light_list(light_index).range / 2) Then
    '                If MapData(x - 1, y - 1).light_value(2) <> color Then MapData(x, y).light_value(0) = 0
    '                If MapData(x + 1, y + 1).light_value(0) <> color Then MapData(x, y).light_value(2) = 0
    '                If MapData(x + 1, y - 1).light_value(1) <> color Then MapData(x, y).light_value(1) = 0
    '                If MapData(x - 1, y + 1).light_value(3) <> color Then MapData(x, y).light_value(3) = 0
    '                If MapData(x - 1, y).light_value(0) = color Then MapData(x, y).light_value(3) = color
    '                If MapData(x + 1, y).light_value(1) = color And MapData(x + 1, y).light_value(2) = color Then
    '                    MapData(x, y).light_value(0) = &HFFFF0FFF
    '                    MapData(x, y).light_value(3) = &HFFFF0FFF
    '                End If
    '                If MapData(x, y - 1).light_value(2) = color Then MapData(x, y).light_value(3) = color
    '                If MapData(x, y + 1).light_value(3) = color Then MapData(x, y).light_value(1) = color''

    '            Else
    '                MapData(x, y).light_value(0) = &HFFFFFFFF
    '                MapData(x, y).light_value(1) = &HFFFFFFFF
    '                MapData(x, y).light_value(2) = &HFFFFFFFF
    '                MapData(x, y).light_value(3) = &HFFFFFFFF
    '            End If
    '        End If
    '    Next i
    'Next ia
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = color
            MapData(X, min_y).light_value(2) = color
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = color
            MapData(X, max_y).light_value(3) = color
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = color
            MapData(min_x, Y).light_value(3) = color
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = color
            MapData(max_x, Y).light_value(1) = color
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = color
                MapData(X, Y).light_value(1) = color
                MapData(X, Y).light_value(2) = color
                MapData(X, Y).light_value(3) = color
            End If
        Next Y
    Next X
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub

Public Function RemoveCharAparence(ByVal CharIndex As Integer, ByVal Muerto As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As E_Heading, _
                        ByVal tempint As Integer, ByVal tempweap As Integer, ByVal tempesc As Integer, _
                        ByVal Fx As Integer, ByVal Loops As Integer)
                        
            CharIndex = CharIndex
            charlist(CharIndex).Muerto = Muerto = 500
            charlist(CharIndex).Body = BodyData(Body)
            charlist(CharIndex).Head = HeadData(Head)
            charlist(CharIndex).Heading = Heading
            charlist(CharIndex).FxIndex = Fx
            charlist(CharIndex).Fx.Loops = Loops
            tempint = tempint
            If tempint <> 0 Then charlist(CharIndex).Arma = WeaponAnimData(tempint)
            tempint = tempweap
            If tempint <> 0 Then charlist(CharIndex).Escudo = ShieldAnimData(tempint)
            tempint = tempesc
            If tempint <> 0 Then charlist(CharIndex).Casco = CascoAnimData(tempint)
        Call SetCharacterFx(CharIndex, charlist(CharIndex).FxIndex, charlist(CharIndex).Fx.Loops)
            
            Call RefreshAllChars
End Function

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)

    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = Y + Height
        .left = X
        .Right = X + Width
        .top = Y
    End With
    
    With Outline
        .bottom = Y + Height + 2
        .left = X - 2
        .Right = X + Width + 2
        .top = Y - 2
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub
Public Sub DrawInv()
    Static re As RECT
    re.left = 0
    re.top = 0
    re.bottom = 160
    re.Right = 160
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
    Inventario.DrawInventory
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmMain.picInv.hWnd, ByVal 0
End Sub

'----------------------------------------PARTICULAS---------------------------------
Private Sub Convert_Heading_to_Direction(ByVal Heading As Long, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addy As Long
    Dim addx As Long
    
    'Figure out which way to move
    Select Case Heading
        Case 1
            addy = -1
    
        Case 2
            addx = 1
    
        Case 3
            addy = 1
            
        Case 4
            addx = -1
    
    End Select
    
    direction_x = direction_x + addx
    direction_y = direction_y + addy
End Sub
 

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
    
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
    End If
End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new particle effect
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    MapData(map_x, map_y).particle_group = particle_group_index
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/24/2003
'
'**************************************************************

    If no_move = False Then
                If temp_particle.alive_counter = 0 Then
                    InitGrh temp_particle.Grh, grh_index, alpha_blend
                    temp_particle.X = RandomNumber(x1, x2)
                    temp_particle.Y = RandomNumber(y1, y2)
                    temp_particle.vector_x = RandomNumber(vecx1, vecx2)
                    temp_particle.vector_y = RandomNumber(vecy1, vecy2)
                    temp_particle.angle = angle
                    temp_particle.alive_counter = RandomNumber(life1, life2)
                    temp_particle.friction = fric
                Else
                    'Continue old particle
                    'Do gravity
                    If gravity = True Then
                        temp_particle.vector_y = temp_particle.vector_y + grav_strength
                        If temp_particle.Y > 0 Then
                            'bounce
                            temp_particle.vector_y = bounce_strength
                        End If
                    End If
                    'Do rotation
                   If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                                
                    If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
                End If

        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.grhindex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
 
    'Draw it
    If temp_particle.Grh.grhindex Then
        Grh_Render temp_particle.Grh, _
                temp_particle.X + screen_x, _
                temp_particle.Y + screen_y, _
                rgb_list(), 1, False, True _
                , True
    End If
End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
    If UserMinHP = 0 Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
        
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
                            
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
    
            Particle_Group_Map_Pos_Set = True
        End If
    End If
End Function

Public Function Particle_Group_Move(ByVal particle_group_index As Long, ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim nX As Integer
    Dim nY As Integer
    
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Particle_Group_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
    
        map_x = particle_group_list(particle_group_index).map_x
        map_y = particle_group_list(particle_group_index).map_y
        
        nX = map_x
        nY = map_y
        
        Convert_Heading_to_Direction Heading, nX, nY
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
            'Move it
            particle_group_list(particle_group_index).map_x = nX
            particle_group_list(particle_group_index).map_y = nY
            
            Particle_Group_Move = True
        End If
    End If
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    Exit For
                End If
            Next i
        End If
    End If
    
    particle_group_list(particle_group_index) = temp
            
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    'color
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey

    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Sub

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        
    End If

End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************

Dim i As Integer

For i = 1 To charlist(char_index).particle_count
    If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).particle_group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) Then
        For i = 1 To charlist(char_index).particle_count
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
    End If
    
End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long

    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.FrameCounter = 1
            'End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.grhindex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
    
    If GrhData(Grh.grhindex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * base_tile_size) + base_tile_size
    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        Width, Height, alpha_blend ', grh.angle
End Sub

Private Function Char_Check(ByVal char_index As Integer) As Boolean
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
    
End Function

Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False, Optional ByVal KillAnim As Boolean = 0)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
On Error Resume Next
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
   
    If Grh.grhindex = 0 Then Exit Sub
       
    'Animation
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1
            If Grh.Loops <> -1 Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
 
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    'If Not Grh_Check(Grh.grhindex) Then Exit Sub
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
       
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.grhindex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
   
    If v_centered Then
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
        End If
    End If
   
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend
 
End Sub

Public Sub GrhRenderToHdc(ByVal grh_index As Long, desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)

    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long

    If grh_index <= 0 Then Exit Sub

    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If


        file_path = DirGraficos & GrhData(grh_index).FileNum & ".bmp"
        
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
            
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
        If transparent = False Then
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        Else
            MaskDC = CreateCompatibleDC(desthDC)
            
            PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
            Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
            
            'Render tranparently
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
            
            Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
            DeleteDC MaskDC
        End If
        
        Call DeleteObject(SelectObject(hdcsrc, PrevObj))
        DeleteDC hdcsrc

    Exit Sub
End Sub

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim TransColor As Long
    Dim ColorKey As String
    
    ColorKey = "0"
    TransColor = &H0

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For X = src_x To src_width + src_x
            If GetPixel(hdcsrc, X, Y) = TransColor Then
                SetPixel MaskDC, X, Y, vbWhite
                SetPixel hdcsrc, X, Y, vbBlack
            Else
                SetPixel MaskDC, X, Y, vbBlack
            End If
        Next X
    Next Y
End Sub

'---------------------------------------FIN PARTICULAS-------------------------------



