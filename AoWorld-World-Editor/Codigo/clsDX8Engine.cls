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
'MOTOR DX8 POR LADDER
Option Explicit

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim font_count      As Long
Dim font_last       As Long

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
Private Type CUSTOMVERTEX
    X As Single         'x in screen space.
    Y As Single         'y in screen space.
    Z As Single         'normalized z.
    rhw As Single       'normalized z rhw.
    color As Long       'vertex color.
End Type
Dim texture      As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS                     As Integer
Private FramesPerSecCounter    As Integer
Private timerElapsedTime       As Single
Private timerTicksPerFrame     As Double

Public engineBaseSpeed         As Single
Private lFrameTimer            As Long
Private lFrameLimiter          As Long
Private ScrollPixelsPerFrameX  As Byte
Private ScrollPixelsPerFrameY  As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop            As Integer
Private MainViewLeft           As Integer
Private MainDestRect           As RECT
Private MainViewRect           As RECT
Private BackBufferRect         As RECT

Private MainViewWidth          As Integer
Private MainViewHeight         As Integer

Private MouseTileX             As Integer
Private MouseTileY             As Integer
Private RLluvia(7)             As RECT  'RECT de la lluvia
Private iFrameIndex            As Byte  'Frame actual de la LL
Private llTick                 As Long  'Contador
Private LTLluvia(4)            As Integer
Private WindowTileWidth        As Integer
Private WindowTileHeight       As Integer

Private HalfWindowTileWidth    As Integer
Private HalfWindowTileHeight   As Integer
Private Const GrhFogata        As Integer = 1521

Private Type Light

    active As Boolean 'Do we ignore this light?
    id As Long
    map_X As Integer 'Coordinates
    map_Y As Integer
    color As Long 'Start colour
    range As Byte

End Type

'Light list
Dim light_list() As Light
Dim light_count  As Long
Dim light_last   As Long

Private Type Particle

    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
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

'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group

    active As Boolean
    id As Long
    map_X As Integer
    map_Y As Integer
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
    
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
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
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type

'Particle system
Dim particle_group_list() As particle_group
Dim particle_group_count  As Long
Dim particle_group_last   As Long

'Loaded Particle groups list
Dim StreamData()          As particle_group
Dim TotalStreams          As Long

Dim valoreBlur            As Long
Dim dimeTex               As Long
Dim tex                   As Direct3DTexture8
Dim D3DbackBuffer         As Direct3DSurface8
Dim zTarget               As Direct3DSurface8
Dim stencil               As Direct3DSurface8
Dim superTex              As Direct3DSurface8
Dim blur                  As Boolean
Public blur_factor        As Byte

Dim bump_map_texture      As Direct3DTexture8
Dim bump_map_texture_ex   As Direct3DTexture8
Dim bump_map_supported    As Boolean
Dim bump_map_powa         As Boolean
'Sets a Grh animation to loop indefinitely.

#Const HARDCODED = False 'True ' == MÁS FPS ^^

Public Function change_day_effect(ByVal o_r As Long, ByVal og As Long, ByVal ob As Long, ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long) As Long
    
    On Error GoTo change_day_effect_Err
    
    Dim r As Long
    Dim g As Long
    Dim b As Long
    r = o_r
    g = og
    b = ob
    
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
        b = b - 1
    ElseIf og < r Then
        b = b + 1

    End If

    If r > 255 Then r = 255
    If g > 255 Then g = 255
    If b > 255 Then b = 255
    day_r_old = r
    day_g_old = g
    day_b_old = b

    If r < 0 Or g < 0 Or b < 0 Then
        r = 1
        g = 1
        b = 1
    End If

    change_day_effect = ARGB(r, g, b, 255)

    
    Exit Function

change_day_effect_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.change_day_effect", Erl)
    Resume Next
    
End Function

Public Sub setup_ambient()
    
    On Error GoTo setup_ambient_Err
    

    'Noche 87, 61, 43
    luz_dia(0).r = 87
    luz_dia(0).g = 61
    luz_dia(0).b = 43
    luz_dia(1).r = 143
    luz_dia(1).g = 137
    luz_dia(1).b = 135
    luz_dia(2).r = 143
    luz_dia(2).g = 137
    luz_dia(2).b = 135
    luz_dia(3).r = 143
    luz_dia(3).g = 137
    luz_dia(3).b = 135
    '4 am 124,117,91
    luz_dia(4).r = 143
    luz_dia(4).g = 137
    luz_dia(4).b = 135
    '5,6 am 143,137,135
    luz_dia(5).r = 143
    luz_dia(5).g = 137
    luz_dia(5).b = 135
    luz_dia(6).r = 143
    luz_dia(6).g = 137
    luz_dia(6).b = 135
    '7 am 212,205,207
    luz_dia(7).r = 212
    luz_dia(7).g = 205
    luz_dia(7).b = 207
    luz_dia(8).r = 212
    luz_dia(8).g = 205
    luz_dia(8).b = 207
    luz_dia(9).r = 212
    luz_dia(9).g = 205
    luz_dia(9).b = 207
    luz_dia(10).r = 212
    luz_dia(10).g = 205
    luz_dia(10).b = 207
    luz_dia(11).r = 212
    luz_dia(11).g = 205
    luz_dia(11).b = 207
    luz_dia(12).r = 212
    luz_dia(12).g = 205
    luz_dia(12).b = 207
    'Dia 255, 255, 255
    luz_dia(12).r = 255
    luz_dia(12).g = 255
    luz_dia(12).b = 255
    luz_dia(13).r = 255
    luz_dia(13).g = 255
    luz_dia(13).b = 255
    'Medio Dia 255, 200, 255
    luz_dia(14).r = 255
    luz_dia(14).g = 250
    luz_dia(14).b = 255
    luz_dia(15).r = 255
    luz_dia(15).g = 240
    luz_dia(15).b = 255
    luz_dia(16).r = 255
    luz_dia(16).g = 230
    luz_dia(16).b = 255
    '17/18 0, 100, 255
    luz_dia(17).r = 230
    luz_dia(17).g = 230
    luz_dia(17).b = 255
    '18/19 0, 100, 255
    luz_dia(18).r = 230
    luz_dia(18).g = 230
    luz_dia(18).b = 255
    '19/20 156, 142, 83
    luz_dia(19).r = 156
    luz_dia(19).g = 142
    luz_dia(19).b = 83
    luz_dia(20).r = 87
    luz_dia(20).g = 61
    luz_dia(20).b = 43
    luz_dia(21).r = 87
    luz_dia(21).g = 61
    luz_dia(21).b = 43
    luz_dia(22).r = 87
    luz_dia(22).g = 61
    luz_dia(22).b = 43
    luz_dia(23).r = 87
    luz_dia(23).g = 61
    luz_dia(23).b = 43
    luz_dia(24).r = 87
    luz_dia(24).g = 61
    luz_dia(24).b = 43
            
    
    Exit Sub

setup_ambient_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.setup_ambient", Erl)
    Resume Next
    
End Sub

Public Function GetElapsedTime() As Single
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    
    On Error GoTo GetElapsedTime_Err
    
    Dim start_time    As Currency
    Static end_time   As Currency
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

    
    Exit Function

GetElapsedTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.GetElapsedTime", Erl)
    Resume Next
    
End Function

Public Sub Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal Format As Long, Optional ByVal shadow As Boolean = False)
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    
    On Error GoTo Text_Render_Err
    
    Dim TextRect   As RECT
    Dim ShadowRect As RECT
    
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.Bottom = Top + Height
    TextRect.Right = Left + Width
    
    If shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 2
        ShadowRect.Bottom = (Top + Height) - 1
        ShadowRect.Right = (Left + Width) - 2
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, Format

    End If
    
    D3DX.DrawText font, color, Text, TextRect, Format

    
    Exit Sub

Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Text_Render", Erl)
    Resume Next
    
End Sub

Public Sub Text_Render_ext(ByVal Text As String, ByVal Top As Long, ByVal Left As Long, ByVal Width As Long, ByVal Height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False)
    
    On Error GoTo Text_Render_ext_Err
    

    If center = True Then
        Call Text_Render(font_list(1), Text, Top, Left, Width, Height, color, fa_center, shadow)
    Else
        Call Text_Render(font_list(1), Text, Top, Left, Width, Height, color, DT_TOP Or DT_LEFT, shadow)

    End If

    
    Exit Sub

Text_Render_ext_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Text_Render_ext", Erl)
    Resume Next
    
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal Style As String, ByVal bold As Boolean, ByVal italic As Boolean, ByVal Size As Long)
    
    On Error GoTo Font_Make_Err
    

    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)

    End If

    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt       As New StdFont
    fnt.name = Style
    fnt.Size = Size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)

    
    Exit Sub

Font_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Font_Make", Erl)
    Resume Next
    
End Sub

Public Function Font_Create(ByVal Style As String, ByVal Size As Long, ByVal bold As Boolean, ByVal italic As Boolean) As Long

    On Error GoTo ErrorHandler:

    Font_Create = Font_Next_Open
    Font_Make Font_Create, Style, bold, italic, Size
ErrorHandler:
    Font_Create = 0

End Function

Private Function Font_Next_Open() As Long
    
    On Error GoTo Font_Next_Open_Err
    
    Font_Next_Open = font_last + 1

    
    Exit Function

Font_Next_Open_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Font_Next_Open", Erl)
    Resume Next
    
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
    
    On Error GoTo Font_Check_Err
    

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True

    End If

    
    Exit Function

Font_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Font_Check", Erl)
    Resume Next
    
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    
    On Error GoTo MakeVector_Err
    
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z

    
    Exit Function

MakeVector_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.MakeVector", Erl)
    Resume Next
    
End Function

Public Sub Engine_Init()
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    'On Error GoTo ErrHandler:
    
    On Error GoTo Engine_Init_Err
    

    Dim DispMode    As D3DDISPLAYMODE
    Dim DispModeBK  As D3DDISPLAYMODE
    Dim D3DWindow   As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = 3200
        .BackBufferHeight = 3200
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = FrmMain.renderer.hWnd

    End With

    DispMode.Format = D3DFMT_X8R8G8B8

    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.Format = DispModeBK.Format

        End If

    Else
        bump_map_supported = False
        DispMode.Format = DispModeBK.Format

    End If

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FrmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    HalfWindowTileHeight = (FrmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (FrmMain.renderer.ScaleWidth / 32) \ 2
    
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
    
    Dim viewport As D3DVIEWPORT8
    viewport.Width = FrmMain.renderer.ScaleWidth
    viewport.Height = FrmMain.renderer.ScaleHeight
    viewport.MinZ = 1
    viewport.MaxZ = -1
    D3DDevice.SetViewport viewport
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.017
    
    ReDim MapData(1 To 1, 1 To 1) As MapBlock
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    UserPos.X = 50
    UserPos.Y = 50
    
    MinXBorder = 1 + (FrmMain.renderer.ScaleWidth / 64)
    MaxXBorder = MapSize.Width - (FrmMain.renderer.ScaleWidth / 64)
    MinYBorder = 1 + (FrmMain.renderer.ScaleHeight / 64)
    MaxYBorder = MapSize.Height - (FrmMain.renderer.ScaleHeight / 64)
    
    Font_Create "Tahoma", 8, True, 0
    Font_Create "Verdana", 8, False, 0
    
    COLOR_WHITE(0) = -1
    COLOR_WHITE(1) = -1
    COLOR_WHITE(2) = -1
    COLOR_WHITE(3) = -1
    
    bRunning = True
    Exit Sub
ErrHandler:
    Debug.Print "Error Number Returned: " & Err.Number
    bRunning = False

    
    Exit Sub

Engine_Init_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Engine_Init", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Deinit()
    
    On Error GoTo Engine_Deinit_Err
    
    Erase MapData
    Erase CharList

    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing
    End

    
    Exit Sub

Engine_Deinit_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Engine_Deinit", Erl)
    Resume Next
    
End Sub

Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    
    On Error GoTo CreateTLVertex_Err
    
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv

    
    Exit Function

CreateTLVertex_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.CreateTLVertex", Erl)
    Resume Next
    
End Function

Public Sub Engine_ActFPS()
    
    On Error GoTo Engine_ActFPS_Err
    

    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount

    End If

    
    Exit Sub

Engine_ActFPS_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Engine_ActFPS", Erl)
    Resume Next
    
End Sub

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
    
    On Error GoTo InMapBounds_Err
    

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If X < 1 Or X > MapSize.Width Or Y < 1 Or Y > MapSize.Height Then
        Exit Function

    End If
    
    InMapBounds = True

    
    Exit Function

InMapBounds_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.InMapBounds", Erl)
    Resume Next
    
End Function

Public Sub Draw_GrhIndex(ByVal grh_index As Long, ByVal X As Integer, ByVal Y As Integer)
    
    On Error GoTo Draw_GrhIndex_Err
    

    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, X, Y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY

    
    Exit Sub

Draw_GrhIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Draw_GrhIndex", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhIndexColor(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    
    On Error GoTo Draw_GrhIndexColor_Err
    

    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(120, 120, 120)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, X, Y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, True

    
    Exit Sub

Draw_GrhIndexColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Draw_GrhIndexColor", Erl)
    Resume Next
    
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, ByRef rgb_list() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_X As Integer = 1, Optional ByVal map_Y As Integer = 1, Optional ByVal angle As Single)

    On Error Resume Next

    Dim CurrentGrhIndex As Long

    If Grh.grhindex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.speed)

            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1

                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0 '

                    End If '''

                End If

            End If

        End If

    End If
    
    If GrhData(Grh.grhindex).NumFrames = 0 Then
        Call AddtoRichTextBox(FrmMain.RichTextBox1, "Hay un error en el grh " & Grh.grhindex & " (" & map_X & ", " & map_Y & ")", 255, 100, 100, False, False, False)
        Exit Sub

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    If CurrentGrhIndex = 0 Then
        Dim IX As Integer
        Dim IY As Integer
    
        For IY = 1 To 100
            For IX = 1 To 100

                If MapData(IX, IY).Graphic(1).grhindex = Grh.grhindex Then
                    '  Debug.Print "El tile que falla es: " & ix & "-" & iy
                    MapData(IX, IY).Graphic(1).grhindex = 0

                End If

                If MapData(IX, IY).Graphic(2).grhindex = Grh.grhindex Then
                    '   Debug.Print "El tile que falla es: " & ix & "-" & iy
                    MapData(IX, IY).Graphic(2).grhindex = 0

                End If

                If MapData(IX, IY).Graphic(3).grhindex = Grh.grhindex Then
                    '    Debug.Print "El tile que falla es: " & ix & "-" & iy
                    MapData(IX, IY).Graphic(3).grhindex = 0

                End If
        
                If MapData(IX, IY).Graphic(4).grhindex = Grh.grhindex Then
                    '    Debug.Print "El tile que falla es: " & ix & "-" & iy
                    MapData(IX, IY).Graphic(4).grhindex = 0

                End If

            Next IX
        Next IY

    End If

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

        End If

    End If

    Device_Box_Textured_Render CurrentGrhIndex, X, Y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, rgb_list, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle
    'exits:

End Sub

Public Function Map_Base_Light_Get() As Long
    '**************************************************************
    'Author: Aaron Perkins - Modified by Augusto José Rando
    'Last Modify Date: 6/12/2005
    '
    '**************************************************************
    
    On Error GoTo Map_Base_Light_Get_Err
    
    Map_Base_Light_Get = map_base_light

    
    Exit Function

Map_Base_Light_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Map_Base_Light_Get", Erl)
    Resume Next
    
End Function

Public Function Map_Base_Light_Set(ByVal Base_light As Long)
    
    On Error GoTo Map_Base_Light_Set_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    If map_base_light <> Base_light Then
        If Base_light = 0 Then
            map_base_light = &HFFFFFFFF
        Else
            map_base_light = Base_light
        End If
        
    End If

    
    Exit Function

Map_Base_Light_Set_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Map_Base_Light_Set", Erl)
    Resume Next
    
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal LAYER As Byte, Optional ByVal light_base_color As Long = -1, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    
    On Error GoTo Map_Fill_Err
    
    Dim X As Integer
    Dim Y As Integer
    
    'Base light color
    If light_base_color <> -1 Then
        If Not Map_Base_Light_Set(light_base_color) Then Exit Function

    End If
        
    Map_Fill = True

    
    Exit Function

Map_Fill_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Map_Fill", Erl)
    Resume Next
    
End Function

Public Sub Render()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    On Error Resume Next

    D3DDevice.BeginScene

    If FrmMain.PreviewGrh.Visible = True Then

        ' Call modPaneles.VistaPreviaDeSup
    End If
    
    Static RE            As RECT
    
    With RE
        .Left = 0
        .Top = 0
        .Bottom = FrmMain.renderer.ScaleHeight
        .Right = FrmMain.renderer.ScaleWidth

    End With

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0

    ShowNextFrame

    If bRefreshRadar = True Then
        Call RefreshAllChars
        bRefreshRadar = False

    End If

    Text_Render font_list(1), "FPS: " & FPS, 10, 10, 70, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    Text_Render font_list(1), SeleccionAncho & "," & SeleccionAlto, 30, 10, 70, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    
    Text_Render font_list(1), SeleccionFX - SeleccionIX + 1 & "," & SeleccionFY - SeleccionIY + 1, 40, 10, 70, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    
    Text_Render font_list(1), "Ultimo Click: X: " & UltimoClickX & " Y: " & UltimoClickY, 20, 10, 150, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    If SelectedZona > 0 Then
        Text_Render font_list(1), "Zona seleccionada: " & SelectedZona & " - " & Zona(SelectedZona).Zona_name, 60, 10, 300, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    End If
    D3DDevice.EndScene

    D3DDevice.Present RE, ByVal 0, 0, ByVal 0

    lFrameLimiter = GetTickCount
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

End Sub

Sub ShowNextFrame()
    
    On Error GoTo ShowNextFrame_Err
    
    'Set SurfaceDB.Used = New Collection
    
    Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, 0, 0)

    'SurfaceDB.UnloadUnused
    
    Exit Sub

ShowNextFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.ShowNextFrame", Erl)
    Resume Next
    
End Sub

Sub RenderScreen(ByVal TileX As Integer, _
                 ByVal TileY As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus) ver ReyarB
    'Renders everything to the viewport
    '**************************************************************
    
    On Error GoTo RenderScreen_Err
    
    Dim Y                As Integer     'Keeps track of where on map we are
    Dim X                As Integer     'Keeps track of where on map we are
    Dim bCapa            As Byte
    Dim Sobre            As Long
    Dim screenminY       As Integer  'Start Y pos on current screen
    Dim screenmaxY       As Integer  'End Y pos on current screen
    Dim screenminX       As Integer  'Start X pos on current screen
    Dim screenmaxX       As Integer  'End X pos on current screen
    Dim minY             As Integer  'Start Y pos on current map
    Dim maxY             As Integer  'End Y pos on current map
    Dim minX             As Integer  'Start X pos on current map
    Dim maxX             As Integer  'End X pos on current map
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim CurrentGrhIndex  As Long
    Dim offX             As Integer
    Dim offY             As Integer
    Dim Grh              As Grh
    Dim TempChar         As char
    Dim i                As Integer
    
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    '******************************************************
    'Ultima modificacion 08/05/2020 por ReyarB
    '******************************************************
    If minY < 1 Then
        minYOffset = 1 - minY
        minY = 1
    End If
    
    If maxY > MapSize.Height Then maxY = MapSize.Height
    
    If minX < 1 Then
        minXOffset = 1 - minX
        minX = 1
    End If
    
    If maxX > MapSize.Width Then maxX = MapSize.Width
    
    'If we can, we render around the view area to make it smoother ver ReyarB
    If screenminY > 1 Then
        screenminY = screenminY - 1
    Else
        screenminY = screenminY - 1
        'ScreenY = 1
    End If
    
    If screenmaxY < MapSize.Height Then screenmaxY = screenmaxY + 1
    
    If screenminX > 1 Then
        screenminX = screenminX - 1
    Else
        screenminX = screenminX - 1
        'ScreenX = 1
    End If
    
    If screenmaxX < MapSize.Width Then screenmaxX = screenmaxX + 1

    If Val(FrmMain.cCapas.Text) >= 1 And (FrmMain.cCapas.Text) <= 4 Then
        bCapa = Val(FrmMain.cCapas.Text)
    Else
        bCapa = 1

    End If

    GenerarVista 'Loopzer

    'If screenmaxX > 100 Then screenmaxX = 100

    If screenminX < 1 Then
        ScreenX = 1 - screenminX
        screenminX = 1
    End If

    If screenminY < 1 Then
        ScreenY = 1 - screenminY
        screenminY = 1
    End If
    'Draw floor layer
    
    For Y = screenminY To screenmaxY
    
        For X = screenminX To screenmaxX

            If X > MapSize.Width Or X < 1 Or Y > MapSize.Height Or Y < 1 Then GoTo continue1



            'Layer 1 **********************************
            
            If VerCapa1 Then
            
                Call Draw_Grh(MapData(X, Y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(X, Y).light_value, False, X, Y)

            End If
            
            'If MapData(X, y).Graphic(5).grhindex <> 1 Then
                            
            '    Call Draw_Grh(MapData(X, y).Graphic(5), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(X, y).light_value, False, X, y)
                        
            'End If
                        
            'Layer 2  **********************************

            '******************************************
            

continue1:
            ScreenX = ScreenX + 1

            'If X > 100 Then Exit For
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1

        'If y > 99 Then Exit For

    Next Y
    
    ScreenY = minYOffset - TileBufferSize
    
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
        
            If X > MapSize.Width Or X < 1 Or Y > MapSize.Height Or Y < 1 Then GoTo continue2
        
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            With MapData(X, Y)

                '******************************************
                If MapData(X, Y).Graphic(2).grhindex <> 0 And VerCapa2 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, False, X, Y)

                End If
                
                'If MapData(X, y).Graphic(6).grhindex <> 1 Then
                            
                '    Call Draw_Grh(MapData(X, y).Graphic(6), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(X, y).light_value, False, X, y)
                        
                'End If

                'Object Layer **********************************
                If .OBJInfo.ObjIndex <> 0 And VerObjetos Then
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)

                End If

                '***********************************************
                
                'Char layer **********************************

                If .CharIndex <> 0 And VerNpcs Then
                
                    Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
          
                End If
                
                'If MapData(X, Y).Graphic(7).grhindex <> 1 Then
                            
                'Call Draw_Grh(MapData(X, Y).Graphic(7), _
                 (ScreenX - 1) * 32 + PixelOffsetX, _
                 (ScreenY - 1) * 32 + PixelOffsetY, _
                 0, 1, MapData(X, Y).light_value, False, X, Y)
                        
                ' End If
      
                '************************************************
            End With
continue2:
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - 5
    
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
        
            If X > MapSize.Width Or X < 1 Or Y > MapSize.Height Or Y < 1 Then GoTo continue3
        
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
    
            'Layer 3 *****************************************
            If MapData(X, Y).Graphic(3).grhindex <> 0 And VerCapa3 Then
                Call Draw_Grh(MapData(X, Y).Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)

            End If
continue3:
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - 5
    
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
        
            If X > MapSize.Width Or X < 1 Or Y > MapSize.Height Or Y < 1 Then GoTo continue4
            
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            If MapData(X, Y).particle_group > 0 And VerParticulas Then
                Call Particle_Group_Render(MapData(X, Y).particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

            End If
continue4:
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - 5
    
    Dim TmpRGB(3) As Long
    TmpRGB(0) = D3DColorARGB(255, 255, 255, 255)
    TmpRGB(1) = TmpRGB(0)
    TmpRGB(2) = TmpRGB(0)
    TmpRGB(3) = TmpRGB(0)
        
    Dim HeadingIt As Byte
                    
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
        
            If X > MapSize.Width Or X < 1 Or Y > MapSize.Height Or Y < 1 Then GoTo continue5
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
                    
            'Layer 4 *****************************************
            If MapData(X, Y).Graphic(4).grhindex <> 0 And (FrmMain.mnuVerCapa4.Checked = True) Then
                Call Draw_Grh(MapData(X, Y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value, AlphaTecho, X, Y)

            End If
 
            '************************************************
            
            If VerBlockeados And MapData(X, Y).Blocked > 0 Then
                HeadingIt = 1

                For i = 1 To 4 ' Headings

                    If (MapData(X, Y).Blocked And HeadingIt) <> 0 Then
                        Grh.grhindex = 1150 + i
                        Grh.FrameCounter = 1
                        Grh.Started = 0
                        Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TmpRGB(), False, X, Y)

                    End If

                    HeadingIt = HeadingIt * 2
                Next

            End If
            
            If MapData(X, Y).TileExit.Map <> 0 And VerTranslados Then
                Grh.grhindex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TmpRGB(), True, X, Y)

            End If
            
            If VerTriggers Then

                If MapData(X, Y).Trigger <> 0 Then

                    Dim cadena   As String

                    Dim caracter As String

                    Dim POSX     As Integer

                    Dim jj       As Byte

                    POSX = 0
                    cadena = str(MapData(X, Y).Trigger)
                    
                    For jj = 1 To Len(cadena)
                        caracter = mid(cadena, jj, 1)

                        If IsNumeric(caracter) Then
                            Grh.grhindex = caracter + 21452
                            Grh.FrameCounter = 1
                            Grh.Started = 0
                            Call Draw_Grh(Grh, PixelOffsetXTemp - 6 + POSX, PixelOffsetYTemp - 12, 1, 1, TmpRGB(), False, X, Y)
                            POSX = POSX + 8

                        End If

                    Next jj

                End If
                


            End If
          
            If Seleccionando And SeleccionFX > 0 Then
                Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
                If SeleccionIX > SeleccionFX Then
                    X1 = SeleccionFX
                    X2 = SeleccionIX
                Else
                    X1 = SeleccionIX
                    X2 = SeleccionFX
                End If
                If SeleccionIY > SeleccionFY Then
                    Y1 = SeleccionFY
                    Y2 = SeleccionIY
                Else
                    Y1 = SeleccionIY
                    Y2 = SeleccionFY
                End If
                If X >= X1 And Y >= Y1 Then
                    If X <= X2 And Y <= Y2 Then
                        Engine_Draw_Box PixelOffsetXTemp, PixelOffsetYTemp, 30, 30, D3DColorARGB(75, 0, 200, 100)
                    End If
                End If

            End If

continue5:
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
        
    
    If FrmMain.cSeleccionarSuperficie.value = True Then
        Dim mGrh As Grh
        Dim aux As Long
        Dim dy  As Integer
        Dim dX  As Integer
        If frmConfigSup.DespMosaic.value = vbChecked Then
            dy = Val(frmConfigSup.DMLargo.Text)
            dX = Val(frmConfigSup.DMAncho.Text)
        Else
            dy = 0
            dX = 0

        End If
        
        If frmConfigSup.mAncho.Text = "0" Then frmConfigSup.mAncho.Text = "1"
        If frmConfigSup.mLargo.Text = "0" Then frmConfigSup.mLargo.Text = "1"

        If FrmMain.mnuAutoCompletarSuperficies.Checked = False Then
            aux = Val(FrmMain.cGrh.Text) + (((SobreY + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((SobreX + dX) Mod frmConfigSup.mAncho.Text)
            mGrh.grhindex = aux
            InitGrh mGrh, aux
            Call Draw_Grh(mGrh, (SobreX - UserPos.X + 18) * 32 + PixelOffsetX, (SobreY - UserPos.Y + 13) * 32 + PixelOffsetY, 1, 1, MapData(1, 1).light_value, 0, X, Y)
            
        Else
            For X = 1 To frmConfigSup.mAncho.Text
                For Y = 1 To frmConfigSup.mLargo.Text
                     aux = Val(FrmMain.cGrh.Text) + (Y - 1) * frmConfigSup.mAncho.Text + X - 1
                     mGrh.grhindex = aux
                     InitGrh mGrh, aux
                     Call Draw_Grh(mGrh, (SobreX - UserPos.X + 18 + X - 1) * 32 + PixelOffsetX, (SobreY - UserPos.Y + 13 + Y - 1) * 32 + PixelOffsetY, 1, 1, MapData(1, 1).light_value, 0, X, Y)
            
                Next Y
            Next X

        End If
        'Call Draw_Grh(mGrh, (SobreX - UserPos.X + 18) * 32 + PixelOffsetX, (SobreY - UserPos.Y + 13) * 32 + PixelOffsetY, 1, 1, MapData(1, 1).light_value, 0, X, Y)
            
    End If
    
    'Previsualizar seleccion
    If ModoPegar Then
        For Y = 0 To SeleccionAlto - 1
            For X = 0 To SeleccionAncho - 1
                If Y + SobreY > MapSize.Height Then Exit For
                If X + SobreX > MapSize.Width Then Exit For
                Call Draw_Grh(SeleccionMap(X, Y).Graphic(1), (X + SobreX - UserPos.X + 18) * 32 + PixelOffsetX, (Y + SobreY - UserPos.Y + 13) * 32 + PixelOffsetY, 1, 1, SeleccionMap(X, Y).light_value, 230, X, Y)
            Next X
        Next Y
        
        For Y = 0 To SeleccionAlto - 1
            For X = 0 To SeleccionAncho - 1
                If Y + SobreY > MapSize.Height Then Exit For
                If X + SobreX > MapSize.Width Then Exit For
                Call Draw_Grh(SeleccionMap(X, Y).Graphic(2), (X + SobreX - UserPos.X + 18) * 32 + PixelOffsetX, (Y + SobreY - UserPos.Y + 13) * 32 + PixelOffsetY, 1, 1, SeleccionMap(X, Y).light_value, 230, X, Y)
            Next X
        Next Y
        
        For Y = 0 To SeleccionAlto - 1
            For X = 0 To SeleccionAncho - 1
                If Y + SobreY > MapSize.Height Then Exit For
                If X + SobreX > MapSize.Width Then Exit For
                Call Draw_Grh(SeleccionMap(X, Y).Graphic(3), (X + SobreX - UserPos.X + 18) * 32 + PixelOffsetX, (Y + SobreY - UserPos.Y + 13) * 32 + PixelOffsetY, 1, 1, SeleccionMap(X, Y).light_value, 230, X, Y)
            Next X
        Next Y
        
        For Y = 0 To SeleccionAlto - 1
            For X = 0 To SeleccionAncho - 1
                If Y + SobreY > MapSize.Height Then Exit For
                If X + SobreX > MapSize.Width Then Exit For
                Call Draw_Grh(SeleccionMap(X, Y).Graphic(4), (X + SobreX - UserPos.X + 18) * 32 + PixelOffsetX, (Y + SobreY - UserPos.Y + 13) * 32 + PixelOffsetY, 1, 1, SeleccionMap(X, Y).light_value, 230, X, Y)
            Next X
        Next Y
    End If
    


    
    
    

    ScreenY = minYOffset - 5
              
    TmpRGB(0) = D3DColorARGB(254, 128, 128, 128)
    TmpRGB(1) = D3DColorARGB(255, 128, 128, 128)
    TmpRGB(2) = D3DColorARGB(255, 128, 128, 128)
    TmpRGB(3) = D3DColorARGB(255, 128, 128, 128)
   
    If VerGrilla Then
        Dim A As Byte
        Dim b As Byte
    
        For A = 0 To 35
            For b = 0 To 26

                Call Draw_GrhIndex(2, 32 * A, 32 * b)
                '
                'Engine_Draw_Box 32 * A, 32 * b, 30, 30, D3DColorARGB(10, 255, 255, 255)
            Next b
        Next A
        
    End If
    
    If VerMarco Then
        Call Draw_GrhIndex(2476, 212, 128)
    End If
    Dim color As Long
 
    color = ColorBlanco
    Call DrawBox(CLng((1 - UserPos.X + 18)) * 32, CLng((1 - UserPos.Y + 13)) * 32, CLng((MapSize.Width - UserPos.X + 19)) * 32, CLng((MapSize.Height - UserPos.Y + 14)) * 32, color)

    'Estoy creando una zona o spawn
    If RectanguloModo > 0 And RectanguloX > 0 Then
       Call DrawBox(CLng((RectanguloX - UserPos.X + 18)) * 32, CLng((RectanguloY - UserPos.Y + 13)) * 32, CLng((RectanguloX2 - UserPos.X + 19)) * 32, CLng((RectanguloY2 - UserPos.Y + 14)) * 32, color)
    End If
    
    If FrmMain.mnuZonas.Checked Then
        For i = 1 To NumZonas
            color = ColorAmarillo
            If i = SelectedZona Then
                color = ColorNaranja
            End If
            With Zona(i)
                If .Map = UserMap And .X > 0 And .Deleted = 0 Then
                    Call DrawBox(CLng((.X - UserPos.X + 18)) * 32, CLng((.Y - UserPos.Y + 13)) * 32, CLng((.X2 - UserPos.X + 19)) * 32, CLng((.Y2 - UserPos.Y + 14)) * 32, color)
                    
                    If .X - UserPos.X > -20 And .Y - UserPos.Y > -15 Then
                        Call Text_Render(font_list(1), i & " - " & .Zona_name, CLng((.Y - UserPos.Y + 13)) * 32 + 4, CLng((.X - UserPos.X + 18)) * 32 + 6, 250, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True)
                    End If
                End If
            End With
        Next i
    End If
    If FrmMain.mnuNpcSpawn.Checked Then
        For i = 1 To NumSpawns
            color = ColorVerde
            If i = SelectedSpawn Then
                color = ColorCian
            End If
            With NpcSpawn(i)
                If .Map = UserMap And .X > 0 And .Deleted = 0 Then
                    Call DrawBox(CLng((.X - UserPos.X + 18)) * 32, CLng((.Y - UserPos.Y + 13)) * 32, CLng((.X2 - UserPos.X + 19)) * 32, CLng((.Y2 - UserPos.Y + 14)) * 32, color)
                    
                    If .X - UserPos.X > -20 And .Y - UserPos.Y > -15 Then
                        Call Text_Render(font_list(1), "Spawn: " & i, CLng((.Y - UserPos.Y + 13)) * 32 + 4 + 30, CLng((.X - UserPos.X + 18)) * 32 + 6, 250, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True)
                        Y = 30
                        For X = 1 To .CantNpcs
                            Y = Y + 15
                            Call Text_Render(font_list(1), "   " & .NPCs(X).Cantidad & " x " & NpcData(.NPCs(X).NpcIndex).name & " (" & .NPCs(X).NpcIndex & ")", CLng((.Y - UserPos.Y + 13)) * 32 + 4 + Y, CLng((.X - UserPos.X + 18)) * 32 + 6, 250, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True)
                            
                        Next X
                    End If
                End If
            End With
        Next i
    End If
    
    
    Exit Sub

RenderScreen_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.RenderScreen", Erl)
    Resume Next
    
End Sub
Private Sub DrawBox(ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal color As Long)
    Dim Lines(7) As TLVERTEX
    
    If X < 0 Then X = -1
    If Y < 0 Then Y = -1
    
    If X > 2000 Or Y > 2000 Then Exit Sub
    If X2 > 2000 Then X2 = 2000
    If Y2 > 2000 Then Y2 = 2000

    
    Lines(0) = Geometry_Create_TLVertex(X, Y, 0, 1, color, 0, 0, 0)
    Lines(1) = Geometry_Create_TLVertex(X2, Y, 0, 1, color, 0, 0, 0)
    Lines(2) = Geometry_Create_TLVertex(X2, Y, 0, 1, color, 0, 0, 0)
    Lines(3) = Geometry_Create_TLVertex(X2, Y2, 0, 1, color, 0, 0, 0)
    Lines(4) = Geometry_Create_TLVertex(X2, Y2, 0, 1, color, 0, 0, 0)
    Lines(5) = Geometry_Create_TLVertex(X, Y2, 0, 1, color, 0, 0, 0)
    Lines(6) = Geometry_Create_TLVertex(X, Y2, 0, 1, color, 0, 0, 0)
    Lines(7) = Geometry_Create_TLVertex(X, Y, 0, 1, color, 0, 0, 0)
    
    D3DDevice.SetTexture 0, Nothing
    
    D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 4, Lines(0), Len(Lines(0))

End Sub
Private Function Convert_Tile_To_View_X(ByVal X As Integer) As Integer
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 10/07/2002
    'Convert tile position into position in view area
    '**************************************************************
    'If engine_windowed Then
    
    On Error GoTo Convert_Tile_To_View_X_Err
    
    Convert_Tile_To_View_X = ((X - 1) * 32)

    ' Else
    '  Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    '  End If
    
    Exit Function

Convert_Tile_To_View_X_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Convert_Tile_To_View_X", Erl)
    Resume Next
    
End Function

Private Function Convert_Tile_To_View_Y(ByVal Y As Integer) As Integer
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 10/07/2002
    'Convert tile position into position in view area
    '**************************************************************
    ' If engine_windowed Then
    
    On Error GoTo Convert_Tile_To_View_Y_Err
    
    Convert_Tile_To_View_Y = ((Y - 1) * 32)

    'Else
    '   Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    'End If
    
    Exit Function

Convert_Tile_To_View_Y_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Convert_Tile_To_View_Y", Erl)
    Resume Next
    
End Function

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, ByVal tv As Single) As TLVERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    
    On Error GoTo Geometry_Create_TLVertex_Err
    
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv

    
    Exit Function

Geometry_Create_TLVertex_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Geometry_Create_TLVertex", Erl)
    Resume Next
    
End Function

Private Function Geometry_Create_TLVertex2(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
    'mz
    
    On Error GoTo Geometry_Create_TLVertex2_Err
    
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

    
    Exit Function

Geometry_Create_TLVertex2_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Geometry_Create_TLVertex2", Erl)
    Resume Next
    
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
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
    
    On Error GoTo Geometry_Create_Box_Err
    
    Dim x_center    As Single
    Dim y_center    As Single
    Dim radius      As Single
    Dim x_Cor       As Single
    Dim y_Cor       As Single
    Dim left_point  As Single
    Dim right_point As Single
    Dim temp        As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point

    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius

    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)

    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius

    End If
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)

    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius

    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)

    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius

    End If
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)

    End If

    
    Exit Sub

Geometry_Create_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Geometry_Create_Box", Erl)
    Resume Next
    
End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/15/2003
    'Copies the Textures allowing resizing
    'Modified by Juan Martín Sotuyo Dodero
    '**************************************************************
    
    On Error GoTo Device_Box_Textured_Render_Advance_Err
    
    Static src_rect            As RECT
    Static dest_rect           As RECT
    Static temp_verts(3)       As TLVERTEX
    Static d3dTextures         As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
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
    ' light_value(0) = &HFFFFFFFF 'blind_color
    'light_value(1) = &HFFFFFFFF 'blind_color
    'light_value(2) = &HFFFFFFFF 'blind_color
    'light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
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

    
    Exit Sub

Device_Box_Textured_Render_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Device_Box_Textured_Render_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    
    On Error GoTo Device_Box_Textured_Render_Err
    
    Static src_rect            As RECT
    Static dest_rect           As RECT
    Static temp_verts(3)       As TLVERTEX
    Static d3dTextures         As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)

    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)

    If (light_value(0) = 0) Then light_value(0) = map_base_light
    If (light_value(1) = 0) Then light_value(1) = map_base_light
    If (light_value(2) = 0) Then light_value(2) = map_base_light
    If (light_value(3) = 0) Then light_value(3) = map_base_light
        
    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
                
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
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

    
    Exit Sub

Device_Box_Textured_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Device_Box_Textured_Render", Erl)
    Resume Next
    
End Sub

Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    
    On Error GoTo Engine_MoveScreen_Err
    
    Dim X  As Integer
    Dim Y  As Integer
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

    End If

    
    Exit Sub

Engine_MoveScreen_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Engine_MoveScreen", Erl)
    Resume Next
    
End Sub

Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal X As Integer, ByVal Y As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    
    On Error GoTo Char_Render_Err
    
    Dim Moved As Boolean
    Dim Pos   As Integer
    Dim line  As String
    Dim color As Long
    Dim i     As Long
    
    With CharList(CharIndex)

        If .Heading = 0 Then Exit Sub
        
        'If done moving stop animation
        If Not Moved Then
        
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
           
        If .Body.Walk(.Heading).grhindex Then

            Dim colorz(3) As Long
            'Draw Body
                
            colorz(0) = MapData(X, Y).light_value(0)
            colorz(1) = MapData(X, Y).light_value(1)
            colorz(2) = MapData(X, Y).light_value(2)
            colorz(3) = MapData(X, Y).light_value(3)
    
             If .Body.Walk(.Heading).grhindex Then
               Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, MapData(X, Y).light_value, False, X, Y)

            If .Head.Head(.Heading).grhindex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value, False, X, Y)

            If .Casco.Head(.Heading).grhindex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value, False, X, Y)

            If .Arma.WeaponWalk(.Heading).grhindex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, MapData(X, Y).light_value, False, X, Y)

            If .Escudo.ShieldWalk(.Heading).grhindex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, MapData(X, Y).light_value, False, X, Y)
            'Draw name over head
            
        End If

    End If

End With

    
    Exit Sub

Char_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Char_Render", Erl)
    Resume Next
    
End Sub

Public Sub Start()
    
    On Error GoTo Start_Err
    
    day_r_old = luz_dia(Hour(Time)).r
    day_g_old = luz_dia(Hour(Time)).g
    day_b_old = luz_dia(Hour(Time)).b
    Base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
    
    DoEvents
    vertList(0).X = 0
    vertList(0).Y = 0
    vertList(1).X = 800
    vertList(1).Y = 0
    vertList(2).X = 0
    vertList(2).Y = 600
    vertList(3).X = 800
    vertList(3).Y = 600
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
    
    
    ColorBlanco = D3DColorRGBA(255, 255, 255, 255)
    ColorRojo = D3DColorRGBA(255, 0, 0, 255)
    ColorVerde = D3DColorRGBA(0, 255, 0, 255)
    ColorAmarillo = D3DColorRGBA(255, 255, 0, 255)
    ColorNaranja = D3DColorRGBA(255, 165, 0, 255)
    ColorCian = D3DColorRGBA(0, 255, 255, 255)
    ColorAzul = D3DColorRGBA(0, 0, 255, 255)
   
    Do While prgRun

        If FrmMain.WindowState <> vbMinimized And FrmMain.Visible = True Then
            CheckKeys
            engine.Render
        Else
            Sleep 10&

        End If

       DoEvents
        Rem Limitar FPS

        While GetTickCount - lFrameLimiter < 10

            Sleep 5
        Wend
    Loop
    engine.Engine_Deinit

    
    Exit Sub

Start_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Start", Erl)
    Resume Next
    
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************

End Sub

Public Sub Char_Move_by_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************
    
    On Error GoTo Char_Move_by_Head_Err
    
    Dim addx As Integer
    Dim addy As Integer
    Dim X    As Integer
    Dim Y    As Integer
    Dim nX   As Integer
    Dim nY   As Integer
    
    With CharList(CharIndex)
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

    
    Exit Sub

Char_Move_by_Head_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Char_Move_by_Head", Erl)
    Resume Next
    
End Sub

Public Sub Char_Move_by_Pos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim X        As Integer
    Dim Y        As Integer
    Dim addx     As Integer
    Dim addy     As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
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

    End With
    
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    
    On Error GoTo EstaPCarea_Err
    

    With CharList(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder

    End With

    
    Exit Function

EstaPCarea_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.EstaPCarea", Erl)
    Resume Next
    
End Function

Private Function Engine_FToDW(F As Single) As Long
    ' single > long
    
    On Error GoTo Engine_FToDW_Err
    
    Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, F
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW

    
    Exit Function

Engine_FToDW_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Engine_FToDW", Erl)
    Resume Next
    
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
    
    On Error GoTo VectorToRGBA_Err
    
    Dim r As Integer, g As Integer, b As Integer, A As Integer
    r = 127 * Vec.X + 128
    g = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    A = 255 * fHeight
    VectorToRGBA = D3DColorARGB(A, r, g, b)

    
    Exit Function

VectorToRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.VectorToRGBA", Erl)
    Resume Next
    
End Function

Public Function Light_Remove(ByVal light_index As Long) As Boolean
    
    On Error GoTo Light_Remove_Err
    

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

    
    Exit Function

Light_Remove_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Remove", Erl)
    Resume Next
    
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
    
    On Error GoTo Light_Color_Value_Get_Err
    

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

    
    Exit Function

Light_Color_Value_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Color_Value_Get", Erl)
    Resume Next
    
End Function

Public Function Light_Create(ByVal map_X As Integer, ByVal map_Y As Integer, Optional ByVal color_value As Long = &HFFFFFFFF, Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
    
    On Error GoTo Light_Create_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Returns the light_index if successful, else 0
    'Edited by Juan Martín Sotuyo Dodero
    '**************************************************************
    If InMapBounds(map_X, map_Y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_X, map_Y, color_value, range, id

    End If

    
    Exit Function

Light_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Create", Erl)
    Resume Next
    
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_X As Integer, ByVal map_Y As Integer) As Boolean
    
    On Error GoTo Light_Move_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Returns true if successful, else false
    '**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then

        'Make sure it's a legal move
        If InMapBounds(map_X, map_Y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_X = map_X
            light_list(light_index).map_Y = map_Y
    
            Light_Move = True
            
        End If

    End If

    
    Exit Function

Light_Move_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Move", Erl)
    Resume Next
    
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 15/05/2002
    'Returns true if successful, else false
    '**************************************************************
    
    On Error GoTo Light_Move_By_Head_Err
    
    Dim map_X As Integer
    Dim map_Y As Integer
    Dim nX    As Integer
    Dim nY    As Integer
    Dim addy  As Integer
    Dim addx  As Integer

    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function

    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_X = light_list(light_index).map_X
        map_Y = light_list(light_index).map_Y

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
        
        nX = map_X + addx
        nY = map_Y + addy
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_X = nX
            light_list(light_index).map_Y = nY
    
            Light_Move_By_Head = True
            
        End If

    End If

    
    Exit Function

Light_Move_By_Head_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Move_By_Head", Erl)
    Resume Next
    
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_X As Integer, ByVal map_Y As Integer, ByVal rgb_value As Long, ByVal range As Long, Optional ByVal id As Long)
    
    On Error GoTo Light_Make_Err
    

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
    
    light_list(light_index).map_X = map_X
    light_list(light_index).map_Y = map_Y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id

    
    Exit Sub

Light_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Make", Erl)
    Resume Next
    
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
    
    On Error GoTo Light_Check_Err
    

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

    
    Exit Function

Light_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Check", Erl)
    Resume Next
    
End Function

Public Sub Light_Render_All()
    
    On Error GoTo Light_Render_All_Err
    

    LightA.LightRenderAll

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

    
    Exit Sub

Light_Render_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Render_All", Erl)
    Resume Next
    
End Sub

Private Sub Light_Render(ByVal light_index As Long)
    'menduz
    
    On Error GoTo Light_Render_Err
    

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X     As Integer
    Dim Y     As Integer
    Dim ia    As Single
    Dim i     As Integer
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_X - light_list(light_index).range
    min_y = light_list(light_index).map_Y - light_list(light_index).range
    max_x = light_list(light_index).map_X + light_list(light_index).range
    max_y = light_list(light_index).map_Y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    MapData(light_list(light_index).map_X, light_list(light_index).map_Y).light_value(0) = color
    MapData(light_list(light_index).map_X, light_list(light_index).map_Y).light_value(1) = color
    MapData(light_list(light_index).map_X, light_list(light_index).map_Y).light_value(2) = color
    MapData(light_list(light_index).map_X, light_list(light_index).map_Y).light_value(3) = color
   
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

    
    Exit Sub

Light_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Render", Erl)
    Resume Next
    
End Sub

Private Function Light_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:
    
    If light_last = 0 Then
        Light_Next_Open = 1
        Exit Function
    End If

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
    
    On Error GoTo Light_Remove_All_Err
    
    Dim Index As Long
    
    For Index = 1 To light_last

        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index

        End If

    Next Index
    
    Light_Remove_All = True

    
    Exit Function

Light_Remove_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Remove_All", Erl)
    Resume Next
    
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    
    On Error GoTo Light_Destroy_Err
    
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

    
    Exit Sub

Light_Destroy_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Destroy", Erl)
    Resume Next
    
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
    '***************************************'
    'Author: Juan Martín Sotuyo Dodero
    'Last modified: 3/31/2003
    'Correctly erases a light
    '***************************************'
    
    On Error GoTo Light_Erase_Err
    
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X     As Integer
    Dim Y     As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_X - light_list(light_index).range
    min_y = light_list(light_index).map_Y - light_list(light_index).range
    max_x = light_list(light_index).map_X + light_list(light_index).range
    max_y = light_list(light_index).map_Y + light_list(light_index).range
    
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

    
    Exit Sub

Light_Erase_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Light_Erase", Erl)
    Resume Next
    
End Sub

Public Sub DrawInv()

End Sub

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)
    
    On Error GoTo Draw_FilledBox_Err
    

    Static box_rect     As RECT
    Static Outline      As RECT
    Static rgb_list(3)  As Long
    Static rgb_list2(3) As Long
    Static Vertex(3)    As TLVERTEX
    Static Vertex2(3)   As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .Bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y

    End With
    
    With Outline
        .Bottom = Y + Height + 2
        .Left = X - 2
        .Right = X + Width + 2
        .Top = Y - 2

    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))

    
    Exit Sub

Draw_FilledBox_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Draw_FilledBox", Erl)
    Resume Next
    
End Sub

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 11/19/2003
    'Similar to Grh_Render, but let´s you resize the Grh
    '**************************************************************
    On Error Resume Next

    Dim tile_width  As Integer
    Dim tile_height As Integer
    Dim grh_index   As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)

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
    grh_index = GrhData(Grh.grhindex).Frames(1)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.grhindex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(Grh.grhindex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32

    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, Width, Height, alpha_blend, Grh.angle

End Sub

Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)

    On Error Resume Next

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    'Modified by Juan Martín Sotuyo Dodero
    'Added centering
    '**************************************************************
    Dim tile_width  As Integer
    Dim tile_height As Integer
    Dim grh_index   As Long
    
    If Grh.grhindex = 0 Then Exit Sub
        
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)

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
    If Not Grh_Check(Grh.grhindex) Then Exit Sub
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
    Device_Box_Textured_Render grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).sX, GrhData(grh_index).sY, alpha_blend, Grh.angle

End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    
    On Error GoTo Grh_Check_Err
    

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= MaxGrhs Then
        Grh_Check = GrhData(grh_index).NumFrames

    End If

    
    Exit Function

Grh_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Grh_Check", Erl)
    Resume Next
    
End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean
    
    On Error GoTo Char_Check_Err
    

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (CharList(char_index).Heading > 0)

    End If
    
    
    Exit Function

Char_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Char_Check", Erl)
    Resume Next
    
End Function

Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function

    End If
    
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

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
    
    On Error GoTo Particle_Group_Check_Err
    

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

    
    Exit Function

Particle_Group_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Check", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Create(ByVal map_X As Integer, ByVal map_Y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
   Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Last Modify Date: 5/14/2003
    'Returns the particle_group_index if successful, else 0
    'Modified by Juan Martín Sotuyo Dodero
    'Modified by Augusto José Rando
    '**************************************************************
    
    On Error GoTo Particle_Group_Create_Err
    
    
    If (map_X <> -1) And (map_Y <> -1) Then
        If Map_Particle_Group_Get(map_X, map_Y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_X, map_Y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey

        End If

    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_X, map_Y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey

    End If

    
    Exit Function

Particle_Group_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Create", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
    
    On Error GoTo Particle_Group_Remove_Err
    

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

    
    Exit Function

Particle_Group_Remove_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Remove", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Remove_All() As Boolean
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    
    On Error GoTo Particle_Group_Remove_All_Err
    
    Dim Index As Long
    
    For Index = 1 To particle_group_last

        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index

        End If

    Next Index
    
    Particle_Group_Remove_All = True

    
    Exit Function

Particle_Group_Remove_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Remove_All", Erl)
    Resume Next
    
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

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    
    On Error GoTo Particle_Group_Destroy_Err
    
    Dim temp As particle_group
    
    MapData(particle_group_list(particle_group_index).map_X, particle_group_list(particle_group_index).map_Y).particle_group = 0
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

    
    Exit Sub

Particle_Group_Destroy_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Destroy", Erl)
    Resume Next
    
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_X As Integer, ByVal map_Y As Integer, _
   ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    
    On Error GoTo Particle_Group_Make_Err
    
                                
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
    
    'Map pos
    If (map_X <> -1) And (map_Y <> -1) Then
        particle_group_list(particle_group_index).map_X = map_X
        particle_group_list(particle_group_index).map_Y = map_Y

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
    If stream_type < 255 Then particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
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
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'Escribo La particula en el mapdata(x,y).particle_group :P
    If (map_X <> -1) And (map_Y <> -1) Then
        MapData(map_X, map_Y).particle_group = particle_group_index

    End If
    
    
    Exit Sub

Particle_Group_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Make", Erl)
    Resume Next
    
End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
    
    On Error GoTo Particle_Group_Render_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/15/2003
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    Dim loopc            As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move          As Boolean
    
    'Set colors

    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    
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
        
            Dim i As Byte
            '  Debug.Print particle_group_list(particle_group_index).grh_index_count
       
            ' For i = 1 To particle_group_list(particle_group_index).grh_index_count
            'Debug.Print particle_group_list(particle_group_index).grh_index_list(i)
            '   Next i

            ' Debug.Print particle_group_list(particle_group_index).grh_index_list(RandomNumber2(1, particle_group_list(particle_group_index).grh_index_count))
        
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
               screen_x, screen_y, _
               particle_group_list(particle_group_index).grh_index_list(1), _
               temp_rgb(), _
               particle_group_list(particle_group_index).alpha_blend, no_move, _
               particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
               particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
               particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
               particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
               particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
               particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
               particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
               particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
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
   
        Particle_Group_Destroy particle_group_index

    End If

    
    Exit Sub

Particle_Group_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Group_Render", Erl)
    Resume Next
    
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
   ByVal grh_index As Long, ByRef rgb_list() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
   Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    
    On Error GoTo Particle_Render_Err
    
                           
    '**************************************************************

    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/15/2003
    '**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index, alpha_blend
            temp_particle.X = RandomNumber(X1, X2) - (32 / 2)
            temp_particle.Y = RandomNumber(Y1, Y2) - (32 / 2)
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
    
    temp_particle.grh_resize = grh_resize
    temp_particle.grh_resizex = grh_resizex
    temp_particle.grh_resizey = grh_resizey

    'Draw i
    If grh_resize = True Then
        If temp_particle.Grh.grhindex Then
            Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub

        End If

    End If

    If temp_particle.Grh.grhindex Then
        Grh_Render temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, rgb_list(), True, True, alpha_blend

    End If

    
    Exit Sub

Particle_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Render", Erl)
    Resume Next
    
End Sub

Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long
    
    On Error GoTo Particle_Type_Get_Err
    

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type

    End If

    
    Exit Function

Particle_Type_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Particle_Type_Get", Erl)
    Resume Next
    
End Function

Public Function Map_Particle_Group_Get(ByVal map_X As Integer, ByVal map_Y As Integer) As Long
    
    On Error GoTo Map_Particle_Group_Get_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/20/2003
    'Checks to see if a tile position has a particle_group_index and return it
    '*****************************************************************
    If InMapBounds(map_X, map_Y) Then
        Map_Particle_Group_Get = MapData(map_X, map_Y).particle_group
    Else
        Map_Particle_Group_Get = 0

    End If

    
    Exit Function

Map_Particle_Group_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.Map_Particle_Group_Get", Erl)
    Resume Next
    
End Function

Public Sub MapCapture(ByRef Format As Boolean, ByVal ConBordes As Boolean, mx As Integer, my As Integer)
    
    On Error GoTo MapCapture_Err
    

    '*************************************************
    'Author: Torres Patricio(Pato)
    'Last modified:12/03/11
    '*************************************************
    
    Dim D3DWindow        As D3DPRESENT_PARAMETERS
    Dim Y                As Long     'Keeps track of where on map we are
    Dim X                As Long     'Keeps track of where on map we are
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim Grh              As Grh      'Temp Grh for show tile and blocked
    Static srcRect       As RECTL
    Static destRect      As RECTL

    Dim viewport As D3DVIEWPORT8
    viewport.Width = 3200
    viewport.Height = 3200
    viewport.MinZ = 1
    viewport.MaxZ = -1




    
    Dim maxX As Integer
    Dim maxY As Integer
    Dim minX As Integer
    Dim minY As Integer
    
    
    minX = mx + 1 - 8
    minY = my + 1 - 8
    maxX = minX + 100 + 8
    maxY = minY + 100 + 8
    
    If minX < 1 Then minX = 1
    If minY < 1 Then minY = 1
    
    If maxX > MapSize.Width Then maxX = MapSize.Width
    If maxY > MapSize.Height Then maxY = MapSize.Height
    
    DoEvents
    frmRenderer.picMap.Width = 800
    frmRenderer.picMap.Height = 800
    D3DDevice.SetViewport viewport
        Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0)
    
    Call D3DDevice.BeginScene
    
    DoEvents
    For Y = minY To maxY
        For X = minX To maxX
        
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).grhindex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), (X - minX) * 32, (Y - minY) * 32, 0, 1, MapData(X, Y).light_value())

            End If

            '******************************************
        
        Next X
    Next Y
    
    'Draw floor layer 2
    For Y = minY To maxY
        For X = minX To maxX
        
            'Layer 2 **********************************
            If (MapData(X, Y).Graphic(2).grhindex <> 0) And VerCapa2 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (X - minX) * 32, (Y - minY) * 32, 1, 1, MapData(X, Y).light_value())

            End If

            '******************************************
        
        Next X
    Next Y

    'Draw Transparent Layers
    For Y = minY To maxY
        For X = minX To maxX
                
            PixelOffsetXTemp = (X - minX) * 32
            PixelOffsetYTemp = (Y - minY) * 32
        
            With MapData(X, Y)
            
                'Object Layer **********************************
                If (.ObjGrh.grhindex <> 0) And VerObjetos Then
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value())

                End If

                '***********************************************
            
                'Layer 3 *****************************************
                If (.Graphic(3).grhindex <> 0) And VerCapa3 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value())

                End If

                '************************************************
            
            End With

        Next X
    Next Y
    
    'Draw layer 4
    For Y = minY To maxY
        For X = minX To maxX

            With MapData(X, Y)
            
                'Layer 4 **********************************
                If (.Graphic(4).grhindex <> 0) And VerCapa4 Then
                    
                    'Draw
                    Call Draw_Grh(.Graphic(4), (X - minX) * 32, (Y - minY) * 32, 1, 1, MapData(X, Y).light_value())
                    
                End If

                '**********************************
            
            End With

        Next X
    Next Y

    'Draw trans, bloqs, triggers and select tiles
    For Y = minY To maxY
        For X = minX To maxX

            With MapData(X, Y)
                PixelOffsetXTemp = (X - minX) * 32
                PixelOffsetYTemp = (Y - minY) * 32
            
                '**********************************
                Grh.FrameCounter = 1
                Grh.Started = 0

                If (.TileExit.Map <> 0) And VerTranslados Then
                    Grh.grhindex = 3
                
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 0, MapData(X, Y).light_value())

                End If
            
                'Show blocked tiles
                If (.Blocked = 1) And VerBlockeados Then
                    Grh.grhindex = 4
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 0, MapData(X, Y).light_value())

                End If

                '******************************************
            
            End With

        Next X
    Next Y

    ToWorldMap2 = True
    DoEvents
    Call D3DDevice.EndScene
    
    With srcRect
            .Left = 0
            .Top = 0
            .Right = 3200
            .Bottom = 3200
    End With
    


    
    Call D3DDevice.Present(srcRect, ByVal 0, frmRenderer.picMap.hWnd, ByVal 0)

    DoEvents

    viewport.Width = FrmMain.renderer.ScaleWidth
    viewport.Height = FrmMain.renderer.ScaleHeight
    viewport.MinZ = 1
    viewport.MaxZ = -1
    D3DDevice.SetViewport viewport
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''Guardo la imagen''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    Call frmRenderer.Capturar_Imagen(frmRenderer.picMap, frmRenderer.picMap)
     
    'Si no existe la carpeta de MiniMapas, la hacemos.
    If Not FileExist(App.Path & "\Render\", vbDirectory) Then
        Call MkDir(App.Path & "\Render\")
    End If
    
    If Not FileExist(App.Path & "\Mapas\", vbDirectory) Then
        Call MkDir(App.Path & "\Mapas\")
    End If
    
    If Not FileExist(App.Path & "\..\Resources\", vbDirectory) Then
        Call MkDir(App.Path & "\..\Resources\")
    End If
    
    If Not FileExist(App.Path & "\..\Resources\Minimapas\", vbDirectory) Then
        Call MkDir(App.Path & "\..\Resources\Minimapas\")
    End If
    
    
     'SaveImageAs App.Path & "\Render\" & MapName & ".png", FrmRender.picMap.hdc, FrmRender.picMap.ScaleWidth, FrmRender.picMap.ScaleHeight, CLng(100)
    
    SavePicture frmRenderer.picMap.Image, App.Path & "\Render\Partials\Mapa" & UserMap & "-" & minX & "-" & minY & "-" & MapSize.Width & "-" & MapSize.Height & ".bmp"
    'FrmRender.PicGrande.PaintPicture FrmRender.picMap.Image, ((minX - 1) \ 200) * 800, ((minY - 1) \ 200) * 800, 800, 800
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        DoEvents
    Exit Sub

MapCapture_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsDX8Engine.MapCapture", Erl)
    Resume Next
    
End Sub

