VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMinimap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public x As Integer
Public y As Integer
Public Opacity As Byte

Private Const MINIMAP_SIZE = 120
Private Const MINIMAP_BORDER = 5
Private Const MINIMAP_RADIUS = MINIMAP_SIZE / 2

Private MinimapTexture As Direct3DTexture8
Private MinimapSurface As Direct3DSurface8
Private MinimapProjection As D3DMATRIX

Private Const MINIMAP_BORDER_GRH = 26188
Private Const MINIMAP_MASK_GRAPHIC = 9793

Private Dragging As Boolean
Private DragOffsetX As Integer
Private DragOffsetY As Integer

Public Sub Class_Initialize()
    x = MINIMAP_BORDER * 2 + MINIMAP_RADIUS
    y = 732 - MINIMAP_BORDER * 2 - MINIMAP_RADIUS
    Opacity = 255
End Sub

Public Sub SetupMinimapTexture()
    ' Inicializo la textura
    Set MinimapTexture = DirectDevice.CreateTexture(MINIMAP_SIZE, MINIMAP_SIZE, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set MinimapSurface = MinimapTexture.GetSurfaceLevel(0)
    Call D3DXMatrixOrthoOffCenterLH(MinimapProjection, 0, MINIMAP_SIZE, MINIMAP_SIZE, 0, -1#, 1#)
End Sub

Public Sub Draw()
    If QueRender <> 0 Then Exit Sub
    If Opacity = 0 Then Exit Sub
    
    Dim Color(3) As RGBA
    Call RGBAList(Color, 255, 255, 255, Opacity)

    ' Dibujo el marco
    Call Draw_GrhIndexColor(MINIMAP_BORDER_GRH, x - MINIMAP_RADIUS - MINIMAP_BORDER, y - MINIMAP_RADIUS - MINIMAP_BORDER, Color)
    ' Detengo la escena actual
    Call Engine_PauseScene
    
    ' Dibujo en la texture la sección del mapa en donde está el usuario
    Call DirectDevice.SetRenderTarget(MinimapSurface, pZBuffer, 0)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, MinimapProjection)
    Call Engine_BeginScene
    
    Dim mapX As Integer, mapY As Integer
    mapX = (UserPos.x - AddtoUserPos.x) * 2 - OffsetCounterX / 16 - MINIMAP_RADIUS
    mapY = (UserPos.y - AddtoUserPos.y) * 2 - OffsetCounterY / 16 - MINIMAP_RADIUS
    
    Call Batch_Textured_Box_Advance(0, 0, MINIMAP_SIZE, MINIMAP_SIZE, mapX, mapY, UserMap + MINIMAP_PNG, MINIMAP_SIZE, MINIMAP_SIZE, COLOR_WHITE)
    
    Call Engine_PauseScene
    Call ResetRenderTarget
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    
    Call Engine_ResumeScene

    ' Cambio el pixel shader
    Call DirectDevice.SetPixelShader(MaskPixelShader)
    
    ' Seteo la mascara
    Dim Ignore As Long
    Call DirectDevice.SetTexture(1, SurfaceDB.GetTexture(MINIMAP_MASK_GRAPHIC, Ignore, Ignore))

    Call DirectDevice.SetPixelShaderConstant(0, RGBA_2_D3DCOLORVALUE(Color(0)), 1)
    
    With SpriteBatch
        Call .SetTexture(MinimapTexture)
        Call .SetAlpha(False)
        Call .Draw(x - MINIMAP_RADIUS, y - MINIMAP_RADIUS, MINIMAP_SIZE, MINIMAP_SIZE, COLOR_WHITE)
        Call .Flush
    End With
    
    ' Reseteo el pixel shader
    Call DirectDevice.SetPixelShader(0)
    
    With charlist(UserCharIndex)
        Dim CharScale As Single
        CharScale = 1 / 6
        Dim BodyWidth As Integer, BodyHeight As Integer
        BodyWidth = GrhData(.Body.Walk(.Heading).GrhIndex).pixelWidth * CharScale
        BodyHeight = GrhData(.Body.Walk(.Heading).GrhIndex).pixelHeight * CharScale
        
        If .Muerto Then Call Copy_RGBAList_WithAlpha(Color, Color, Opacity * 0.5)
    
        If .Navegando Then
            Call Grh_Render_Advance(.Body.Walk(.Heading), x - BodyWidth * 0.5, y + 2 - BodyHeight, BodyHeight, BodyWidth, Color, False, False)
        Else
            Call Grh_Render_Advance(.Body.Walk(.Heading), x - BodyWidth * 0.5, y + 1 - BodyHeight, BodyHeight, BodyWidth, Color, False, False)
            
            If .Head.Head(.Heading).GrhIndex Then
                Dim HeadWidth As Integer, HeadHeight As Integer
                HeadWidth = GrhData(.Head.Head(.Heading).GrhIndex).pixelWidth * CharScale
                HeadHeight = GrhData(.Head.Head(.Heading).GrhIndex).pixelHeight * CharScale
            
                Call Grh_Render_Advance(.Head.Head(.Heading), x - HeadWidth * 0.5, y + 1 - HeadHeight + .Body.HeadOffset.y * CharScale, HeadHeight, HeadWidth, Color, False, False)
            End If
            
        End If
    End With

End Sub

Public Sub MouseDown(ByVal Button As Integer, ByVal MouseX As Integer, ByVal MouseY As Integer)
    If Opacity = 0 Then Exit Sub
    If Button = vbLeftButton Then
        Dim dX As Long, dY As Long
        dX = MouseX - x
        dY = MouseY - y
    
        If dX * dX + dY * dY <= MINIMAP_RADIUS * MINIMAP_RADIUS Then
            Dragging = True
            DragOffsetX = dX
            DragOffsetY = dY
        End If
    End If
End Sub

Public Sub MouseMove(ByVal MouseX As Integer, ByVal MouseY As Integer)
    If Dragging Then
        x = Clamp(MouseX - DragOffsetX, MINIMAP_BORDER + MINIMAP_RADIUS, MainViewWidth - MINIMAP_BORDER - MINIMAP_RADIUS)
        y = Clamp(MouseY - DragOffsetY, MINIMAP_BORDER + MINIMAP_RADIUS, MainViewHeight - MINIMAP_BORDER - MINIMAP_RADIUS)
    End If
End Sub

Public Function MouseUp(ByVal Button As Integer, ByVal MouseX As Integer, ByVal MouseY As Integer) As Boolean
    If Opacity = 0 Then Exit Function
    If Dragging Then
        Dragging = False
    ElseIf EsGM And Button = vbRightButton Then
        Dim dX As Long, dY As Long
        dX = MouseX - x
        dY = MouseY - y

        If dX * dX + dY * dY <= MINIMAP_RADIUS * MINIMAP_RADIUS Then
            Call WriteWarpChar("YO", UserMap, UserPos.x + dX \ 2, UserPos.y + dY \ 2)
            MouseUp = True
        End If
    End If
End Function
