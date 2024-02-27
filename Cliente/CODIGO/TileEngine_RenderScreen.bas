Attribute VB_Name = "TileEngine_RenderScreen"
Option Explicit

'Letter showing on screen
Public letter_text           As String
Public letter_grh            As grh
Public map_letter_grh        As grh
Public map_letter_grh_next   As Long
Public map_letter_a          As Single
Public map_letter_fadestatus As Byte

Public Const MINIMAP_PNG As Integer = 9800

Sub RenderScreen(ByVal center_x As Integer, ByVal center_y As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal HalfTileWidth As Integer, ByVal HalfTileHeight As Integer)
    
    On Error GoTo RenderScreen_Err
    

    '**************************************************************
    ' Author: Aaron Perkins
    ' Last Modify Date: 23/11/2020
    ' Modified by: Juan Martín Sotuyo Dodero (Maraxus)
    ' Last modified by: Alexis Caraballo (WyroX)
    ' Renders everything to the viewport
    '**************************************************************
    
    
    
    Dim X                   As Integer, Y As Integer
    
    Dim ScreenX             As Integer, ScreenY As Integer
    
    Dim MinX                As Integer, MinY As Integer ' Top-left corner
    
    Dim MaxX                As Integer, MaxY As Integer ' Bottom-right corner
    
    Dim MinBufferedX        As Integer, MinBufferedY As Integer ' Start tile buffered
    
    Dim MaxBufferedX        As Integer, MaxBufferedY As Integer ' End tile buffered

    Dim StartX              As Integer, StartY As Integer ' Pixel start

    Dim StartBufferedX      As Integer, StartBufferedY As Integer ' Pixel start buffered
    
    Dim DeltaTime                   As Long

    Dim TempColor(3)        As RGBA
    Dim ColorBarraPesca(3)  As RGBA

    ' Tiles that are in range
    MinX = center_x - HalfTileWidth
    MinY = center_y - HalfTileHeight
    MaxX = center_x + HalfTileWidth
    MaxY = center_y + HalfTileHeight
    
    If PixelOffsetX < 0 Then
        MaxX = MaxX + 1
    ElseIf PixelOffsetX > 0 Then
        MinX = MinX - 1
    End If
    
    If PixelOffsetY < 0 Then
        MaxY = MaxY + 1
    ElseIf PixelOffsetY > 0 Then
        MinY = MinY - 1
    End If

    MinX = max(MinX, MapSize.XMin)
    MinY = max(MinY, MapSize.YMin)
    MaxX = min(MaxX, MapSize.Width)
    MaxY = min(MaxY, MapSize.Height)

    MinBufferedX = max(MinX - TileBufferSize.Left, MapSize.XMin)
    MinBufferedY = max(MinY - TileBufferSize.Top, MapSize.YMin)
    MaxBufferedX = min(MaxX + TileBufferSize.Right, MapSize.Width)
    MaxBufferedY = min(MaxY + TileBufferSize.Bottom, MapSize.Height)

    ' Pixel offset start
    StartX = (MinX - center_x + HalfWindowTileWidth) * TilePixelWidth + PixelOffsetX + StartScreenX
    StartY = (MinY - center_y + HalfWindowTileHeight) * TilePixelHeight + PixelOffsetY + StartScreenY
    StartBufferedX = (MinBufferedX - center_x + HalfWindowTileWidth) * TilePixelWidth + PixelOffsetX + StartScreenX
    StartBufferedY = (MinBufferedY - center_y + HalfWindowTileHeight) * TilePixelHeight + PixelOffsetY + StartScreenY

    
    If UpdateLights Then
        Call RestaurarLuz
        Call MapUpdateGlobalLightRender
        UpdateLights = False
    End If
    
    'Call SpriteBatch.BeginPrecalculated(StartX, StartY)
    
    ' *********************************
    ' Layer 1 loop
    ScreenY = StartY
    For y = MinY To MaxY
        ScreenX = StartX
        For x = MinX To MaxX
            With MapData(rrX(x), rrY(y))

                ' Layer 1 *********************************
                Call Draw_Grh_Precalculated(ScreenX, ScreenY, .Graphic(1), .light_value, (.Blocked And FLAG_AGUA) <> 0, (.Blocked And FLAG_LAVA) <> 0, X, Y, MinX, MaxX, MinY, MaxY)
                '******************************************

            End With
            ScreenX = ScreenX + TilePixelWidth
        Next x
        ScreenY = ScreenY + TilePixelHeight
    Next y

    'Call SpriteBatch.EndPrecalculated

    ' *********************************
    ' Layer 2 & small objects loop
    Call DirectDevice.SetRenderState(D3DRS_ALPHATESTENABLE, True) ' Para no pisar los reflejos
    
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            With MapData(rrX(x), rrY(y))
                
                ' Layer 2 *********************************
                If .Graphic(2).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(2), ScreenX, ScreenY, 1, 1, .light_value)
                End If
                '******************************************
            
            End With

            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
    
 
    

    Dim grhSpellArea As grh
    grhSpellArea.GrhIndex = 20058
    
    Dim temp_color(3) As RGBA
    
    Call SetRGBA(temp_color(0), 255, 20, 25, 255)
    Call SetRGBA(temp_color(1), 0, 255, 25, 255)
    Call SetRGBA(temp_color(2), 55, 255, 55, 255)
    Call SetRGBA(temp_color(3), 145, 70, 70, 255)
    
   ' Call SetRGBA(MapData(rrX(15), rrY(15)).light_value(0), 255, 20, 20)
    'size 96x96 - mitad = 48
    If casteaArea And MouseX > 0 And MouseY > 0 And frmMain.MousePointer = 2 Then
        Call Draw_Grh(grhSpellArea, MouseX - 48, MouseY - 48, 0, 1, temp_color, True, 70)
    End If
    
     ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            With MapData(rrX(x), rrY(y))
                
                ' Objects *********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Select Case ObjData(.OBJInfo.ObjIndex).ObjType
                        Case eObjType.otArboles, eObjType.otPuertas, eObjType.otTeleport, eObjType.otCarteles, eObjType.OtPozos, eObjType.otYacimiento, eObjType.OtCorreo, eObjType.otFragua, eObjType.OtDecoraciones
                            Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value)

                        Case Else
                            ' Objetos en el suelo (items, decorativos, etc)
                            
                             If ((.Blocked And FLAG_AGUA) <> 0) And .Graphic(2).GrhIndex = 0 Then
                             
                                object_angle = (object_angle + (timerElapsedTime * 0.002))
                                
                                .light_value(1).A = 85
                                .light_value(3).A = 85
                                
                                Call Draw_Grh_ItemInWater(.ObjGrh, ScreenX, ScreenY, False, False, .light_value, False, , , (object_angle + (x Mod 100) * 45 + (y Mod 100) * 90))
                                
                                .light_value(1).A = 255
                                .light_value(3).A = 255
                                .light_value(0).A = 255
                                .light_value(2).A = 255
                            Else
                                Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value)
                            End If
                    End Select
                End If
                '******************************************

            End With
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y

    Call DirectDevice.SetRenderState(D3DRS_ALPHATESTENABLE, False)
    
    ' *********************************
    '  Layer 3 & chars
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            With MapData(rrX(x), rrY(y))
                ' Chars ***********************************
                If .charindex = UserCharIndex Then 'evitamos reenderizar un clon del usuario
                    If x <> UserPos.x Or y <> UserPos.y Then
                        .charindex = 0
                    End If
                End If
                
                'If .CharFantasma.Activo Then
                '
                '    If .CharFantasma.AlphaB > 0 Then
                '
                '        .CharFantasma.AlphaB = .CharFantasma.AlphaB - (timerTicksPerFrame * 30)
                '
                '        'Redondeamos a 0 para prevenir errores
                '        If .CharFantasma.AlphaB < 0 Then .CharFantasma.AlphaB = 0
                '
                '        Call Copy_RGBAList_WithAlpha(TempColor, .light_value, .CharFantasma.AlphaB)
                '
                '        'Seteamos el color
                '        If .CharFantasma.Heading = 1 Or .CharFantasma.Heading = 2 Then
                '            Call Draw_Grh(.CharFantasma.Escudo, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Body, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Head, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Casco, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Arma, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                '        Else
                '            Call Draw_Grh(.CharFantasma.Body, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Head, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Escudo, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Casco, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                '            Call Draw_Grh(.CharFantasma.Arma, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                '        End If
'
'                    Else
'                        .CharFantasma.Activo = False
'
'                    End If
'
'                End If
                
                If .charindex <> 0 Then
                    If charlist(.charindex).active = 1 Then
                        'If mascota.visible And .charindex = UserCharIndex Then
                          '  Call Mascota_Render(.charindex, PixelOffsetX, PixelOffsetY)
                        'End If
                        Call Char_Render(.charindex, ScreenX, ScreenY, x, y)
                    End If
                End If
                '******************************************
                
            End With
            ScreenX = ScreenX + TilePixelWidth
        Next x
        
        ' Recorremos de nuevo esta fila para dibujar objetos grandes y capa 3 encima de chars
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            With MapData(rrX(x), rrY(y))
                ' Objects *********************************
                If .ObjGrh.GrhIndex <> 0 Then
                           
                    Select Case ObjData(.OBJInfo.ObjIndex).ObjType
                         
                        Case eObjType.otArboles
                          
                            Call Draw_Sombra(.ObjGrh, ScreenX, ScreenY, 1, 1, False, x, y)

                            ' Debajo del arbol
                            If Abs(UserPos.x - x) < 3 And (Abs(UserPos.y - y)) < 8 And (Abs(UserPos.y) < y) Then
    
                                If .ArbolAlphaTimer <= 0 Then
                                    .ArbolAlphaTimer = lastMove
                                End If
    
                                DeltaTime = FrameTime - .ArbolAlphaTimer
    
                                Call Copy_RGBAList_WithAlpha(TempColor, .light_value, IIf(DeltaTime > ARBOL_ALPHA_TIME, ARBOL_MIN_ALPHA, 255 - DeltaTime / ARBOL_ALPHA_TIME * (255 - ARBOL_MIN_ALPHA)))
                                Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, TempColor, False)
    
                            Else    ' Lejos del arbol
                                If .ArbolAlphaTimer = 0 Then
                                    Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False)
    
                                Else
                                    If .ArbolAlphaTimer > 0 Then
                                        .ArbolAlphaTimer = -lastMove
                                    End If
    
                                    DeltaTime = FrameTime + .ArbolAlphaTimer
    
                                    If DeltaTime > ARBOL_ALPHA_TIME Then
                                        .ArbolAlphaTimer = 0
                                        Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False)
                                    Else
                                        Call Copy_RGBAList_WithAlpha(TempColor, .light_value, ARBOL_MIN_ALPHA + DeltaTime * (255 - ARBOL_MIN_ALPHA) / ARBOL_ALPHA_TIME)
                                        Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, TempColor, False)
                                    End If
                                End If
    
                            End If
                        
                        Case eObjType.otPuertas, eObjType.otTeleport, eObjType.otCarteles, eObjType.OtPozos, eObjType.otYacimiento, eObjType.OtCorreo, eObjType.otYunque, eObjType.otFragua, eObjType.OtDecoraciones
                            ' Objetos grandes (menos árboles)
                            Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False)
                            
                        'Case Else
                        '    Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
                    
                    End Select
                End If
                '******************************************
                
                'Layer 3 **********************************
                If .Graphic(3).GrhIndex <> 0 Then

                    If (.Blocked And FLAG_ARBOL) <> 0 Then
                        
                        
                       ' Call Draw_Sombra(.Graphic(3), ScreenX, ScreenY, 1, 1, False, x, y)

                        ' Debajo del arbol
                        If Abs(UserPos.x - x) <= 3 And (Abs(UserPos.y - y)) < 8 And (Abs(UserPos.y) < y) Then

                            If .ArbolAlphaTimer <= 0 Then
                                .ArbolAlphaTimer = lastMove
                            End If

                            DeltaTime = FrameTime - .ArbolAlphaTimer

                            Call Copy_RGBAList_WithAlpha(TempColor, .light_value, IIf(DeltaTime > ARBOL_ALPHA_TIME, ARBOL_MIN_ALPHA, 255 - DeltaTime / ARBOL_ALPHA_TIME * (255 - ARBOL_MIN_ALPHA)))
                            Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, TempColor, False)

                        Else    ' Lejos del arbol
                            If .ArbolAlphaTimer = 0 Then
                                Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False)

                            Else
                                If .ArbolAlphaTimer > 0 Then
                                    .ArbolAlphaTimer = -lastMove
                                End If

                                DeltaTime = FrameTime + .ArbolAlphaTimer

                                If DeltaTime > ARBOL_ALPHA_TIME Then
                                    .ArbolAlphaTimer = 0
                                    Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False)
                                Else
                                    Call Copy_RGBAList_WithAlpha(TempColor, .light_value, ARBOL_MIN_ALPHA + DeltaTime * (255 - ARBOL_MIN_ALPHA) / ARBOL_ALPHA_TIME)
                                    Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, TempColor, False)
                                End If
                            End If

                        End If

                    Else
                        If AgregarSombra(.Graphic(3).GrhIndex) Then
                            Call Draw_Sombra(.Graphic(3), ScreenX, ScreenY, 1, 1, False, x, y)
                        End If

                        Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False)

                    End If

                End If
                '******************************************
            End With
            
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
    
    'If InfoItemsEnRender And tX And tY Then
        'With MapData(rrX(tX), rrY(tY))
        '    If .OBJInfo.ObjIndex Then
        '        If Not ObjData(.OBJInfo.ObjIndex).Agarrable Then
        '            Dim Text As String, Amount As String
        '            If .OBJInfo.Amount > 1000 Then
        '                Amount = Round(.OBJInfo.Amount * 0.001, 1) & "K"
        '            Else
        '                Amount = .OBJInfo.Amount
        '            End If
        '            Text = ObjData(.OBJInfo.ObjIndex).Name & " (" & Amount & ")"
        '            Call Engine_Text_Render(Text, MouseX + 15, MouseY, COLOR_WHITE, , , , 160)
        '        End If
        '    End If
        'End With
    'End If

    ' *********************************
    ' Particles loop
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            With MapData(rrX(x), rrY(y))
                ' Particles *******************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX + 16, ScreenY + 16)
                End If
                '******************************************
            End With
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
 
    ' *********************************
    ' Layer 4 loop
    If HayLayer4 Then

        ' Actualizo techos
        Dim Trigger As eTrigger
        For Trigger = LBound(RoofsLight) To UBound(RoofsLight)

            ' Si estoy bajo este techo
            If Trigger = MapData(rrX(UserPos.x), rrY(UserPos.y)).Trigger Then
            
                If RoofsLight(Trigger) > 0 Then
                    ' Reduzco el alpha
                    RoofsLight(Trigger) = RoofsLight(Trigger) - timerTicksPerFrame * 48
                    If RoofsLight(Trigger) < 0 Then RoofsLight(Trigger) = 0
                End If

            ElseIf RoofsLight(Trigger) < 255 Then
            
                ' Aumento el alpha
                RoofsLight(Trigger) = RoofsLight(Trigger) + timerTicksPerFrame * 48
                If RoofsLight(Trigger) > 255 Then RoofsLight(Trigger) = 255
            
            End If
            
        Next

        ScreenY = StartBufferedY

        For y = MinBufferedY To MaxBufferedY
            ScreenX = StartBufferedX
            
            For x = MinBufferedX To MaxBufferedX
                With MapData(rrX(x), rrY(y))
                    ' Layer 4 - roofs *******************************
                    If .Graphic(4).GrhIndex Then

                        Trigger = NearRoof(x, y)

                        If Trigger Then
                            Call Copy_RGBAList_WithAlpha(TempColor, .light_value, RoofsLight(Trigger))
                            Call Draw_Grh(.Graphic(4), ScreenX, ScreenY, 1, 1, TempColor)
                        Else
                            Call Draw_Grh(.Graphic(4), ScreenX, ScreenY, 1, 1, .light_value)
                        End If
    
                    End If
                    '******************************************
                End With
                ScreenX = ScreenX + TilePixelWidth
            Next x
            
            ScreenY = ScreenY + TilePixelHeight
        Next y
        
    End If
    

    
    
    
  
    
       
    If mascota.dialog <> "" And mascota.visible Then
        Call Engine_Text_Render(mascota.dialog, mascota.PosX + 14 - CInt(Engine_Text_Width(mascota.dialog, True) / 2) + 150, mascota.PosY - Engine_Text_Height(mascota.dialog, True) - 25 + 150, mascota_text_color(), 1, True, , mascota.Color(0).A)
    End If
        
    
    
    ' *********************************
    ' FXs and dialogs loop
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            With MapData(rrX(x), rrY(y))


                ' Dialogs *******************************
                If .charindex <> 0 Then
                
                    If charlist(.charindex).active = 1 Then
                    
                        Call Char_TextRender(.charindex, ScreenX, ScreenY, x, y)
                    
                    End If
                    
                End If
                '******************************************

                ' Render text value *******************************
                Dim i As Long
                If False Then
                    For i = 1 To UBound(.DialogEffects)
                        With .DialogEffects(i)
                            If LenB(.Text) <> 0 Then
                                Dim DialogTime As Long
                                DialogTime = FrameTime - .Start
            
                                If DialogTime > 1300 Then
                                    .Text = vbNullString
                                Else
                                    If DialogTime > 900 Then
                                        Call RGBAList(TempColor, .Color.r, .Color.G, .Color.B, .Color.A * (1300 - DialogTime) * 0.0025)
                                    Else
                                        Call RGBAList(TempColor, .Color.r, .Color.G, .Color.B, .Color.A)
                                    End If
                            
                                    Engine_Text_Render_Efect 0, .Text, ScreenX + 16 - Int(Engine_Text_Width(.Text, False) * 0.5) + .offset.x, ScreenY - Engine_Text_Height(.Text, False) + .offset.y - DialogTime * 0.025, TempColor, 1, False
                    
                                End If
                            End If
                        End With
                    Next
                End If
                '******************************************

                ' FXs *******************************
                If .FxCount > 0 Then
                    For i = 1 To .FxCount

                        If .FxList(i).FxIndex > 0 And .FxList(i).started <> 0 Then
    
                            Call RGBAList(TempColor, 255, 255, 255, 220)

                            If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                            
                                Call Draw_GrhFX(.FxList(i), ScreenX + FxData(.FxList(i).FxIndex).OffsetX, ScreenY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, TempColor(), False)

                            Else
                                Call Draw_GrhFX(.FxList(i), ScreenX + FxData(.FxList(i).FxIndex).OffsetX, ScreenY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, TempColor(), True)

                            End If

                        End If

                        If .FxList(i).started = 0 Then .FxList(i).FxIndex = 0

                    Next i

                    If .FxList(.FxCount).started = 0 Then .FxCount = .FxCount - 1

                End If
                '******************************************
                End With
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y

    If bRain Then
    
        If MapDat.Lluvia Then
        
            'Screen positions were hardcoded by now
            ScreenX = 250
            ScreenY = 0
            
            Call Particle_Group_Render(MeteoParticle, ScreenX, ScreenY)
            
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

        End If

    End If

    If AlphaNiebla Then
    
        If MapDat.Niebla Then Call Engine_Weather_UpdateFog

    End If

    If bNieve Then
    
        If MapDat.Nieve Then
        
            If Graficos_Particulas.Engine_MeteoParticle_Get <> 0 Then
            
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                
                Call Particle_Group_Render(MeteoParticle, ScreenX, ScreenY)

            End If

        End If

    End If
    
    Call Effect_Render_All
    
    
    'Call Engine_Draw_Load(300, 350, 250, 250, RGBA_From_Comp(200, 200, 200, 255))
    Call renderCooldowns(710, 582)
    
    'Consola renderizada
    Call ConsoleDialog.Draw
    
    Call Minimap.Draw
    
        'Render minimapa
    If UserCharIndex > 0 And (minimapa_visible Or minimapa_alpha > 0) Then
        Dim x_map As Integer
        Dim y_map As Integer
        
        If minimapa_visible And minimapa_alpha < 255 Then
            minimapa_alpha = minimapa_alpha + timerTicksPerFrame * 40
            If minimapa_alpha > 255 Then minimapa_alpha = 255
        ElseIf Not minimapa_visible And minimapa_alpha > 0 Then
            minimapa_alpha = minimapa_alpha - timerTicksPerFrame * 40
            If minimapa_alpha < 0 Then minimapa_alpha = 0
        End If
        
        x_map = (UserPos.x - AddtoUserPos.x) * 2 - OffsetCounterX / 16 - (MainViewWidth / 2)
        y_map = (UserPos.y - AddtoUserPos.y) * 2 - OffsetCounterY / 16 - (MainViewHeight / 2)

        Call RGBAList(color_minimap, 255, 255, 255, minimapa_alpha)

        Dim srcX As Integer
        Dim srcY As Integer
        Dim pjX As Integer
        Dim pjY As Integer
        Dim mapX As Integer
        Dim mapY As Integer
        
        If MapSize.Width * 2 > MainViewWidth Then
            If x_map < 0 Then
                srcX = 0
                pjX = x_map
            ElseIf x_map > MapSize.Width * 2 - MainViewWidth Then
                srcx = MapSize.Width * 2 - MainViewWidth
                pjX = x_map - MapSize.Width * 2 + MainViewWidth
            Else
                srcX = x_map
                pjX = 0
            End If
        Else
            pjX = UserPos.x * 2 - MainViewWidth / 2
        End If
        
        If MapSize.Height * 2 > MainViewHeight Then
            If y_map < 0 Then
                srcY = 0
                pjY = y_map
            ElseIf y_map > MapSize.Height * 2 - MainViewHeight Then
                srcY = MapSize.Height * 2 - MainViewHeight
                pjY = y_map - MapSize.Height * 2 + MainViewHeight
            Else
                srcY = y_map
                pjY = 0
            End If
        Else
            pjY = UserPos.y * 2 - MainViewHeight / 2
        End If
        
        If MapSize.Width * 2 < MainViewWidth Then
            mapX = (MainViewWidth - MapSize.Width * 2) / 2
        End If
        If MapSize.Height * 2 < MainViewHeight Then
            mapY = (MainViewHeight - MapSize.Height * 2) / 2
        End If
        
        pjX = pjX + mapX
        pjY = pjY + mapY
        
        Call Batch_Textured_Box_Advance(mapX, mapY, IIf(MapSize.Width * 2 > MainViewWidth, MainViewWidth, MapSize.Width * 2), IIf(MapSize.Height * 2 > MainViewHeight, MainViewHeight, MapSize.Height * 2), srcX - 1, srcY, UserMap + MINIMAP_PNG, IIf(MapSize.Width * 2 > MainViewWidth, MainViewWidth, MapSize.Width * 2), IIf(MapSize.Height * 2 > MainViewHeight, MainViewHeight, MapSize.Height * 2), color_minimap, False, 0)

        
        With charlist(UserCharIndex)
            Dim CharScale As Single
            CharScale = IIf(MapSize.Width * 2 > MainViewWidth, 1 / 6, 1 / 3)
            Dim BodyWidth As Integer, BodyHeight As Integer
            BodyWidth = GrhData(.Body.Walk(.Heading).GrhIndex).pixelWidth * CharScale
            BodyHeight = GrhData(.Body.Walk(.Heading).GrhIndex).pixelHeight * CharScale
            
            Dim Color(3) As RGBA
            Call Copy_RGBAList_WithAlpha(Color, color_minimap, IIf(.Muerto, minimapa_alpha * 0.5, minimapa_alpha))
        
            If .Navegando Then
                Call Grh_Render_Advance(.Body.Walk(.Heading), pjX + Int((MainViewWidth - BodyWidth) * 0.5), pjY + MainViewHeight * 0.5 + 2 - BodyHeight, BodyHeight, BodyWidth, Color, False, False)
            Else
                Call Grh_Render_Advance(.Body.Walk(.Heading), pjX + Int((MainViewWidth - BodyWidth) * 0.5), pjY + MainViewHeight * 0.5 + 1 - BodyHeight, BodyHeight, BodyWidth, Color, False, False)
                
                If .Head.Head(.Heading).GrhIndex Then
                    Dim HeadWidth As Integer, HeadHeight As Integer
                    HeadWidth = GrhData(.Head.Head(.Heading).GrhIndex).pixelWidth * CharScale
                    HeadHeight = GrhData(.Head.Head(.Heading).GrhIndex).pixelHeight * CharScale
                
                    Call Grh_Render_Advance(.Head.Head(.Heading), pjX + Int((MainViewWidth - HeadWidth) * 0.5), pjY + MainViewHeight * 0.5 + 1 - HeadHeight + .Body.HeadOffset.y * CharScale, HeadHeight, HeadWidth, Color, False, False)
                End If
                
            End If
        End With
        
        If MouseMapX > 0 And MouseMapY > 0 Then
            Call Engine_Text_Render(MouseMapX & ", " & MouseMapY, MouseX, MouseY + 24, color_minimap, 1, False, , minimapa_alpha)
            If MouseZona > 0 Then
                Call Engine_Text_Render(MapZona(UserMap).Zona(MouseZona).Zona_name, MouseX, MouseY + 42, color_minimap, 1, False, , minimapa_alpha)
            End If
        End If
        
    End If
    
    If InvasionActual Then
        
        Call Engine_Draw_Box(190, 550, 356, 36, RGBA_From_Comp(0, 0, 0, 200))
        Call Engine_Draw_Box(193, 553, 3.5 * InvasionPorcentajeVida, 30, RGBA_From_Comp(20, 196, 255, 200))
        
        Call Engine_Draw_Box(340, 586, 54, 9, RGBA_From_Comp(0, 0, 0, 200))
        Call Engine_Draw_Box(342, 588, 0.5 * InvasionPorcentajeTiempo, 5, RGBA_From_Comp(220, 200, 0, 200))
        
    End If

    If Pregunta Then
        
        Call Engine_Draw_Box(283, 170, 190, 100, RGBA_From_Comp(150, 20, 3, 200))
        Call Engine_Draw_Box(288, 175, 180, 90, RGBA_From_Comp(25, 25, 23, 200))

        Dim preguntaGrh As grh
        Call InitGrh(preguntaGrh, 32120)
        
        Call Engine_Text_Render(PreguntaScreen, 290, 180, COLOR_WHITE, 1, True)
        
        Call Draw_Grh(preguntaGrh, 416, 233, 1, 0, COLOR_WHITE, False)

    End If

    If cartel Then

        Call RGBAList(TempColor, 255, 255, 255, 220)
        
        Dim TempGrh  As grh
        Call InitGrh(TempGrh, GrhCartel)
        
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, TempColor(), False)
        
        Call Engine_Text_Render(Leyenda, CInt(clicX - 100), CInt(clicY - 130), TempColor(), 1, False)

    End If


    Call RenderScreen_NombreMapa
    
    
    If PescandoEspecial Then
        Call RGBAList(ColorBarraPesca, 255, 255, 255)
        Dim grh As grh
        grh.GrhIndex = GRH_BARRA_PESCA
        Call Draw_Grh(grh, 239, 550, 0, 0, ColorBarraPesca())
        grh.GrhIndex = GRH_CURSOR_PESCA
        Call Draw_Grh(grh, 271 + PosicionBarra, 558, 0, 0, ColorBarraPesca())
        For i = 1 To MAX_INTENTOS
            If intentosPesca(i) = 1 Then
                grh.GrhIndex = GRH_CIRCULO_VERDE
                Call Draw_Grh(grh, 394 + (i * 10), 573, 0, 0, ColorBarraPesca())
            ElseIf intentosPesca(i) = 2 Then
                grh.GrhIndex = GRH_CIRCULO_ROJO
                Call Draw_Grh(grh, 394 + (i * 10), 573, 0, 0, ColorBarraPesca())
            End If
        Next i
                
        If PosicionBarra <= 0 Then
            DireccionBarra = 1
            PuedeIntentar = True
        ElseIf PosicionBarra > 199 Then
            DireccionBarra = -1
            PuedeIntentar = True
        End If
        If PosicionBarra < 0 Then
            PosicionBarra = 0
        ElseIf PosicionBarra > 199 Then
            PosicionBarra = 199
        End If
        '90 - 111 es incluido (saca el pecesito)
        PosicionBarra = PosicionBarra + (DireccionBarra * VelocidadBarra * timerElapsedTime * 0.2)
        
        
        If (GetTickCount() - startTimePezEspecial) >= 20000 Then
            PescandoEspecial = False
            Call AddToConsole("El pez ha roto tu linea de pesca.", 255, 0, 0, 1, 0)
            Call WriteRomperCania
        End If
        
    End If
    
    If cartel_visible Then Call RenderScreen_Cartel
    Exit Sub

RenderScreen_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_RenderScreen.RenderScreen", Erl)
    Resume Next
    
End Sub

Private Sub RenderScreen_NombreMapa()
    
    On Error GoTo RenderScreen_NombreMapa_Err
    
    
    If map_letter_fadestatus > 0 Then
    
        If map_letter_fadestatus = 1 Then
        
            map_letter_a = map_letter_a + (timerTicksPerFrame * 3.5)

            If map_letter_a >= 255 Then
                map_letter_a = 255
                map_letter_fadestatus = 2

            End If

        Else
        
            map_letter_a = map_letter_a - (timerTicksPerFrame * 3.5)

            If map_letter_a <= 0 Then
            
                map_letter_fadestatus = 0
                map_letter_a = 0
                 
                If map_letter_grh_next > 0 Then
                    map_letter_grh.GrhIndex = map_letter_grh_next
                    map_letter_fadestatus = 1
                    map_letter_grh_next = 0

                End If
                
            End If

        End If

    End If
    
    If Len(letter_text) Then
        
        Dim Color(3) As RGBA
        Call RGBAList(Color(), 179, 95, 0, map_letter_a)
        
        Call Grh_Render(letter_grh, 250, 300, Color())
        
        Call Engine_Text_RenderGrande(letter_text, 360 - Engine_Text_Width(letter_text, False, 4) / 2, 0 + ConsoleDialog.console_height + 10, Color(), 5, False, , CInt(map_letter_a))

    End If
    
    If ZonaSeguraAlpha > 0 Or ZonaSeguraDir <> 0 Then
        ZonaSeguraAlpha = ZonaSeguraAlpha - timerTicksPerFrame * 3.5 * ZonaSeguraDir
        If ZonaSeguraAlpha < 0 Then
            ZonaSeguraAlpha = 0
            ZonaSeguraDir = -1
        End If
        If ZonaSeguraAlpha > 255 Then
            ZonaSeguraAlpha = 255
            ZonaSeguraDir = 0
        End If
        Dim textZona As String
        textZona = IIf(ZoneSegura = 1, "Has ingresado a una zona segura", "Has salido de una zona segura")
        Call Engine_Text_Render(textZona, 360 - Engine_Text_Width(textZona, False, 1) / 2, 0 + ConsoleDialog.console_height + 10 + 55, Color(), 1, False, , CInt(255 - ZonaSeguraAlpha))
    End If

    
    Exit Sub

RenderScreen_NombreMapa_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_RenderScreen.RenderScreen_NombreMapa", Erl)
    Resume Next
    
End Sub

