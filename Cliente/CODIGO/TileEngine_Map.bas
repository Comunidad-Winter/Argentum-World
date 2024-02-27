Attribute VB_Name = "TileEngine_Map"
Option Explicit

Sub SwitchMap(ByVal map As Integer)
    
    On Error GoTo SwitchMap_Err
    
    If map = 1 Then Exit Sub
    
    Call LoadMap(map)
    
    Exit Sub

SwitchMap_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.SwitchMap", Erl)
    Resume Next
    
End Sub

Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo HayAgua_Err
    

    With MapData(rrX(x), rrY(y)).Graphic(1)
            HayAgua = (.GrhIndex >= 1505 And .GrhIndex <= 1520) Or _
                        (.GrhIndex >= 124 And .GrhIndex <= 139) Or _
                        (.GrhIndex >= 24223 And .GrhIndex <= 24238) Or _
                        (.GrhIndex >= 24303 And .GrhIndex <= 24318) Or _
                        (.GrhIndex >= 468 And .GrhIndex <= 483) Or _
                        (.GrhIndex >= 44668 And .GrhIndex <= 44683) Or _
                        (.GrhIndex >= 24143 And .GrhIndex <= 24158) Or _
                        (.GrhIndex >= 12628 And .GrhIndex <= 12643) Or _
                        (.GrhIndex >= 2948 And .GrhIndex <= 2963)
    End With

    
    Exit Function

HayAgua_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.HayAgua", Erl)
    Resume Next
    
End Function

Function HayLava(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo HayLava_Err
    

    With MapData(rrX(x), rrY(y)).Graphic(1)
        HayLava = .GrhIndex >= 57400 And .GrhIndex <= 57415
    End With

    
    Exit Function

HayLava_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.HayLava", Erl)
    Resume Next
    
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
    
    On Error GoTo EsArbol_Err
    
    EsArbol = GrhIndex = 304 Or GrhIndex = 305 Or GrhIndex = 641 Or GrhIndex = 643 Or GrhIndex = 644 Or GrhIndex = 647 Or GrhIndex = 735 Or GrhIndex = 1121 Or GrhIndex = 1126 Or GrhIndex = 2931 Or _
              GrhIndex = 12309 Or GrhIndex = 12310 Or GrhIndex = 16833 Or GrhIndex = 16834 Or GrhIndex = 7020 Or GrhIndex = 11903 Or GrhIndex = 11904 Or _
              GrhIndex = 11905 Or GrhIndex = 11906 Or GrhIndex = 12160 Or GrhIndex = 15698 Or GrhIndex = 14504 Or GrhIndex = 15697 Or _
              (GrhIndex >= 12581 And GrhIndex <= 12586) Or (GrhIndex >= 12164 And GrhIndex <= 12179) Or _
              (GrhIndex >= 14950 And GrhIndex <= 14965) Or (GrhIndex >= 14967 And GrhIndex <= 14980) Or (GrhIndex >= 14982 And GrhIndex <= 14988) Or _
              (GrhIndex >= 26075 And GrhIndex <= 26081) Or GrhIndex = 26192 Or (GrhIndex >= 32142 And GrhIndex <= 32162) Or (GrhIndex >= 32343 And GrhIndex <= 32352) Or _
              (GrhIndex >= 55626 And GrhIndex <= 55640) Or GrhIndex = 55642 Or _
              (GrhIndex >= 50985 And GrhIndex <= 50991) Or (GrhIndex >= 2547 And GrhIndex <= 2549) Or (GrhIndex >= 6597 And GrhIndex <= 6598) Or (GrhIndex >= 15108 And GrhIndex <= 15110) Or GrhIndex = 11904 Or GrhIndex = 11905 Or GrhIndex = 11906 Or GrhIndex = 12160 Or _
              GrhIndex = 7220 Or GrhIndex = 50990 Or GrhIndex = 6597 Or GrhIndex = 6598 Or GrhIndex = 2548 Or GrhIndex = 2549 Or _
              GrhIndex = 463 Or GrhIndex = 1880 Or GrhIndex = 1878 Or GrhIndex = 9513 Or GrhIndex = 9514 Or GrhIndex = 9515 Or GrhIndex = 9518 Or GrhIndex = 9519 Or GrhIndex = 9520 Or GrhIndex = 9529 Or _
              GrhIndex = 55633 Or GrhIndex = 55627 Or GrhIndex = 15510 Or GrhIndex = 14775 Or GrhIndex = 14687
              
    
    Exit Function

EsArbol_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.EsArbol", Erl)
    Resume Next
    
End Function

Function AgregarSombra(ByVal GrhIndex As Long) As Boolean
    
    On Error GoTo AgregarSombra_Err
    
    AgregarSombra = GrhIndex = 5624 Or GrhIndex = 5625 Or GrhIndex = 5626 Or GrhIndex = 5627 Or GrhIndex = 51716

    
    Exit Function

AgregarSombra_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.AgregarSombra", Erl)
    Resume Next
    
End Function

Public Function EsObjetoFijo(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo EsObjetoFijo_Err
    
    Dim ObjIndex As Integer
    ObjIndex = MapData(rrX(x), rrY(y)).OBJInfo.ObjIndex
    
    Dim ObjType As eObjType
    ObjType = ObjData(ObjIndex).ObjType
    
    EsObjetoFijo = ObjType = eObjType.otForos Or ObjType = eObjType.otCarteles Or ObjType = eObjType.otArboles Or ObjType = eObjType.otYacimiento Or ObjType = eObjType.OtDecoraciones

    
    Exit Function

EsObjetoFijo_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.EsObjetoFijo", Erl)
    Resume Next
    
End Function

Public Function Letter_Set(ByVal grh_index As Long, ByVal text_string As String) As Boolean
    '*****************************************************************
    'Author: Augusto Jos� Rando
    '*****************************************************************
    
    On Error GoTo Letter_Set_Err
    
    letter_text = text_string
    letter_grh.GrhIndex = grh_index
    Letter_Set = True
    map_letter_fadestatus = 1

    
    Exit Function

Letter_Set_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.Letter_Set", Erl)
    Resume Next
    
End Function



Public Sub SetGlobalLight(ByVal base_light As Long)
    
    On Error GoTo SetGlobalLight_Err
    
    Call Long_2_RGBA(global_light, base_light)
    global_light.A = 255
    light_transition = 1#
    
    Exit Sub

SetGlobalLight_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.SetGlobalLight", Erl)
    Resume Next
    
End Sub

Public Function Map_FX_Group_Next_Open(ByVal x As Integer, ByVal y As Integer) As Integer

    '*****************************************************************
    'Author: Augusto Jos� Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If MapData(rrX(x), rrY(y)).FxCount = 0 Then
        MapData(rrX(x), rrY(y)).FxCount = 1
        ReDim MapData(rrX(x), rrY(y)).FxList(1 To 1)
        Map_FX_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until MapData(rrX(x), rrY(y)).FxList(loopc).FxIndex = 0

        If loopc = MapData(rrX(x), rrY(y)).FxCount Then
            Map_FX_Group_Next_Open = MapData(rrX(x), rrY(y)).FxCount + 1
            MapData(rrX(x), rrY(y)).FxCount = Map_FX_Group_Next_Open
            ReDim Preserve MapData(rrX(x), rrY(y)).FxList(1 To Map_FX_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop

    Map_FX_Group_Next_Open = loopc
    Exit Function

ErrorHandler:
    MapData(rrX(x), rrY(y)).FxCount = 1
    ReDim MapData(rrX(x), rrY(y)).FxList(1 To 1)
    Map_FX_Group_Next_Open = 1

End Function

Public Sub Draw_Sombra(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Integer = 1, Optional ByVal map_y As Integer = 1, Optional ByVal Angle As Single)
    
    On Error GoTo Draw_Sombra_Err

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
    End If

    If GrhData(grh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
    End If

    Call Batch_Textured_Box_Shadow(x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, GrhData(CurrentGrhIndex).FileNum, MapData(rrX(map_x), rrY(map_y)).light_value)
    
    Exit Sub

Draw_Sombra_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.Draw_Sombra", Erl)
    Resume Next
    
End Sub

Sub Engine_Weather_UpdateFog()
    
    On Error GoTo Engine_Weather_UpdateFog_Err
    

    '*****************************************************************
    'Update the fog effects
    '*****************************************************************
    Dim TempGrh     As grh

    Dim i           As Long

    Dim x           As Long

    Dim y           As Long

    Dim cc(3)       As RGBA

    Dim ElapsedTime As Single

    ElapsedTime = Engine_ElapsedTime

    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop

    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop

    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop

    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)

    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop

    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop

    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop

    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop

    Call InitGrh(TempGrh, 32014)

    x = 2
    y = -1

    Call RGBAList(cc, 255, 255, 255, AlphaNiebla)

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX2, (y * 512) + WeatherFogY2, 0, 0, cc()
        x = x + 1

        If x > (1 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i
            
    'Render fog 1
    TempGrh.GrhIndex = 32015
    x = 0
    y = 0

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 0, 0, cc()
        x = x + 1

        If x > (2 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i

    
    Exit Sub

Engine_Weather_UpdateFog_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.Engine_Weather_UpdateFog", Erl)
    Resume Next
    
End Sub

Sub MapUpdateGlobalLight()
    
    On Error GoTo MapUpdateGlobalLight_Err
    

    Dim x As Integer, y As Integer
    
    ' Reseteamos toda la luz del mapa
    For y = 1 To MapSize.Height
        For x = 1 To MapSize.Width
            With MapData(rrX(x), rrY(y))
            
                .light_value(0) = global_light
                .light_value(1) = global_light
                .light_value(2) = global_light
                .light_value(3) = global_light
                
            End With
        Next x
    Next y
    
    Exit Sub

MapUpdateGlobalLight_Err:
    Call RegistrarError(err.Number, err.Description, "TileEngine_Map.MapUpdateGlobalLight", Erl)
    Resume Next
    
End Sub

Sub MapUpdateGlobalLightRender()
    
    On Error GoTo MapUpdateGlobalLight_Err
    

    Dim x As Integer, y As Integer
    Dim MinX As Long, MinY As Long, MaxX As Long, MaxY As Long
    MinX = 1
    MinY = 1
    MaxX = 100
    MaxY = 100
    
    ' Reseteamos toda la luz del mapa
    For y = MinY To MaxY
        For x = MinX To MaxX
            With MapData(rrX(x), rrY(y))
            
                .light_value(0) = global_light
                .light_value(1) = global_light
                .light_value(2) = global_light
                .light_value(3) = global_light
                
            End With
        Next x
    Next y
    
    Call LucesRedondas.LightRenderAll(MinX, MinY, MaxX, MaxY) '(MinX, MinY, MaxX, MaxY)
    Call LucesCuadradas.Light_Render_All(MinX, MinY, MaxX, MaxY)  '(MinX, MinY, MaxX, MaxY)
        
    Exit Sub

MapUpdateGlobalLight_Err:
   ' Call RegistrarError(Err.Number, Err.Description, "TileEngine_Map.MapUpdateGlobalLightRender", Erl)
    Resume Next
    
End Sub
