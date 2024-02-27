VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0

Private Type tLight

    RGBcolor As D3DCOLORVALUE
    active As Boolean
    map_X As Integer
    map_Y As Integer
    range As Byte

End Type
 
Private light_list() As tLight
Private NumLights    As Integer

Public Function Create_Light_To_Map(ByVal map_X As Integer, ByVal map_Y As Integer, Optional range As Byte = 3, Optional ByVal red As Byte = 255, Optional ByVal green = 255, Optional ByVal blue As Byte = 255)
    
    On Error GoTo Create_Light_To_Map_Err
    
    NumLights = NumLights + 1
   
    ReDim Preserve light_list(1 To NumLights) As tLight
   
    'Le damos color
    light_list(NumLights).RGBcolor.r = red
    light_list(NumLights).RGBcolor.g = green
    light_list(NumLights).RGBcolor.b = blue
   
    'Alpha (Si borras esto RE KB!!)
    light_list(NumLights).RGBcolor.A = 255
   
    'Un rango determinado
    light_list(NumLights).range = range
   
    'La activamos para que se pueda renderizar
    light_list(NumLights).active = True
   
    'Pos en el mapa
    light_list(NumLights).map_X = map_X
    light_list(NumLights).map_Y = map_Y
   
    'LightA.LightRenderAll

    
    Exit Function

Create_Light_To_Map_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLight.Create_Light_To_Map", Erl)
    Resume Next
    
End Function

Public Function Delete_Light_To_Map(ByVal X As Integer, ByVal Y As Integer)
    
    On Error GoTo Delete_Light_To_Map_Err
    
   
    Dim i As Long
   
    For i = 1 To NumLights

        If light_list(i).map_X = X And light_list(i).map_Y = Y Then
            Delete_Light_To_Index i
            Exit Function

        End If

    Next i
 
    
    Exit Function

Delete_Light_To_Map_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLight.Delete_Light_To_Map", Erl)
    Resume Next
    
End Function

Public Function Delete_Light_To_Index(ByVal light_index As Integer)

    On Error Resume Next

    Dim min_x As Long
    Dim min_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim Ya    As Integer
    Dim Xa    As Integer
 
    light_list(light_index).active = False
    NumLights = NumLights - 1
 
    min_x = light_list(light_index).map_X - light_list(light_index).range
    max_x = light_list(light_index).map_X + light_list(light_index).range
    min_y = light_list(light_index).map_Y - light_list(light_index).range
    max_y = light_list(light_index).map_Y + light_list(light_index).range
    
    If min_x <= 0 Then min_x = 1
    If max_x > MapSize.Width Then min_x = MapSize.Width
        
    If min_y <= 0 Then min_y = 1
        
    If max_y > MapSize.Height Then max_y = MapSize.Height
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                MapData(Xa, Ya).light_value(0) = 0
                MapData(Xa, Ya).light_value(1) = 0
                MapData(Xa, Ya).light_value(2) = 0
                MapData(Xa, Ya).light_value(3) = 0

            End If

        Next Xa
    Next Ya
    
    LightA.LightRenderAll
   
End Function

Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    
    On Error GoTo LightCalculate_Err
    
    Dim XDist        As Single
    Dim YDist        As Single
    Dim VertexDist   As Single
    Dim pRadio       As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 2
   
    XDist = LightX + 1 - XCoord
    YDist = LightY + 1 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.r), Round(CurrentColor.g), Round(CurrentColor.b))
        '        If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight

    End If

    
    Exit Function

LightCalculate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLight.LightCalculate", Erl)
    Resume Next
    
End Function
 
Private Sub LightRender(ByVal light_index As Integer)
 
    On Error Resume Next

    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
   
    Dim min_x        As Long
    Dim min_y        As Long
    Dim max_x        As Long
    Dim max_y        As Long
    Dim color        As Long
    Dim Ya           As Integer
    Dim Xa           As Integer
   
    Dim TileLight    As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor   As D3DCOLORVALUE
   
    Dim XCoord       As Long
    Dim YCoord       As Long
    
    Dim Ambiental As Long
    Ambiental = ColorAmb
    If Ambiental = 0 Then Ambiental = -1
                             
    Dim r, g, b As Byte
    b = (Ambiental And 16711680) / 65536
    g = (Ambiental And 65280) / 256
    r = Ambiental And 255
   
    AmbientColor.r = r
    AmbientColor.g = g
    AmbientColor.b = b
    AmbientColor.A = 255
   
    LightColor = light_list(light_index).RGBcolor
       
    min_x = light_list(light_index).map_X - light_list(light_index).range
    max_x = light_list(light_index).map_X + light_list(light_index).range
    min_y = light_list(light_index).map_Y - light_list(light_index).range
    max_y = light_list(light_index).map_Y + light_list(light_index).range
       
    If min_x <= 0 Then min_x = 1
    If max_x > MapSize.Width Then min_x = MapSize.Width
        
    If min_y <= 0 Then min_y = 1
        
    If max_y > MapSize.Height Then max_y = MapSize.Height

    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 2
                YCoord = Ya * 2
                MapData(Xa, Ya).light_value(1) = LightCalculate(light_list(light_index).range, light_list(light_index).map_X * 2, light_list(light_index).map_Y * 2, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
 
                XCoord = Xa * 2 + 2
                YCoord = Ya * 2
                MapData(Xa, Ya).light_value(3) = LightCalculate(light_list(light_index).range, light_list(light_index).map_X * 2, light_list(light_index).map_Y * 2, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                XCoord = Xa * 2
                YCoord = Ya * 2 + 2
                MapData(Xa, Ya).light_value(0) = LightCalculate(light_list(light_index).range, light_list(light_index).map_X * 2, light_list(light_index).map_Y * 2, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
   
                XCoord = Xa * 2 + 2
                YCoord = Ya * 2 + 2
                MapData(Xa, Ya).light_value(2) = LightCalculate(light_list(light_index).range, light_list(light_index).map_X * 2, light_list(light_index).map_Y * 2, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
               
            End If

        Next Xa
    Next Ya

End Sub

Public Sub LightRenderAll()
    
    On Error GoTo LightRenderAll_Err
    
    Dim i As Long
   
    For i = 1 To NumLights
        LightRender i
            
    Next i

    
    Exit Sub

LightRenderAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLight.LightRenderAll", Erl)
    Resume Next
    
End Sub

Public Function Delete_All_LigthRound()
    
    On Error GoTo Delete_All_LigthRound_Err
    
    Dim i As Long
   
    For i = 1 To NumLights
        Delete_Light_To_Index i
            
    Next i

    
    Exit Function

Delete_All_LigthRound_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsLight.Delete_All_LigthRound", Erl)
    Resume Next
    
End Function

