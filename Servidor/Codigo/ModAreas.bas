Attribute VB_Name = "ModAreas"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit
 
'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type t_ConnGroup

    CountEntrys As Integer
    OptValue As Long
    UserEntrys() As Integer

End Type

Public Type t_ConnArea
    Area() As t_ConnGroup
End Type

Public Const MAX_MAP_X As Integer = 1432
Public Const MAX_MAP_Y As Integer = 1780

Private AreasInfo(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As Long


Public Const USER_NUEVO               As Byte = 255

Public Const AREA_DIM                As Byte = 13
 
'Cuidado:
' ¡¡¡LAS AREAS ESTÃN HARDCODEADAS!
Private CurDay                        As Byte

Private CurHour                       As Byte
 
Public ConnGroups()                   As t_ConnArea

Public Sub InitAreasPre()
        Dim X As Integer
        Dim Y As Integer

106     For X = 1 To MAX_MAP_X
108         For Y = 1 To MAX_MAP_Y
110             AreasInfo(X, Y) = ((X - 1) \ AREA_DIM + 1) + ((MAX_MAP_X - 1) \ AREA_DIM + 1) * ((Y - 1) \ AREA_DIM)
112         Next Y
114     Next X
End Sub

Public Sub InitAreas()
        
        On Error GoTo InitAreas_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC As Integer

        Dim LoopX As Integer
        
        Dim X As Integer
        Dim Y As Integer


        'Setup AutoOptimizacion de areas
116     CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
118     CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece

120     ReDim ConnGroups(1 To NumMaps) As t_ConnArea
    
122     For LoopC = 1 To NumMaps
'124         ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))

            Mapinfo(LoopC).AreaX = ((Mapinfo(LoopC).Width - 1) \ AREA_DIM) + 1
            Mapinfo(LoopC).AreaY = ((Mapinfo(LoopC).Height - 1) \ AREA_DIM) + 1

            ReDim ConnGroups(LoopC).Area(Mapinfo(LoopC).AreaX, Mapinfo(LoopC).AreaY) As t_ConnGroup
            
            For X = 0 To Mapinfo(LoopC).AreaX
                For Y = 0 To Mapinfo(LoopC).AreaY
126                 If ConnGroups(LoopC).Area(X, Y).OptValue = 0 Then ConnGroups(LoopC).Area(X, Y).OptValue = 1
128                 ReDim ConnGroups(LoopC).Area(X, Y).UserEntrys(1 To ConnGroups(LoopC).Area(X, Y).OptValue) As Integer
                Next Y
            Next X
            

130     Next LoopC
        
        Exit Sub

InitAreas_Err:
132     Call TraceError(Err.Number, Err.Description, "ModAreas.InitAreas", Erl)

        
End Sub
 
Public Sub AreasOptimizacion()
        
        On Error GoTo AreasOptimizacion_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
        '**************************************************************
        Dim LoopC      As Long

        Dim tCurDay    As Byte

        Dim tCurHour   As Byte

        Dim EntryValue As Long
        
        Dim X As Integer
        Dim Y As Integer
        
    
100     If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        
102         tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
104         tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        
106         For LoopC = 1 To NumMaps
108             'EntryValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
110             'Call WriteVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
            
112             'ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, tCurDay & "-" & tCurHour))

114             'If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
116
                For X = 1 To Mapinfo(LoopC).AreaX
                    For Y = 1 To Mapinfo(LoopC).AreaY
                        If ConnGroups(LoopC).Area(X, Y).CountEntrys > 0 And ConnGroups(LoopC).Area(X, Y).OptValue > ConnGroups(LoopC).Area(X, Y).CountEntrys Then ReDim Preserve ConnGroups(LoopC).Area(X, Y).UserEntrys(1 To ConnGroups(LoopC).Area(X, Y).CountEntrys) As Integer
                    Next Y
                Next X
                
118         Next LoopC
        
120         CurDay = tCurDay
122         CurHour = tCurHour

        End If

        
        Exit Sub

AreasOptimizacion_Err:
124     Call TraceError(Err.Number, Err.Description, "ModAreas.AreasOptimizacion", Erl)

        
End Sub
 
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, ByVal appear As Byte, Optional ByVal Muerto As Byte = 0)

        On Error GoTo CheckUpdateNeededUser_Err

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Es la función clave del sistema de areas... Es llamada al mover un user
        '**************************************************************
100     If (UserList(UserIndex).AreaId = AreasInfo(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) And Head <> USER_NUEVO) And Muerto = 0 Then Exit Sub
    
        Dim MinX    As Integer, MaxX As Integer, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt As Long, Map As Long

102     With UserList(UserIndex)
104        MinX = ((.Pos.X - 1) \ AREA_DIM) * AREA_DIM + 1
106        MinY = ((.Pos.Y - 1) \ AREA_DIM) * AREA_DIM + 1

           Map = UserList(UserIndex).Pos.Map
        
108         If Head = e_Heading.NORTH Then
                
                MaxX = MinX + AREA_DIM * 2 - 1
                MinX = MinX - AREA_DIM
                
                MaxY = MinY - 1
                MinY = MinY - AREA_DIM
                
        
120         ElseIf Head = e_Heading.SOUTH Then


                MaxX = MinX + AREA_DIM * 2 - 1
                MinX = MinX - AREA_DIM
                
                
                MinY = MinY + AREA_DIM
                MaxY = MinY + AREA_DIM - 1

132         ElseIf Head = e_Heading.WEST Then


                MaxY = MinY + AREA_DIM * 2 - 1
                MinY = MinY - AREA_DIM
                
                MaxX = MinX - 1
                MinX = MinX - AREA_DIM

        
144         ElseIf Head = e_Heading.EAST Then


                MaxY = MinY + AREA_DIM * 2 - 1
                MinY = MinY - AREA_DIM
                
                
                MinX = MinX + AREA_DIM
                MaxX = MinX + AREA_DIM - 1

           
156         ElseIf Head = USER_NUEVO Or Head = 5 Then
                'Esto pasa por cuando cambiamos de mapa o logeamos...
                
                MaxX = MinX + AREA_DIM * 2 - 1
                MaxY = MinY + AREA_DIM * 2 - 1
158             MinY = MinY - AREA_DIM
                MinX = MinX - AREA_DIM
            End If
        
170         If MinY < 1 Then MinY = 1
172         If MinX < 1 Then MinX = 1
174         If MaxY > Mapinfo(Map).Height Then MaxY = Mapinfo(Map).Height
176         If MaxX > Mapinfo(Map).Width Then MaxX = Mapinfo(Map).Width
  
       
            'Esto es para ke el cliente elimine lo "fuera de area..."
180         Call WriteAreaChanged(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, Head)
            If UserList(UserIndex).flags.GMMeSigue > 0 Then
                Call WriteAreaChanged(UserList(UserIndex).flags.GMMeSigue, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, Head)
                Call WriteSendFollowingCharindex(UserList(UserIndex).flags.GMMeSigue, UserList(UserIndex).Char.charindex)
            End If
       
            'Actualizamos!
182         For X = MinX To MaxX
184             For Y = MinY To MaxY
               
                    '<<< User >>>
186                 If MapData(Map).Tile(X, Y).UserIndex Then
                   
188                     TempInt = MapData(Map).Tile(X, Y).UserIndex

190                     If UserIndex <> TempInt Then
                            'NOTIFICO AL USUARIO QUE ESTABA EN EL AREA
                            Call NotifyUser(TempInt, UserIndex)
                            
                            'NOTIFICO AL USUARIO QUE LLEGA AL AREA
                            Call NotifyUser(UserIndex, TempInt)

208                     ElseIf Head = USER_NUEVO Then
210                         'Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y, appear)
                        End If

                    End If
               
                    '<<< Npc >>>
                    If MapData(Map).Tile(X, Y).NpcIndex Then
                        Call MakeNPCChar(False, UserIndex, MapData(Map).Tile(X, Y).NpcIndex, Map, X, Y)
                    End If
212
                 
                    '<<< Item >>>
216                 If MapData(Map).Tile(X, Y).ObjInfo.ObjIndex Then
218                     TempInt = MapData(Map).Tile(X, Y).ObjInfo.ObjIndex

220                     If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
222                         Call WriteObjectCreate(UserIndex, TempInt, MapData(Map).Tile(X, Y).ObjInfo.Amount, X, Y)
                           ' If tmpGM > 0 Then Call WriteObjectCreate(tmpGM, TempInt, MapData(map).Tile(X, Y).ObjInfo.amount, X, Y)
                            
224                         If ObjData(TempInt).OBJType = e_OBJType.otPuertas And InMapBounds(Map, X, Y) Then
226                             Call MostrarBloqueosPuerta(False, UserIndex, X, Y)
                                'If tmpGM > 0 Then Call MostrarBloqueosPuerta(False, tmpGM, X, Y)
                            End If
                            
                        End If

                    End If

                    ' Bloqueo GM
228                 If (MapData(Map).Tile(X, Y).Blocked And e_Block.GM) <> 0 Then
230                     Call Bloquear(False, UserIndex, X, Y, e_Block.ALL_SIDES)
                    End If
                    
232             Next Y
234         Next X
            
            If Head <> USER_NUEVO Then
            
                Dim AreaX As Integer
                Dim AreaY As Integer
                
                If .AreaId > 0 Then
                
                    AreaY = (.AreaId \ ((MAX_MAP_X - 1) \ AREA_DIM + 1)) * AREA_DIM + 1
                    AreaX = (.AreaId Mod ((MAX_MAP_X - 1) \ AREA_DIM + 1)) * AREA_DIM
                
248
    
                    Call QuitarUser(UserIndex, Map, AreaX, AreaY)
                
                End If
                .AreaMap = .Pos.Map
                .AreaId = AreasInfo(.Pos.X, .Pos.Y)
                Call AgregarUser(UserIndex, Map, .Pos.X, .Pos.Y)
            Else
                .AreaMap = .Pos.Map
                .AreaId = AreasInfo(.Pos.X, .Pos.Y)

            End If
                
                
                'Es un gm que está siguiendo a un usuario
                If .flags.SigueUsuario > 0 Then
                 
                 .AreaMap = UserList(.flags.SigueUsuario).AreaMap
                 .AreaId = UserList(.flags.SigueUsuario).AreaId
                
                End If
            
            
            'Es un usuario que está siendo seguido
            If .flags.GMMeSigue > 0 Then
             
                UserList(.flags.GMMeSigue).AreaMap = .Pos.Map
                UserList(.flags.GMMeSigue).AreaId = .AreaId
            
            End If
        End With

        
        Exit Sub

CheckUpdateNeededUser_Err:
250     Call TraceError(Err.Number, Err.Description, "ModAreas.CheckUpdateNeededUser", Erl)

        
End Sub

Private Sub NotifyUser(ByVal UserNotificado As Integer, ByVal UserIngresante As Integer)

    Dim sendChar As Boolean

    sendChar = True

    With UserList(UserNotificado)
        If UserList(UserIngresante).flags.AdminInvisible = 1 Then
            If Not EsGM(UserNotificado) Or CompararPrivilegios(.flags.Privilegios, UserList(UserIngresante).flags.Privilegios) < 0 Then
                sendChar = False
            End If
         ElseIf UserList(UserNotificado).flags.Muerto = 1 And Zona(.ZonaId).Segura = 0 And (UserList(UserNotificado).GuildIndex = 0 Or UserList(UserNotificado).GuildIndex <> UserList(UserIngresante).GuildIndex Or modGuilds.NivelDeClan(UserList(UserIngresante).GuildIndex) < 6) Then
            sendChar = False
        End If
            

        If sendChar Then
            Call MakeUserChar(UserNotificado, UserIngresante, UserList(UserIngresante).Pos.Map, UserList(UserIngresante).Pos.X, UserList(UserIngresante).Pos.Y, 0)
            If UserList(UserIngresante).flags.invisible Or UserList(UserIngresante).flags.Oculto Then
                Call WriteSetInvisible(UserNotificado, UserList(UserIngresante).Char.charindex, True)
            End If
        End If
    End With

End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
        
        On Error GoTo CheckUpdateNeededNpc_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        ' Se llama cuando se mueve un Npc
        '**************************************************************
100     If NpcList(NpcIndex).AreaId = AreasInfo(NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y) Then Exit Sub
    
        Dim MinX    As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer

        Dim TempInt As Long

        Dim appear  As Byte

        Dim Map As Integer
102     appear = 0
    
104     With NpcList(NpcIndex)

105        MinX = ((.Pos.X - 1) \ AREA_DIM) * AREA_DIM + 1
106        MinY = ((.Pos.Y - 1) \ AREA_DIM) * AREA_DIM + 1

           Map = NpcList(NpcIndex).Pos.Map
        
108         If Head = e_Heading.NORTH Then
                
                MaxX = MinX + AREA_DIM * 2 - 1
                MinX = MinX - AREA_DIM
                
                MaxY = MinY - 1
                MinY = MinY - AREA_DIM
                
        
120         ElseIf Head = e_Heading.SOUTH Then


                MaxX = MinX + AREA_DIM * 2 - 1
                MinX = MinX - AREA_DIM
                
                
                MinY = MinY + AREA_DIM
                MaxY = MinY + AREA_DIM - 1

132         ElseIf Head = e_Heading.WEST Then


                MaxY = MinY + AREA_DIM * 2 - 1
                MinY = MinY - AREA_DIM
                
                MaxX = MinX - 1
                MinX = MinX - AREA_DIM

        
144         ElseIf Head = e_Heading.EAST Then


                MaxY = MinY + AREA_DIM * 2 - 1
                MinY = MinY - AREA_DIM
                
                
                MinX = MinX + AREA_DIM
                MaxX = MinX + AREA_DIM - 1

           
156         ElseIf Head = USER_NUEVO Then

                MaxX = MinX + AREA_DIM * 2 - 1
                MaxY = MinY + AREA_DIM * 2 - 1
158             MinY = MinY - AREA_DIM
                MinX = MinX - AREA_DIM
            End If

        
170         If MinY < 1 Then MinY = 1
172         If MinX < 1 Then MinX = 1
174         If MaxY > Mapinfo(Map).Height Then MaxY = Mapinfo(Map).Height
176         If MaxX > Mapinfo(Map).Width Then MaxX = Mapinfo(Map).Width
        
            'Actualizamos!
182         If Mapinfo(.Pos.Map).NumUsers <> 0 Then

184             For X = MinX To MaxX
186                 For Y = MinY To MaxY
                        
188                     If MapData(.Pos.Map).Tile(X, Y).UserIndex Then Call MakeNPCChar(False, MapData(.Pos.Map).Tile(X, Y).UserIndex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)

190                 Next Y
192             Next X

            End If
            
206         .AreaId = AreasInfo(.Pos.X, .Pos.Y)

        End With

        
        Exit Sub

CheckUpdateNeededNpc_Err:
208     Call TraceError(Err.Number, Err.Description, "ModAreas.CheckUpdateNeededNpc", Erl)

        
End Sub
 
Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo QuitarUser_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim TempVal As Long

        Dim LoopC   As Long
        
        Dim AreaX As Integer
        Dim AreaY As Integer
        
        AreaX = ((X - 1) \ AREA_DIM) + 1
        AreaY = ((Y - 1) \ AREA_DIM) + 1
   
        'Search for the user
100     For LoopC = 1 To ConnGroups(Map).Area(AreaX, AreaY).CountEntrys

102         If ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(LoopC) = UserIndex Then Exit For
104     Next LoopC
   
        'Char not found
106     If LoopC > ConnGroups(Map).Area(AreaX, AreaY).CountEntrys Then Exit Sub
   
        'Remove from old map
108     ConnGroups(Map).Area(AreaX, AreaY).CountEntrys = ConnGroups(Map).Area(AreaX, AreaY).CountEntrys - 1
110     TempVal = ConnGroups(Map).Area(AreaX, AreaY).CountEntrys
   
        'Move list back
112     For LoopC = LoopC To TempVal
114         ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(LoopC) = ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(LoopC + 1)
116     Next LoopC
   
118     If TempVal > ConnGroups(Map).Area(AreaX, AreaY).OptValue Then 'Nescesito Redim?
120         ReDim Preserve ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(1 To TempVal) As Integer

        End If
        
        Exit Sub

QuitarUser_Err:
122     Call TraceError(Err.Number, Err.Description, "ModAreas.QuitarUser", Erl)

        
End Sub
 
Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal appear As Byte = 0)
        
        On Error GoTo AgregarUser_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: 04/01/2007
        'Modified by Juan Martín Sotuyo Dodero (Maraxus)
        '   - Now the method checks for repetead users instead of trusting parameters.
        '   - If the character is new to the map, update it
        '**************************************************************
        Dim TempVal As Long

        Dim EsNuevo As Boolean

        Dim i       As Long
        Dim AreaX As Integer
        Dim AreaY As Integer
        
   
100     If Not MapaValido(Map) Then Exit Sub

        AreaX = ((X - 1) \ AREA_DIM) + 1
        AreaY = ((Y - 1) \ AREA_DIM) + 1
   
102     EsNuevo = True
   
        'Prevent adding repeated users
104     For i = 1 To ConnGroups(Map).Area(AreaX, AreaY).CountEntrys

106         If ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(i) = UserIndex Then
108             EsNuevo = False
                Exit For

            End If

110     Next i
   
112     If EsNuevo Then
            'Update map and connection groups data
114         ConnGroups(Map).Area(AreaX, AreaY).CountEntrys = ConnGroups(Map).Area(AreaX, AreaY).CountEntrys + 1
116         TempVal = ConnGroups(Map).Area(AreaX, AreaY).CountEntrys
       
118         If TempVal > ConnGroups(Map).Area(AreaX, AreaY).OptValue Then 'Nescesito Redim
120             ReDim Preserve ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(1 To TempVal) As Integer

            End If
       
122         ConnGroups(Map).Area(AreaX, AreaY).UserEntrys(TempVal) = UserIndex

        End If
        
        Exit Sub

AgregarUser_Err:
136     Call TraceError(Err.Number, Err.Description, "ModAreas.AgregarUser", Erl)

        
End Sub
 
Public Sub AgregarNpc(ByVal NpcIndex As Integer)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        
        On Error GoTo AgregarNpc_Err
        
100     NpcList(NpcIndex).AreaId = 0
      
110     Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)

        
        Exit Sub

AgregarNpc_Err:
112     Call TraceError(Err.Number, Err.Description, "ModAreas.AgregarNpc", Erl)

        
End Sub


Public Function GetUsersArea(ByVal UserIndex As Integer, ByRef Total As Integer) As Integer()

Dim AreaX As Integer
Dim AreaY As Integer

Dim X As Integer
Dim Y As Integer
Dim Map As Integer
Dim index As Integer
Dim LoopC As Integer

Dim result() As Integer

Map = UserList(UserIndex).Pos.Map
AreaX = ((UserList(UserIndex).Pos.X - 1) \ AREA_DIM) + 1
AreaY = ((UserList(UserIndex).Pos.Y - 1) \ AREA_DIM) + 1

For X = AreaX - 1 To AreaX + 1
    For Y = AreaY - 1 To AreaY + 1
        If X > 0 And X <= Mapinfo(Map).AreaX And Y > 0 And Y <= Mapinfo(Map).AreaY Then
            Total = Total + ConnGroups(Map).Area(X, Y).CountEntrys
        End If
    Next Y
Next X
If Total > 0 Then
    ReDim result(1 To Total)
    index = 1
    For X = AreaX - 1 To AreaX + 1
        For Y = AreaY - 1 To AreaY + 1
            If X > 0 And X <= Mapinfo(Map).AreaX And Y > 0 And Y <= Mapinfo(Map).AreaY Then
                For LoopC = 1 To ConnGroups(Map).Area(X, Y).CountEntrys
                    result(index) = ConnGroups(Map).Area(X, Y).UserEntrys(LoopC)
                    index = index + 1
                Next LoopC
            End If
        Next Y
    Next X
End If
GetUsersArea = result


End Function


Public Function GetUsersNpcArea(ByVal NpcIndex As Integer, ByRef Total As Integer) As Integer()

Dim AreaX As Integer
Dim AreaY As Integer

Dim X As Integer
Dim Y As Integer
Dim Map As Integer
Dim index As Integer
Dim LoopC As Integer


Dim result() As Integer

Map = NpcList(NpcIndex).Pos.Map
AreaX = ((NpcList(NpcIndex).Pos.X - 1) \ AREA_DIM) + 1
AreaY = ((NpcList(NpcIndex).Pos.Y - 1) \ AREA_DIM) + 1

For X = AreaX - 1 To AreaX + 1
    For Y = AreaY - 1 To AreaY + 1
        If X > 0 And X <= Mapinfo(Map).AreaX And Y > 0 And Y <= Mapinfo(Map).AreaY Then
            Total = Total + ConnGroups(Map).Area(X, Y).CountEntrys
        End If
    Next Y
Next X
If Total > 0 Then
ReDim result(1 To Total)
    index = 1
    For X = AreaX - 1 To AreaX + 1
        For Y = AreaY - 1 To AreaY + 1
            If X > 0 And X <= Mapinfo(Map).AreaX And Y > 0 And Y <= Mapinfo(Map).AreaY Then
                For LoopC = 1 To ConnGroups(Map).Area(X, Y).CountEntrys
                    result(index) = ConnGroups(Map).Area(X, Y).UserEntrys(LoopC)
                    index = index + 1
                Next LoopC
            End If
        Next Y
    Next X
End If
GetUsersNpcArea = result


End Function

Public Function GetUsersPosArea(ByVal Map As Integer, ByVal pX As Integer, ByVal pY As Integer, ByRef Total As Integer) As Integer()

Dim AreaX As Integer
Dim AreaY As Integer

Dim X As Integer
Dim Y As Integer
Dim index As Integer
Dim LoopC As Integer

Dim result() As Integer

AreaX = ((pX - 1) \ AREA_DIM) + 1
AreaY = ((pY - 1) \ AREA_DIM) + 1

For X = AreaX - 1 To AreaX + 1
    For Y = AreaY - 1 To AreaY + 1
        If X > 0 And X <= Mapinfo(Map).AreaX And Y > 0 And Y <= Mapinfo(Map).AreaY Then
            Total = Total + ConnGroups(Map).Area(X, Y).CountEntrys
        End If
    Next Y
Next X
If Total > 0 Then
    ReDim result(1 To Total)
    index = 1
    For X = AreaX - 1 To AreaX + 1
        For Y = AreaY - 1 To AreaY + 1
            If X > 0 And X <= Mapinfo(Map).AreaX And Y > 0 And Y <= Mapinfo(Map).AreaY Then
                For LoopC = 1 To ConnGroups(Map).Area(X, Y).CountEntrys
                    result(index) = ConnGroups(Map).Area(X, Y).UserEntrys(LoopC)
                    index = index + 1
                Next LoopC
            End If
        Next Y
    Next X
End If
GetUsersPosArea = result


End Function


Sub getUsersAll()
Dim X As Integer
Dim Y As Integer
Dim Areas As String
Dim LoopC As Integer
For LoopC = 1 To NumMaps

    For X = 1 To Mapinfo(LoopC).AreaX
        For Y = 1 To Mapinfo(LoopC).AreaY
            If ConnGroups(LoopC).Area(X, Y).CountEntrys Then
                Areas = Areas & ",(" & X & "," & Y & ")"
            End If
        Next Y
    Next X
Next LoopC
Debug.Print Areas
End Sub
