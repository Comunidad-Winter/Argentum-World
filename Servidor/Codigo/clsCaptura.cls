VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private participantes As Collection

Private cantidad_rondas As Byte
Private nivel_minimo As Byte
Private max_participantes As Byte
Private tiempo_espera As Integer
Private tiempo_ronda As Integer
Private Precio As Long
Private ronda_actual As Byte
Private reintentos As Byte

Private Enum estadoCaptura
    Inscripcion = 1
    IniciaRonda = 2
    EnJuego = 3
    Finalizado = 4
End Enum

Private Estado As estadoCaptura

Private contadorBandera(1 To 2) As Integer

Private rondasGanadas(1 To 2) As Integer


Private colorEquipo(1 To 2) As String
Private mapaEquipo(1 To 2) As Integer
Private xEquipo(1 To 2) As Byte
Private yEquipo(1 To 2) As Byte
Private objEquipo(1 To 2) As Integer
Private Sub Class_Initialize()
    Set participantes = New Collection
    
    colorEquipo(1) = "azul"
    colorEquipo(2) = "rojo"
    
    mapaEquipo(1) = MAP_TEAM_1
    mapaEquipo(2) = MAP_TEAM_2
    
    xEquipo(1) = X_BANDERA_1
    xEquipo(2) = X_BANDERA_2
    yEquipo(1) = Y_BANDERA_1
    yEquipo(2) = Y_BANDERA_2
    
    objEquipo(1) = OBJ_CAPTURA_BANDERA_1
    objEquipo(2) = OBJ_CAPTURA_BANDERA_2
End Sub

Public Sub eliminarParticipante(ByVal UserIndex As Integer)

    Call tirarBandera(UserIndex)

    UserList(UserIndex).flags.jugando_captura = 0
    UserList(UserIndex).flags.jugando_captura_team = 0
    
    Dim p As clsCapturaParticipante
    
    Set p = participantes.Item(CStr(UserIndex))
    
    Call WarpUserChar(UserIndex, p.LastPosMap, p.LastPosx, p.LastPosy)
    
    Call participantes.Remove(CStr(UserIndex))
    
    
End Sub

Public Sub PasarSegundo()

    Dim participante As clsCapturaParticipante
    If Estado = Inscripcion Then
        tiempo_espera = tiempo_espera - 1
        If tiempo_espera Mod 30 = 0 And tiempo_espera > 5 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» En " & tiempo_espera & " segundos se cierran las inscripciones para el evento captura de bandera, escribe /PARTICIPARCAPTURA para anotarte.", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_espera <= 5 And tiempo_espera > 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Cerramos inscripciones en " & tiempo_espera & ".", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_espera = 0 Then
            If participantes.Count < max_participantes / 2 Then
                reintentos = reintentos + 1
                If reintentos >= 5 Then
                    Call finalizarCaptura(-2)
                Else
                    tiempo_espera = 60
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Se extendieron las inscripciones por falta de participantes.", e_FontTypeNames.FONTTYPE_GUILD))
                End If
            Else
                Call iniciarCaptura
            End If
        End If
    ElseIf Estado = IniciaRonda Then
        
        tiempo_ronda = tiempo_ronda - 1
        If tiempo_ronda Mod 30 = 0 And tiempo_ronda > 5 Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» En " & tiempo_ronda & " segundos comienza la ronda " & ronda_actual & ".", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_ronda <= 5 And tiempo_ronda > 0 Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» La ronda comienza en " & tiempo_ronda & ".", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_ronda = 0 Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �Ya!", e_FontTypeNames.FONTTYPE_GUILD))
            
            'Cambias triggers de los mapas para dejar pasar a los usuarios.

            
            For Each participante In participantes
                Call enviarUsuarioBase(participante.UserIndex)
            Next participante
        
            Estado = EnJuego
        End If
    ElseIf Estado = EnJuego Then
        Dim Team As Byte
        Dim estaEnBandera(1 To 2) As Boolean
        For Each participante In participantes
            With UserList(participante.UserIndex)
            Team = .flags.jugando_captura_team
            If .Pos.Map = mapaEquipo(Team) Then 'Si estoy en mi mapa
                If Abs(.Pos.X - xEquipo(Team)) <= 8 And Abs(.Pos.Y - yEquipo(Team)) <= 5 Then ' .Pos.x Distance(.Pos.X, .Pos.Y, xEquipo(Team), yEquipo(Team)) <= 8 Then 'Y estoy 5 pasos del centro
                    If TieneObjetos(objEquipo(IIf(Team = 1, 2, 1)), 1, participante.UserIndex) Then 'Y tengo la bandera enemgia
                        contadorBandera(Team) = contadorBandera(Team) + 1
                        estaEnBandera(Team) = True
                        If contadorBandera(Team) >= CAPTURA_TIEMPO_BANDERA Then
                            Call finalizarRonda(Team)
                        Else
                            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �Atenci�n! " & .Name & " del equipo " & colorEquipo(Team) & " tiene la bandera enemiga en su base, si aguanta " & CAPTURA_TIEMPO_BANDERA - contadorBandera(Team) & " segundos m�s ganar� la ronda.", e_FontTypeNames.FONTTYPE_GUILD))
                        End If
                        
                    End If
                End If
            ElseIf .flags.jugando_captura_timer > 0 Then
                .flags.jugando_captura_timer = .flags.jugando_captura_timer - 1
                If .flags.jugando_captura_timer = 0 Then
                    Call enviarUsuarioBase(participante.UserIndex)
                End If
            End If
            End With
        Next participante
        If Not estaEnBandera(1) Then
            contadorBandera(1) = 0
        End If
        
        If Not estaEnBandera(2) Then
            contadorBandera(2) = 0
        End If
    End If
End Sub
Private Function sortParticipantes() As Collection
    Dim i As Integer
    Dim e As Integer
    Dim UI As Integer
    Dim added As Boolean
    Set sortParticipantes = New Collection
    
    For i = 1 To participantes.Count
        added = False
        For e = 1 To sortParticipantes.Count
            UI = participantes.Item(i).UserIndex
            If UserList(UI).Stats.ELV > UserList(sortParticipantes.Item(e).UserIndex).Stats.ELV Then
                Call sortParticipantes.Add(participantes.Item(i), CStr(UI), e)
                added = True
                Exit For
            End If
        Next e
        If Not added Then
            Call sortParticipantes.Add(participantes.Item(i), CStr(UI))
        End If
    Next i
    
End Function
Private Sub iniciarCaptura()

    tiempo_espera = 0
    
    
    'Ordenar equipos
    Dim Team As Byte
    Dim Counter As Byte
    Dim participante As clsCapturaParticipante
    
    If participantes.Count Mod 2 <> 0 Then
        Dim UI As Integer
        UI = participantes(participantes.Count).UserIndex
        Call eliminarParticipante(UI)
        Call WriteConsoleMsg(UI, "Has sido eliminado del evento por ser el �ltimo en ingresar y los equipos eran impares.", e_FontTypeNames.FONTTYPE_INFOBOLD)
    End If
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» El evento captura la bandera ha iniciado.", e_FontTypeNames.FONTTYPE_GUILD))
    
    Dim sortedParticipantes As Collection
    Set sortedParticipantes = sortParticipantes()
    
    Team = 1
    Counter = 1
    'Armamos los equipos segun level para que sea lo mas parejo.
    For Each participante In sortedParticipantes
        participante.Team = Team
        UserList(participante.UserIndex).flags.jugando_captura_team = Team
        UserList(participante.UserIndex).flags.tiene_bandera = 0
         Call SendData(SendTarget.ToAll, participante.UserIndex, PrepareMessageUpdateFlag(UserList(participante.UserIndex).flags.tiene_bandera, UserList(participante.UserIndex).Char.charindex))
        Call WriteConsoleMsg(participante.UserIndex, "Te ha tocado participar en el equipo " & colorEquipo(participante.Team) & " deber�s proteger a tus compa�eros de equipo y capturar la bandera del equipo contrario para ganar.", e_FontTypeNames.FONTTYPE_GUILD)
        Counter = Counter + 1
        If Counter = 2 Then
            Team = IIf(Team = 1, 2, 1)
            Counter = 0
        End If
    Next participante
    
    


    Call iniciarRonda
End Sub

Public Sub muereUsuario(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        If .flags.jugando_captura_team = 0 Then
            Exit Sub
        End If
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» " & .Name & " del equipo " & colorEquipo(.flags.jugando_captura_team) & " ha muerto en batalla.", e_FontTypeNames.FONTTYPE_GUILD))
        
        Call tirarBandera(UserIndex)
        
        .flags.jugando_captura_muertes = .flags.jugando_captura_muertes + 1
        .flags.jugando_captura_timer = .flags.jugando_captura_muertes * CAPTURA_TIEMPO_POR_MUERTE + CAPTURA_TIEMPO_MUERTE
        
        
        Call WriteConsoleMsg(UserIndex, "Volver�s a tu base en " & .flags.jugando_captura_timer & " segundos.", e_FontTypeNames.FONTTYPE_GUILD)
        
        
        Call enviarUsuarioSala(UserIndex)
    
        Call ResucitarOCurar(UserIndex)
       
    End With
End Sub

Public Function tomaBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    With UserList(UserIndex)
        Dim banderaTeam As Byte
        
        If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
            banderaTeam = 1
        ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
            banderaTeam = 2
        Else
            tomaBandera = True
            Exit Function
        End If
        
        
        If banderaTeam = .flags.jugando_captura_team Then
             If .Pos.X = xEquipo(banderaTeam) And .Pos.Y = yEquipo(banderaTeam) Then
                tomaBandera = False
                Exit Function
            End If
        End If
        'Si toma la bandera del equipo contrario
        If UserList(UserIndex).flags.jugando_captura_team <> banderaTeam Then
            UserList(UserIndex).flags.tiene_bandera = banderaTeam
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(UserList(UserIndex).flags.tiene_bandera, UserList(UserIndex).Char.charindex))
        Else
            UserList(UserIndex).flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(UserList(UserIndex).flags.tiene_bandera, UserList(UserIndex).Char.charindex))
        End If
        tomaBandera = True
    End With
End Function


Public Sub quitarBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    With UserList(UserIndex)
        Dim banderaTeam As Byte
        
        If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
            banderaTeam = 1
        ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
            banderaTeam = 2
        Else
            Exit Sub
        End If
                
        If banderaTeam <> .flags.jugando_captura_team Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �Atenci�n! " & .Name & " del equipo " & colorEquipo(.flags.jugando_captura_team) & " ha tomado la bandera del equipo " & colorEquipo(banderaTeam) & ".", e_FontTypeNames.FONTTYPE_GUILD))
            UserList(UserIndex).flags.tiene_bandera = banderaTeam
        Else
            UserList(UserIndex).flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, UserList(UserIndex).Char.charindex))
            Call QuitarObjetos(ObjIndex, 1, UserIndex)
            Call restaurarBandera(banderaTeam)
            
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �Atenci�n! La bandera del equipo " & colorEquipo(banderaTeam) & " ha sido recuperada.", e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

Public Sub tiraBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    'El usuario tira la bandera
    With UserList(UserIndex)
        Dim Team As Byte
        
        If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
            Team = 1
        ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
            Team = 2
        Else
            Exit Sub
        End If
        
        contadorBandera(Team) = 0
        
        UserList(UserIndex).flags.tiene_bandera = 0
        
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(UserList(UserIndex).flags.tiene_bandera, UserList(UserIndex).Char.charindex))
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �Atenci�n! la bandera del equipo " & colorEquipo(Team) & " se encuentra en (" & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ")", e_FontTypeNames.FONTTYPE_GUILD))
        
    End With
End Sub

Public Sub tirarBandera(ByVal UserIndex As Integer)
    'Verificamos si tiene la bandera
    Dim Team As Byte
    Dim objBandera As Integer
    With UserList(UserIndex)
        Team = .flags.jugando_captura_team
        If Team = 0 Then
            Exit Sub
        ElseIf Team = 1 Then
            objBandera = OBJ_CAPTURA_BANDERA_2
        ElseIf Team = 2 Then
            objBandera = OBJ_CAPTURA_BANDERA_1
        End If
    
        
        If TieneObjetos(objBandera, 1, UserIndex) Then
            UserList(UserIndex).flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, UserList(UserIndex).Char.charindex))
            Call QuitarObjetos(objBandera, 1, UserIndex)
            
            If Estado = EnJuego Then
                Dim MiObj As t_Obj
        
                MiObj.Amount = 1
                MiObj.ObjIndex = objBandera
            
                Call TirarItemAlPiso(.Pos, MiObj, False)
                Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �Atenci�n! la bandera del equipo " & colorEquipo(Team) & " se encuentra en (" & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ")", e_FontTypeNames.FONTTYPE_GUILD))
            End If
            
        End If
    End With
End Sub

Private Sub enviarUsuarioBase(ByVal UserIndex As Integer)
    Dim pos_map As Integer
    Dim pos_x As Integer
    Dim pos_y As Integer
    If UserList(UserIndex).flags.jugando_captura_team = 1 Then
        pos_map = MAP_TEAM_1
        pos_x = X_TEAM_1
        pos_y = Y_TEAM_1
    ElseIf UserList(UserIndex).flags.jugando_captura_team = 2 Then
        pos_map = MAP_TEAM_2
        pos_x = X_TEAM_2
        pos_y = Y_TEAM_2
    Else
        Exit Sub
    End If
    Call FindLegalPos(UserIndex, pos_map, pos_x, pos_y)
    Call WarpUserChar(UserIndex, pos_map, pos_x, pos_y, True)

End Sub

Private Sub enviarUsuarioSala(ByVal UserIndex As Integer)
    Dim pos_x As Integer, pos_y As Integer
    pos_x = RandomNumber(MIN_SALA_ESPERA_x, MAX_SALA_ESPERA_x)
    pos_y = RandomNumber(MIN_SALA_ESPERA_y, MAX_SALA_ESPERA_y)
    
    'No.
    Call FindLegalPos(UserIndex, MAP_SALA_ESPERA, pos_x, pos_y)
    Call WarpUserChar(UserIndex, MAP_SALA_ESPERA, pos_x, pos_y, True)
End Sub

Private Sub iniciarRonda()
    Estado = IniciaRonda
    ronda_actual = ronda_actual + 1
    tiempo_ronda = CAPTURA_TIEMPO_INICIO_RONDA
    
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» La ronda " & ronda_actual & " dar� comienzo en " & tiempo_ronda & " segundos.", e_FontTypeNames.FONTTYPE_GUILD))

    
    Dim participante As clsCapturaParticipante
    For Each participante In participantes
        Call SendData(SendTarget.ToAll, participante.UserIndex, PrepareMessageUpdateFlag(0, UserList(participante.UserIndex).Char.charindex))
        UserList(participante.UserIndex).flags.jugando_captura_muertes = 0
        UserList(participante.UserIndex).flags.jugando_captura_timer = 0
        Call tirarBandera(participante.UserIndex)
        Call enviarUsuarioSala(participante.UserIndex)
    Next participante
    
    'Ponemos las banderas de nuevo
    Call restaurarBandera(1)
    Call restaurarBandera(2)
    

End Sub

Private Sub restaurarBandera(ByVal Team As Byte)
    Dim MiObj As t_Obj
    Dim MiPos As t_WorldPos
    MiObj.Amount = 1
    MiObj.ObjIndex = objEquipo(Team)
    MiPos.Map = mapaEquipo(Team)
    MiPos.X = xEquipo(Team)
    MiPos.Y = yEquipo(Team)
    
    If MapData(MiPos.Map).Tile(MiPos.X, MiPos.Y).ObjInfo.Amount <= 0 Then
        Call TirarItemAlPiso(MiPos, MiObj, False)
    End If
    
End Sub


Public Sub finalizarCaptura(Optional Team As Integer = -1)

    Dim participante As clsCapturaParticipante
    Estado = Finalizado
    If Team = -2 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Se ha cancelado el evento de captura la bandera por falta de participantes.", e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf Team = -1 Then 'Se aborto el evento
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» El evento de captura la bandera ha sido cancelado. Se te ha devuelto el precio de inscripci�n.", e_FontTypeNames.FONTTYPE_GUILD))
         For Each participante In participantes
            UserList(participante.UserIndex).Stats.GLD = UserList(participante.UserIndex).Stats.GLD + Precio
            Call WriteUpdateGold(participante.UserIndex)
        Next participante
    ElseIf Team = 0 Then 'Hubo empate
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» �El evento de captura la bandera ha terminado en empate!.", e_FontTypeNames.FONTTYPE_GUILD))
    Else
    
        'Hay premios???
    
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» �El evento de captura la bandera ha terminado! Ganador el equipo " & colorEquipo(Team) & ".", e_FontTypeNames.FONTTYPE_GUILD))
    End If

    
    For Each participante In participantes
        Call eliminarParticipante(participante.UserIndex)
    Next participante

    'Destruyo la instancia
    Set InstanciaCaptura = Nothing
End Sub

Public Sub finalizarRonda(ByVal Team As Byte)

    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» �El equipo " & colorEquipo(Team) & " ha ganado la ronda " & ronda_actual & "!", e_FontTypeNames.FONTTYPE_GUILD))
    'El sonido de fundar clan para hacerlo mas magestuoso, mirar si hay un sonido mejor
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
    
    contadorBandera(1) = 0
    contadorBandera(2) = 0
    
    
    rondasGanadas(Team) = rondasGanadas(Team) + 1
    If ronda_actual >= cantidad_rondas Then
        'Se termina la joda
        
        If rondasGanadas(1) > rondasGanadas(2) Then
            Team = 1
        ElseIf rondasGanadas(2) > rondasGanadas(1) Then
            Team = 2
        Else
            Team = 0
        End If
        
        Call finalizarCaptura(CInt(Team))
    Else
        Call iniciarRonda
    End If
End Sub

Public Function Inicializar(ByVal MaxParticipantes As Long, ByVal CantidadRondas As Byte, ByVal NivelMinimo As Byte, ByVal precio_ As Long) As Boolean
    
    'Agregar random en comando para que no lo macreen rr-magax
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Se ha iniciado un Captura de la bandera de " & MaxParticipantes & " participantes, el precio de inscripci�n es de " & PonerPuntos(precio_) & " y el nivel m�nimo para ingresar es " & NivelMinimo & ", para participar env�a /PARTICIPARCAPTURA Las inscripciones se cerraran en 3 minutos.", e_FontTypeNames.FONTTYPE_GUILD))
        
    max_participantes = MaxParticipantes
    cantidad_rondas = CantidadRondas
    nivel_minimo = NivelMinimo
    tiempo_espera = CAPTURA_TIEMPO_ESPERA
    Precio = precio_
    Estado = Inscripcion
End Function

Public Sub inscribirse(ByVal UserIndex As Integer)
    'Validaciones al inscribirse
    
    With UserList(UserIndex)
        
        If Estado <> Inscripcion Then
            Call WriteConsoleMsg(UserIndex, "Ya se ha cerrado la inscripci�n para el evento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.GLD < Precio Then
            Call WriteConsoleMsg(UserIndex, "No tienes oro suficiente para ingresar al evento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.jugando_captura = 1 Then
            Call WriteConsoleMsg(UserIndex, "Ya est�s anotado en el evento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Nivel
        If .Stats.ELV < nivel_minimo Then
            Call WriteConsoleMsg(UserIndex, "Para ingresar al evento deber�s ser al menos nivel " & nivel_minimo & ".", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Zona segura
        If Zona(.ZonaId).Segura = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solamente podr�s ingresar al evento en zona segura.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'cupos completos
        If participantes.Count >= max_participantes Then
            Call WriteConsoleMsg(UserIndex, "Los cupos ya est�n completos.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'que no est� en carcel
        If .Counters.Pena > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes anotarte al evento mientras est�s encarcelado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'que no est� muerto
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes anotarte a un evento estando muerto.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'que no est� comerciando
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No puedes ingresar al evento mientras est�s comerciando.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Montado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes ingresar al evento estando montado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes ingresar al evento mientras est�s navegando.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim participante As clsCapturaParticipante
        Set participante = New clsCapturaParticipante
                
        participante.UserIndex = UserIndex
        participante.LastPosMap = UserList(UserIndex).Pos.Map
        participante.LastPosx = UserList(UserIndex).Pos.X
        participante.LastPosy = UserList(UserIndex).Pos.Y
        
        
        Call participantes.Add(participante, CStr(UserIndex))
        
        UserList(UserIndex).flags.jugando_captura = 1
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
        
        Call WriteUpdateGold(UserIndex)
        Call enviarUsuarioSala(UserIndex)
        
        If participantes.Count >= max_participantes Then
            Call iniciarCaptura
        End If
        
    End With
End Sub
