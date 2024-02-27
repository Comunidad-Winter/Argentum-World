Attribute VB_Name = "modMapIO"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
'MOTOR DX8 POR LADDER
''
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit
Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat
Public UserMap As Integer
''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByVal Filename As String) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo FalloFile

    Dim nFileNum  As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open Filename For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1

End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    
    On Error GoTo FileExist_Err
    

    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    If LenB(Dir(File, FileType)) = 0 Then
        FileExist = False
    Else
        FileExist = True

    End If

    
    Exit Function

FileExist_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.FileExist", Erl)
    Resume Next
    
End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByVal Path As String)
    
    On Error GoTo AbrirMapa_Err
    

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    FrmMain.MousePointer = vbHourglass
    Dim ind As Integer
    ind = InStrRev(Path, "\") + 5
    UserMap = mid$(Path, ind, Len(Path) - ind - 3)
    FrmMain.Label16.Caption = "Map " & UserMap
    
    SelectedZona = 0
    SelectedSpawn = 0
    
    SurfaceDB.UnloadAll
    
    Call Load_Map_Data_CSM(Path)
    
    CargarMinimap

    
    FrmMain.MousePointer = vbDefault
    

    
    Exit Sub

AbrirMapa_Err:

    FrmMain.MousePointer = vbDefault
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.AbrirMapa", Erl)
    Resume Next
    
End Sub
Public Sub CargarMinimap()
Dim PathMini As String
    PathMini = App.Path & "\..\Resources\Minimapas\Mapa" & UserMap & "x1.bmp"
    
    
    If MapSize.Height / MapSize.Width > 745 / 420 Then
        FrmMain.MiniMap.Height = 745
        FrmMain.MiniMap.Width = MapSize.Width * (FrmMain.MiniMap.Height / MapSize.Height)
    Else
        FrmMain.MiniMap.Width = 420
        FrmMain.MiniMap.Height = MapSize.Height * (FrmMain.MiniMap.Width / MapSize.Width)
    End If
    
   
    
    FrmMain.ApuntadorRadar.Width = 36 * (FrmMain.MiniMap.Width / MapSize.Width)
    FrmMain.ApuntadorRadar.Height = 26 * (FrmMain.MiniMap.Width / MapSize.Width)
    
    If FileExist(PathMini, vbNormal) Then
        FrmMain.MiniMap.Picture = LoadPicture(PathMini)
        FrmMain.MiniMap.ScaleMode = 3
        FrmMain.MiniMap.AutoRedraw = True
        FrmMain.MiniMap.PaintPicture FrmMain.MiniMap.Picture, _
        0, 0, FrmMain.MiniMap.ScaleWidth, FrmMain.MiniMap.ScaleHeight, _
        0, 0, _
        FrmMain.MiniMap.Picture.Width / 26.46, _
        FrmMain.MiniMap.Picture.Height / 26.46
        FrmMain.MiniMap.Picture = FrmMain.MiniMap.Image
        
        Call FrmMain.DibujarZonas
    Else
        FrmMain.MiniMap.Picture = Nothing
    End If
End Sub




''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    On Error GoTo DeseaGuardarMapa_Err
    

    If MapInfo.Changed = 1 Then
        FrmMain.mnuGuardarMapa_Click

    End If

    
    Exit Sub

DeseaGuardarMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.DeseaGuardarMapa", Erl)
    Resume Next
    
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 21/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc As Integer
    Dim Y     As Integer
    Dim X     As Integer

    bAutoGuardarMapaCount = 0

    'frmMain.mnuUtirialNuevoFormato.Checked = True
    FrmMain.mnuReAbrirMapa.Enabled = False
    FrmMain.TimAutoGuardarMapa.Enabled = False

    MapaCargado = False

    FrmMain.MousePointer = 11
    
    MapSize.XMin = 1
    MapSize.YMin = 1
    MapSize.Width = 100
    MapSize.Height = 100
    
    ReDim MapData(1 To 100, 1 To 100)
    'ReDim MapData_Deshacer(1 To maxDeshacer)

    For Y = 1 To MapSize.Height
        For X = 1 To MapSize.Width
    
            ' Capa 1
            MapData(X, Y).Graphic(1).grhindex = 1
        
            ' Bloqueos
            MapData(X, Y).Blocked = 0

            ' Capas 2, 3 y 4
            MapData(X, Y).Graphic(2).grhindex = 0
            MapData(X, Y).Graphic(3).grhindex = 0
            MapData(X, Y).Graphic(4).grhindex = 0

            ' NPCs
            If MapData(X, Y).CharIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NpcIndex = 0

            End If

            ' OBJs
            MapData(X, Y).OBJInfo.ObjIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.grhindex = 0

            ' Translados
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        
            ' Triggers
            MapData(X, Y).Trigger = 0
        
            InitGrh MapData(X, Y).Graphic(1), 1
        Next X
    Next Y

    MapInfo.MapVersion = 0
    MapInfo.name = "Nuevo Mapa"
    MapInfo.Music = 0
    MapInfo.PK = True
    MapInfo.MagiaSinEfecto = 0
    MapInfo.InviSinEfecto = 0
    MapInfo.ResuSinEfecto = 0
    MapInfo.Terreno = "BOSQUE"
    MapInfo.Zona = "CAMPO"
    MapInfo.Restringir = "No"
    MapInfo.NoEncriptarMP = 0

    Call MapInfo_Actualizar

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 0
    FrmMain.MousePointer = 0

    ' Vacio deshacer
    modEdicion.Deshacer_Clear

    MapaCargado = True
    EngineRun = True

    'FrmMain.SetFocus

End Sub




' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo MapInfo_Guardar_Err
    

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save

    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))
    
    Call WriteVar(Archivo, MapTitulo, "Light", MapInfo.Light)
    
    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.Backup))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")

    End If

    
    Exit Sub

MapInfo_Guardar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.MapInfo_Guardar", Erl)
    Resume Next
    
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

    Dim Leer  As New clsIniReader
    Dim loopc As Integer
    Dim Path  As String
    MapTitulo = Empty
    Leer.Initialize Archivo

    For loopc = Len(Archivo) To 1 Step -1

        If mid(Archivo, loopc, 1) = "\" Then
            Path = Left(Archivo, loopc)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    MapInfo.Light = Leer.GetValue(MapTitulo, "Light")

    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False

    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.Backup = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    'FORMATO IAO
    MapDat.map_name = MapInfo.name
    MapDat.backup_mode = MapInfo.Backup
    MapDat.restrict_mode = MapInfo.Restringir
    MapDat.music_numberLow = MapInfo.Music
    MapDat.zone = MapInfo.Zona
    MapDat.terrain = MapInfo.Terreno

    MidiMusic = MapDat.music_numberLow
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

   


End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestañas(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    
    On Error GoTo Pestañas_Err
    

    Dim loopc As Integer

    If FormatoIAO Then

        For loopc = Len(Map) To 1 Step -1

            If mid(Map, loopc, 1) = "\" Then
                PATH_Save = Left(Map, loopc)
                Exit For

            End If

        Next
        Map = Right(Map, Len(Map) - (Len(PATH_Save)))

        For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

            If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
                'NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
                'NameMap_Save = Left(Map, loopc)
                Exit For

            End If

        Next

    Else

        For loopc = Len(Map) To 1 Step -1

            If mid(Map, loopc, 1) = "\" Then
                PATH_Save = Left(Map, loopc)
                Exit For

            End If

        Next
        Map = Right(Map, Len(Map) - (Len(PATH_Save)))

        For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

            If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
                NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
                NameMap_Save = Left(Map, loopc)
                Exit For

            End If

        Next


    End If

    
    Exit Sub

Pestañas_Err:
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.Pestañas", Erl)
    Resume Next
    
End Sub
