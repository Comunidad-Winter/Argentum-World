Attribute VB_Name = "modZonas"
Public Type t_ZonaInfo
    Zona_name As String
    
    Deleted As Byte
    Map As Integer
    X As Integer
    Y As Integer
    X2 As Integer
    Y2 As Integer
    
    Backup As Byte
    
    Lluvia As Byte
    Nieve As Byte
    Niebla As Byte
    
    Ambient As String
    Base_light As Long
    Terreno As String
    
    MinLevel As Integer
    MaxLevel As Integer
    Segura As Byte
    Newbie As Byte
    
    SoloClanes As Byte
    SoloFaccion As Byte
    SinMagia As Byte
    SinInvi As Byte
    SinMascotas As Byte
    SinResucitar As Byte
    Interdimensional As Byte
    
    Faccion As Byte
        
    SalidaMap As Integer
    SalidaX As Integer
    SalidaY As Integer
    
    Musica1 As Integer
    Musica2 As Integer
    Musica3 As Integer
    
End Type

Public Type t_NpcSpawn_List
    NpcIndex As Integer
    Cantidad As Integer
End Type

Public Type t_NpcSpawn
    Deleted As Byte
    Map As Integer
    X As Integer
    Y As Integer
    X2 As Integer
    Y2 As Integer
    CantNpcs As Byte
    NPCs() As t_NpcSpawn_List
End Type

Public Zona() As t_ZonaInfo
Public NumZonas As Integer
Public NpcSpawn() As t_NpcSpawn
Public NumSpawns As Integer
Public SelectedZona As Integer
Public SelectedSpawn As Integer

Public Sub LoadZonas()
    On Error GoTo ErrHandler

    Dim i As Integer

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(App.Path & "\..\Resources\Dat\Zonas.dat")
    

    Dim Key As String
    NumZonas = Lector.GetValue("INIT", "Cantidad")
    
    ReDim Zona(1 To NumZonas)
    
    For i = 1 To NumZonas
        Key = "Zona" & i
        With Zona(i)
            
            .Zona_name = Lector.GetValue(Key, "Name")
            .Map = Val(Lector.GetValue(Key, "Map"))
            .Deleted = Val(Lector.GetValue(Key, "Deleted"))
            .X = Val(Lector.GetValue(Key, "X"))
            .Y = Val(Lector.GetValue(Key, "Y"))
            .X2 = Val(Lector.GetValue(Key, "X2"))
            .Y2 = Val(Lector.GetValue(Key, "Y2"))
            
            .Backup = Val(Lector.GetValue(Key, "Backup"))
            .Lluvia = Val(Lector.GetValue(Key, "Lluvia"))
            .Nieve = Val(Lector.GetValue(Key, "Nieve"))
            .Niebla = Val(Lector.GetValue(Key, "Niebla"))
            .MinLevel = Val(Lector.GetValue(Key, "MinLevel"))
            .MaxLevel = Val(Lector.GetValue(Key, "MaxLevel"))
            .Segura = Val(Lector.GetValue(Key, "Segura"))
            .Newbie = Val(Lector.GetValue(Key, "Newbie"))
            .SinMagia = Val(Lector.GetValue(Key, "SinMagia"))
            .SinInvi = Val(Lector.GetValue(Key, "SinInvi"))
            .SinMascotas = Val(Lector.GetValue(Key, "SinMascotas"))
            .SinResucitar = Val(Lector.GetValue(Key, "SinResucitar"))
            .SoloClanes = Val(Lector.GetValue(Key, "SoloClanes"))
            .SoloFaccion = Val(Lector.GetValue(Key, "SoloFaccion"))
            .Interdimensional = Val(Lector.GetValue(Key, "Interdimensional"))

            .Faccion = Val(Lector.GetValue(Key, "Faccion"))
            .Terreno = Lector.GetValue(Key, "Terreno")
            .Ambient = Lector.GetValue(Key, "Ambient")
            .Base_light = Val(Lector.GetValue(Key, "Base_light"))
            .SalidaMap = Val(Lector.GetValue(Key, "SalidaMap"))
            .SalidaX = Val(Lector.GetValue(Key, "SalidaX"))
            .SalidaY = Val(Lector.GetValue(Key, "SalidaY"))
            
            .Musica1 = Val(Lector.GetValue(Key, "Musica1"))
            .Musica2 = Val(Lector.GetValue(Key, "Musica2"))
            .Musica3 = Val(Lector.GetValue(Key, "Musica3"))
        End With
        
        
    Next i
    
    Exit Sub
    
ErrHandler:

End Sub
Public Sub LoadNpcSpawn()
    On Error GoTo ErrHandler

    Dim i As Integer

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(App.Path & "\..\Resources\Dat\NpcSpawns.dat")
    
    Dim Cant As Integer
    Dim Key As String
    NumSpawns = Lector.GetValue("INIT", "Cantidad")
    
    ReDim NpcSpawn(1 To NumSpawns)
    
    For i = 1 To NumSpawns
        Key = "NpcSpawn" & i
        With NpcSpawn(i)
            
            .Deleted = Val(Lector.GetValue(Key, "Deleted"))
            .Map = Val(Lector.GetValue(Key, "Map"))
            .X = Val(Lector.GetValue(Key, "X"))
            .Y = Val(Lector.GetValue(Key, "Y"))
            .X2 = Val(Lector.GetValue(Key, "X2"))
            .Y2 = Val(Lector.GetValue(Key, "Y2"))
            .CantNpcs = Val(Lector.GetValue(Key, "CantNpcs"))
            If .CantNpcs > 0 Then
            ReDim .NPCs(1 To .CantNpcs)
            For X = 1 To .CantNpcs
                .NPCs(X).NpcIndex = Val(Lector.GetValue(Key, "NpcIndex" & X))
                .NPCs(X).Cantidad = Val(Lector.GetValue(Key, "NpcCantidad" & X))
            Next X
            End If
        End With
        
        
    Next i
    
    Exit Sub
    
ErrHandler:

End Sub
Public Function SaveZona(i As Integer, nZona As t_ZonaInfo) As Integer
    Dim Path As String
    Path = App.Path & "\..\Resources\Dat\Zonas.dat"
    
    If i = 0 Then
    
        For i = 1 To NumZonas
            If Zona(i).Deleted = 1 Then
                Exit For
            End If
        Next i
    
        If i > NumZonas Then
            NumZonas = NumZonas + 1
            i = NumZonas
            ReDim Preserve Zona(1 To NumZonas)
        End If
    End If
    


    With nZona
        Call WriteVar(Path, "Init", "Cantidad", NumZonas)
        Call WriteVar(Path, "Zona" & i, "Name", .Zona_name)
        Call WriteVar(Path, "Zona" & i, "Deleted", .Deleted)
        Call WriteVar(Path, "Zona" & i, "Map", .Map)
        Call WriteVar(Path, "Zona" & i, "X", .X)
        Call WriteVar(Path, "Zona" & i, "Y", .Y)
        Call WriteVar(Path, "Zona" & i, "X2", .X2)
        Call WriteVar(Path, "Zona" & i, "Y2", .Y2)
        Call WriteVar(Path, "Zona" & i, "Backup", .Backup)
        Call WriteVar(Path, "Zona" & i, "Lluvia", .Lluvia)
        Call WriteVar(Path, "Zona" & i, "Nieve", .Nieve)
        Call WriteVar(Path, "Zona" & i, "Niebla", .Niebla)
        Call WriteVar(Path, "Zona" & i, "Terreno", .Terreno)
        Call WriteVar(Path, "Zona" & i, "MinLevel", .MinLevel)
        Call WriteVar(Path, "Zona" & i, "MaxLevel", .MaxLevel)
        Call WriteVar(Path, "Zona" & i, "Segura", .Segura)
        Call WriteVar(Path, "Zona" & i, "Newbie", .Newbie)
        Call WriteVar(Path, "Zona" & i, "SoloClanes", .SoloClanes)
        Call WriteVar(Path, "Zona" & i, "SoloFaccion", .SoloFaccion)
        Call WriteVar(Path, "Zona" & i, "Faccion", .Faccion)
        Call WriteVar(Path, "Zona" & i, "SinMagia", .SinMagia)
        Call WriteVar(Path, "Zona" & i, "SinInvi", .SinInvi)
        Call WriteVar(Path, "Zona" & i, "SinMascotas", .SinMascotas)
        Call WriteVar(Path, "Zona" & i, "SinResucitar", .SinResucitar)
        Call WriteVar(Path, "Zona" & i, "SalidaMap", .SalidaMap)
        Call WriteVar(Path, "Zona" & i, "SalidaX", .SalidaX)
        Call WriteVar(Path, "Zona" & i, "SalidaY", .SalidaY)
        Call WriteVar(Path, "Zona" & i, "Ambient", .Ambient)
        Call WriteVar(Path, "Zona" & i, "Interdimensional", .Interdimensional)
    End With
    
    Zona(i) = nZona
    SaveZona = i
End Function
Public Function SaveSpawn(i As Integer, nSpawn As t_NpcSpawn) As Integer
    Dim Path As String
    Dim X As Integer
    Path = App.Path & "\..\Resources\Dat\NpcSpawns.dat"
    
    If i = 0 Then
    
        For i = 1 To NumSpawns
            If NpcSpawn(i).Deleted = 1 Then
                Exit For
            End If
        Next i
    
        If i > NumSpawns Then
            NumSpawns = NumSpawns + 1
            i = NumSpawns
            ReDim Preserve NpcSpawn(1 To NumSpawns)
        End If
    End If
    


    With nSpawn
        Call WriteVar(Path, "Init", "Cantidad", NumSpawns)
        Call WriteVar(Path, "NpcSpawn" & i, "Deleted", .Deleted)
        Call WriteVar(Path, "NpcSpawn" & i, "Map", .Map)
        Call WriteVar(Path, "NpcSpawn" & i, "X", .X)
        Call WriteVar(Path, "NpcSpawn" & i, "Y", .Y)
        Call WriteVar(Path, "NpcSpawn" & i, "X2", .X2)
        Call WriteVar(Path, "NpcSpawn" & i, "Y2", .Y2)
        Call WriteVar(Path, "NpcSpawn" & i, "CantNpcs", .CantNpcs)
        For X = 1 To .CantNpcs
            Call WriteVar(Path, "NpcSpawn" & i, "NpcIndex" & X, .NPCs(X).NpcIndex)
            Call WriteVar(Path, "NpcSpawn" & i, "NpcCantidad" & X, .NPCs(X).Cantidad)
        Next X
        
        
    End With
    
    NpcSpawn(i) = nSpawn
    SaveSpawn = i
End Function
Public Sub DeleteZona(id As Integer)
    Dim Path As String
    Path = App.Path & "\..\Resources\Dat\Zonas.dat"
    If id = NumZonas Then
        NumZonas = NumZonas - 1
        ReDim Preserve Zona(1 To NumZonas)
        Call WriteVar(Path, "Init", "Cantidad", NumZonas)
    Else
        Zona(id).Deleted = 1
    End If
End Sub


Public Sub DeleteSpawn(id As Integer)
    Dim Path As String
    Path = App.Path & "\..\Resources\Dat\NpcSpawns.dat"
    If id = NumSpawns Then
        NumSpawns = NumSpawns - 1
        ReDim Preserve NpcSpawn(1 To NumSpawns)
        Call WriteVar(Path, "Init", "Cantidad", NumSpawns)
    Else
        NpcSpawn(id).Deleted = 1
    End If
End Sub

