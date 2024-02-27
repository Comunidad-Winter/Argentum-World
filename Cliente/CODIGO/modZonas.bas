Attribute VB_Name = "modZonas"
Public Type t_ZonaInfo
    Zona_name As String
    
    Deleted As Byte
    map As Integer
    x As Integer
    y As Integer
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
    
    Faccion As Byte
        
    SalidaMap As Integer
    SalidaX As Integer
    SalidaY As Integer
    
    Musica1 As Integer
    Musica2 As Integer
    Musica3 As Integer
    
End Type

Private Type t_ListZona
    Zona() As t_ZonaInfo
    Cant As Integer
End Type

Public MapZona() As t_ListZona
Public NumZonas As Integer
Public UserZona As Integer
Public ZonaName As String
Public ZoneSegura As Byte
Public ZonaSeguraDir As Integer
Public ZonaSeguraAlpha As Single
Public Sub LoadZonas()
    On Error GoTo ErrHandler
    
    If Not FileExist(RESOURCES_PATH & "/DAT/Zonas.dat", vbArchive) Then
        MsgBox "No se encuentra " & RESOURCES_PATH & "/DAT/Zonas.dat", vbCritical
        End
    End If

    Dim i As Integer

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(RESOURCES_PATH & "/DAT/Zonas.dat")

    Dim Key As String
    NumZonas = Lector.GetValue("INIT", "Cantidad")
    
    ReDim MapZona(1 To NUM_MAPAS)
    
    For i = 1 To NumZonas
        Key = "Zona" & i
        map = Val(Lector.GetValue(Key, "Map"))
        MapZona(map).Cant = MapZona(map).Cant + 1
        ReDim Preserve MapZona(map).Zona(MapZona(map).Cant)
        With MapZona(map).Zona(MapZona(map).Cant)
            
            .Zona_name = Lector.GetValue(Key, "Name")
            .map = map
            .Deleted = Val(Lector.GetValue(Key, "Deleted"))
            .x = Val(Lector.GetValue(Key, "X"))
            .y = Val(Lector.GetValue(Key, "Y"))
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

Public Function getZona(ByVal x As Integer, ByVal y As Integer) As Integer
    Dim i As Integer
    For i = 1 To MapZona(UserMap).Cant
        With MapZona(UserMap).Zona(i)
            If x >= .x And x <= .X2 And y >= .y And y <= .Y2 Then
                getZona = i
                Exit Function
            End If
        End With
    Next i
    getZona = 0
End Function

Public Sub checkZona()

UserZona = getZona(UserPos.x, UserPos.y)

If MapZona(UserMap).Zona(UserZona).Zona_name <> ZonaName Then
    ZonaName = MapZona(UserMap).Zona(UserZona).Zona_name
    frmMain.NameMapa.Caption = ZonaName
    If ZonaName <> "" Then
        Letter_Set 0, ZonaName
    End If
End If

If MapZona(UserMap).Zona(UserZona).Segura <> ZoneSegura Then
    ZoneSegura = MapZona(UserMap).Zona(UserZona).Segura
    If ZonaSeguraAlpha = 0 Then ZonaSeguraAlpha = 255
    ZonaSeguraDir = 1
End If

End Sub
