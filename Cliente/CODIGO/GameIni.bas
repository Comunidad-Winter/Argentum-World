Attribute VB_Name = "GameIni"
Option Explicit

Public RESOURCES_PATH As String

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    
    On Error GoTo IniciarCabecera_Err
    
    Cabecera.desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
    
    Exit Sub
IniciarCabecera_Err:
    Call RegistrarError(err.Number, err.Description, "GameIni.IniciarCabecera", Erl)
    Resume Next
    
End Sub

Public Sub InitClient()
    
    RESOURCES_PATH = GetVar(App.path & "/cliente.ini", "DIRECTORIOS", "Recursos")
    
    If Len(RESOURCES_PATH) = 0 Or Len(dir(RESOURCES_PATH, vbDirectory)) = 0 Then
        RESOURCES_PATH = "../RESOURCES/"
    End If

End Sub
