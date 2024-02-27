Attribute VB_Name = "modServidores"
Option Explicit

Public Type tServidor
    nombre As String
    Host As String
    Puerto As Integer
End Type

Public Servidores() As tServidor

Public Sub LoadServers()
    
    On Error GoTo CargarLst_Err
    
    Dim servidorSeleccionado() As String
    servidorSeleccionado = Split(ServerIndex, ":")
    
    Dim total As Integer
    Dim file_path As String
    
    file_path = RESOURCES_PATH & "/INIT/Servidores.ini"
    
    total = Val(GetVar(file_path, "SERVIDORES", "Total"))
    If total < 1 Then
        ReDim Servidores(1 To 1) As tServidor
        Servidores(1).nombre = "Localhost"
        Servidores(1).Host = "127.0.0.1"
        Servidores(1).Puerto = 7666
    End If

    ReDim Servidores(1 To total) As tServidor
    frmLogear.lstServidores.Clear
    
    Dim k As Integer
    For k = 1 To total
        Servidores(k).nombre = GetVar(file_path, "SERVIDOR" & k, "Nombre")
        Servidores(k).Host = GetVar(file_path, "SERVIDOR" & k, "Host")
        Servidores(k).Puerto = Val(GetVar(file_path, "SERVIDOR" & k, "Puerto"))
        frmLogear.lstServidores.AddItem Servidores(k).nombre
        
        If UBound(servidorSeleccionado) = 1 Then  ' selected
            If servidorSeleccionado(0) = Servidores(k).Host Then
                frmLogear.lstServidores.ListIndex = (k - 1)
            End If
        End If
    Next
    
    If UBound(servidorSeleccionado) <> 1 Then ' default
        servidorSeleccionado = Split(Servidores(1).Host & ":" & Servidores(1).Puerto, ":")
        frmLogear.lstServidores.ListIndex = 0
    End If
     
    frmLogear.txtIp.Text = servidorSeleccionado(0)
    frmLogear.txtPort.Text = servidorSeleccionado(1)
    IPdelServidor = servidorSeleccionado(0)
    PuertoDelServidor = servidorSeleccionado(1)

    Exit Sub

CargarLst_Err:
    Call RegistrarError(err.Number, err.Description, "ModLadder.CargarLst", Erl)
    Resume Next
    
End Sub

'Public Function randomIp() As String
'    Dim id As Long
'    id = RandomNumber(1, 3)
'    Select Case id
'        Case 1
'            randomIp = "45.235.98.33"
'            Exit Function
'        Case 2
'            randomIp = "45.235.98.34"
'            Exit Function
'
'        Case 3
'            randomIp = "45.235.98.35"
'            Exit Function
'    End Select
'End Function
'
'
'Public Function get_logging_server() As String
'    Dim Value As Long
'    Dim k As Integer
'    For k = 1 To 100
'        Value = RandomNumber(1, 100)
'    Next k
'    If Value <= 50 Then
'        get_logging_server = servers_login_connections(1)
'    Else
'        get_logging_server = servers_login_connections(2)
'    End If
'End Function
'
'Public Function SetDefaultServer()
'On Error GoTo SetDefaultServer_Err
'    Dim serverLogin() As String
'    serverLogin = Split(get_logging_server(), ":")
'    IPdelServidorLogin = serverLogin(0)
'    PuertoDelServidorLogin = serverLogin(1)
'    Exit Function
'SetDefaultServer_Err:
'    Call RegistrarError(err.Number, err.Description, "Mod_General.WriteVar", Erl)
'End Function
