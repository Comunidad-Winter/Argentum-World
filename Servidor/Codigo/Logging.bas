Attribute VB_Name = "modLogging"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

 Private Enum type_log
    e_LogearEventoDeSubasta = 0
    e_LogBan = 1
    e_LogCreditosPatreon = 2
    e_LogShopTransactions = 3
    e_LogShopErrors = 4
    e_LogEdicionPaquete = 5
    e_LogMacroServidor = 6
    e_LogMacroCliente = 7
    e_LogVentaCasa = 8
    e_LogCriticEvent = 9
    e_LogEjercitoReal = 10
    e_LogEjercitoCaos = 11
    e_LogError = 12
    e_LogPerformance = 13
    e_LogConsulta = 14
    e_LogClanes = 15
    e_LogGM = 16
    e_LogPremios = 17
    e_LogDatabaseError = 18
    e_LogSecurity = 19
 End Enum


Public Sub LogThis(nErrNo As Long, sLogMsg As String, EventType As LogEventTypeConstants)
    Dim filenum As Integer
    Dim msg As String
    msg = Time & " Error number: " & nErrNo & " | Description: " & sLogMsg & vbNewLine
    filenum = FreeFile
    Debug.Print msg
    
    Dim fileName As String
    Select Case EventType
        Case e_LogearEventoDeSubasta
            fileName = "LogearEventoDeSubasta.log"
        Case e_LogBan
            fileName = "LogBan.log"
        Case e_LogCreditosPatreon
            fileName = "LogCreditosPatreon.log"
        Case e_LogShopTransactions
            fileName = "LogShopTransactions.log"
        Case e_LogShopErrors
            fileName = "LogShopErrors.log"
        Case e_LogEdicionPaquete
            fileName = "LogEdicionPaquete.log"
        Case e_LogMacroServidor
            fileName = "LogMacroServidor.log"
        Case e_LogMacroCliente
            fileName = "LogMacroCliente.log"
        Case e_LogVentaCasa
            fileName = "LogVentaCasa.log"
        Case e_LogCriticEvent
            fileName = "LogCriticEvent.log"
        Case e_LogEjercitoReal
            fileName = "LogEjercitoReal.log"
        Case e_LogEjercitoCaos
            fileName = "LogEjercitoCaos.log"
        Case e_LogError
            fileName = "LogError.log"
        Case e_LogPerformance
            fileName = "LogPerformance.log"
        Case e_LogConsulta
            fileName = "LogConsulta.log"
        Case e_LogClanes
            fileName = "LogClanes.log"
        Case e_LogGM
            fileName = "LogGM.log"
        Case e_LogPremios
            fileName = "LogPremios.log"
        Case e_LogDatabaseError
            fileName = "LogDatabaseError.log"
        Case e_LogSecurity
            fileName = "LogSecurity.log"
    End Select
    
    Open App.Path & "\Logs\errores.log" For Append As filenum
        Print #filenum, msg
    Close filenum
    
End Sub

Public Sub LogearEventoDeSubasta(s As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogearEventoDeSubasta, "[Subastas.log] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal userindex As Integer, ByVal Motivo As String)
On Error GoTo ErrHandler
        Dim s As String
        s = UserList(BannedIndex).Name & " BannedBy " & UserList(userindex).Name & " Reason " & Motivo
        Call LogThis(type_log.e_LogBan, "[Bans] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogCreditosPatreon(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogCreditosPatreon, "[MonetizationCreditosPatreon.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogShopTransactions(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogShopTransactions, "[MonetizationShopTransactions.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogShopErrors(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogShopErrors, "[MonetizationShopErrors.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogEdicionPaquete(texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEdicionPaquete, "[EdicionPaquete.log] " & texto, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroServidor(texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogMacroServidor, "[MacroServidor] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroCliente(texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogMacroCliente, "[MacroCliente] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogVentaCasa, "[Propiedades] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogCriticEvent, "[Eventos.log] " & Desc, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEjercitoReal, "[EjercitoReal.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEjercitoCaos, "[EjercitoCaos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogError(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogError, "[Errores.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogPerformance(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogPerformance, "[Performance.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogConsulta(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogConsulta, "[obtenemos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogClanes(ByVal str As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogClanes, "[Clans.log] " & str, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub
Public Sub LogGM(Name As String, Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogGM, "[" & Name & "] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogPremios(GM As String, UserName As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)
On Error GoTo ErrHandler
        Dim s As String
        s = "Item: " & ObjData(ObjIndex).Name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine _
        & "Motivo: " & Motivo & vbNewLine & vbNewLine
        Call LogThis(type_log.e_LogPremios, s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogSecurity(str As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogSecurity, "[Cheating.log] " & str, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub TraceError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    'Start append text to file
    Dim filenum As Integer
    Dim msg As String
    msg = "Error number: " & Numero & " | Description: " & Descripcion & vbNewLine & "Component: " & Componente & " | Line number: " & Linea
    filenum = FreeFile
    Debug.Print msg
    Open App.Path & "\Logs\errores.log" For Append As filenum
    Print #filenum, msg
    Close filenum

End Sub
