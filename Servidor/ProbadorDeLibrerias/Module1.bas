Attribute VB_Name = "Module1"
Option Explicit

Sub main()
On Error GoTo fallo
    Dim etapa As String
    
    etapa = "msado15.dll"
    
    Dim test1 As ADODB.Connection
    
    etapa = "scrrun.dll"
    
    Dim test2 As Screen
    
    etapa = "msxml6.dll"
    
    Dim test3 As XMLHTTP60
    
    etapa = "Aurora.Network.dll"
    
    Dim test4 As Network.Server
    
    etapa = "MSINET.ocx"
    
    Dim test5 As Inet
    
    etapa = "MSDatGrd.ocx"
    
    Dim test6 As DataGrid
    
    etapa = "AOProgress.ocx"
    
    Dim test7 As uAOProgress

    MsgBox "Todo OK", vbInformation + vbOKOnly

Exit Sub
fallo:

MsgBox "Error en " & etapa & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbCritical + vbOKOnly

End Sub

