Attribute VB_Name = "Application"
'RevolucionAo 1.0
'Pablo Mercavides
'**************************************************************************

Option Explicit

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type: Private HistorialError As UltimoError

Private Const EventFile = "\Argentum.log"

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    
    On Error GoTo IsAppActive_Err
    
    IsAppActive = (GetActiveWindow <> 0)

    
    Exit Function

IsAppActive_Err:
    Call RegistrarError(err.Number, err.Description, "Application.IsAppActive", Erl)
    Resume Next
    
End Function

Public Sub CleanEvents()

    If FileExist(App.path & EventFile, vbArchive) Then
        Kill App.path & EventFile
    End If

End Sub

Public Sub SetEvent(ByVal EventMessage As String)

Dim file As Integer: file = FreeFile
        
Open App.path & EventFile For Append As #file
    Print #file, Date$ & "-" & Time$ & ": " & EventMessage
Close #file

End Sub

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    '**********************************************************
    'Author: Jopi
    'Guarda una descripcion detallada del error en Errores.log
    '**********************************************************
        
        On Error GoTo RegistrarError_Err
    
        
        
        'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
100     If Componente = HistorialError.Componente And _
           Numero = HistorialError.ErrorCode Then
       
           'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
            'x lo que no hace falta registrar el error.
102        ' If HistorialError.Contador = 10 Then
           '     Debug.Print "Mismo error"
           '     Debug.Assert False
           '     Exit Sub
           ' End If
        
            'Agregamos el error al historial.
104         HistorialError.Contador = HistorialError.Contador + 1
        
        Else 'Si NO es igual, reestablecemos el contador.

106         HistorialError.Contador = 0
108         HistorialError.ErrorCode = Numero
110         HistorialError.Componente = Componente
            
        End If
        
        If Not FileExist(App.path & "\logs", vbDirectory) Then
            MkDir App.path & "\logs"
        End If
    
        'Registramos el error en Errores.log
112     Dim file As Integer: file = FreeFile
        
114     Open App.path & "\logs\Errores.log" For Append As #file
    
116         Print #file, "Error: " & Numero
118         Print #file, "Descripcion: " & Descripcion
        
120         Print #file, "Componente: " & Componente

122         If LenB(Linea) <> 0 Then
124             Print #file, "Linea: " & Linea
            End If

126         Print #file, "Fecha y Hora: " & Date$ & "-" & Time$
        
128         Print #file, vbNullString
        
130     Close #file
    
132     Debug.Print "Error: " & Numero & vbNewLine & _
                    "Descripcion: " & Descripcion & vbNewLine & _
                    "Componente: " & Componente & vbNewLine & _
                    "Linea: " & Linea & vbNewLine & _
                    "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
        
        Exit Sub

RegistrarError_Err:
        Call RegistrarError(err.Number, err.Description, "ES.RegistrarError", Erl)

        
End Sub

