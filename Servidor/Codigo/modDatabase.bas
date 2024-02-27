Attribute VB_Name = "modDatabase"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'

'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018
'Rewrited for Argentum by Alexis Caraballo (WyroX)
'October 2020

Option Explicit

Public Database_Enabled     As Boolean
Public Database_Driver      As String
Public Database_Source      As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Private Const MAX_ASYNC     As Byte = 20
Private Current_async       As Byte

Private Connection          As ADODB.Connection
Private Connection_async(1 To MAX_ASYNC)    As ADODB.Connection

Private Builder             As cStringBuilder

Public Function Query(ByVal Text As String, ParamArray Arguments() As Variant) As ADODB.Recordset
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    
    Command.ActiveConnection = Connection
    Command.CommandText = Text
    Command.CommandType = adCmdText
    Command.Prepared = True
    
    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument

    On Error GoTo Query_Err

    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call GetElapsedTime
    End If
    
    Set Query = Command.Execute()
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call LogPerformance("Query: " & Text & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If

    Exit Function
    
Query_Err:
    DBError = Err.Description
    'Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - " & vbCrLf & Text)
End Function

Public Function Execute(ByVal Text As String, ParamArray Arguments() As Variant) As Boolean
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    Exit Function
    Command.ActiveConnection = Connection_async(Current_async)
    Command.CommandText = Text
    Command.CommandType = adCmdText
    Command.Prepared = True

    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument
    
On Error GoTo Execute_Err
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call GetElapsedTime
    End If
    
    Call Command.Execute(, , adAsyncExecute)  ' @TODO: We want some operation to be async
    
    Current_async = Current_async + 1
    
    If Current_async = MAX_ASYNC Then
        Current_async = 1
    End If
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call LogPerformance("Execute: " & Text & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If
    
    Execute = (Err.Number = 0)
    Exit Function
        
Execute_Err:
    
    If (Err.Number <> 0) Then
     '   Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - " & vbCrLf & Text)
    End If
    
End Function


Private Function CreateParameter(ByVal value As Variant, ByVal Direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter
    Set CreateParameter = New ADODB.Parameter
    
    CreateParameter.Direction = Direction
    
    Select Case VarType(value)
        Case VbVarType.vbString
            CreateParameter.Type = adBSTR
            CreateParameter.size = Len(value)
            CreateParameter.value = CStr(value)
        Case VbVarType.vbDecimal
            CreateParameter.Type = adInteger
            CreateParameter.value = CLng(value)
        Case VbVarType.vbByte:
            CreateParameter.Type = adTinyInt
            CreateParameter.value = CByte(value)
        Case VbVarType.vbInteger
            CreateParameter.Type = adSmallInt
            CreateParameter.value = CInt(value)
        Case VbVarType.vbLong
            CreateParameter.Type = adInteger
            CreateParameter.value = CLng(value)
        Case VbVarType.vbBoolean
            CreateParameter.Type = adBoolean
            CreateParameter.value = CBool(value)
        Case VbVarType.vbSingle
            CreateParameter.Type = adSingle
            CreateParameter.value = CSng(value)
        Case VbVarType.vbDouble
            CreateParameter.Type = adDouble
            CreateParameter.value = CDbl(value)
    End Select
End Function


Public Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
        On Error GoTo ErrorHandler
    
100     Dim RS As ADODB.Recordset
        Set RS = Query("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE LOWER(" & ColumnaTest & ") = ?;", ValueTest)

        'Revisamos si recibio un resultado
102     If RS Is Nothing Then Exit Function
        If RS.BOF Or RS.EOF Then Exit Function
        
        'Obtenemos la variable
104     GetDBValue = RS.Fields(ColumnaGet).value

        Exit Function
    
ErrorHandler:
106     'Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.Description)
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)
        
        Exit Function

GetUserValue_Err:
102   '  Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserValue", Erl)
End Function

Public Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
        On Error GoTo ErrorHandler

        Call Execute("UPDATE " & Tabla & " SET " & ColumnaSet & " = ? WHERE " & ColumnaTest & " = ?;", ValueSet, ValueTest)

        Exit Sub
    
ErrorHandler:
102     'Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.Description)
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, value As Variant)
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, value, "UPPER(name)", UCase(CharName))

        Exit Sub

SetUserValue_Err:
102    ' Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValue", Erl)
End Sub

Private Sub SetUserValueByID(ByVal ID As Long, Columna As String, value As Variant)
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, value, "id", ID)

        Exit Sub

SetUserValueByID_Err:
102     'Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValueByID", Erl)
End Sub


Public Sub SaveVotoDatabase(ByVal ID As Long, ByVal Encuestas As Integer)
        
        On Error GoTo SaveVotoDatabase_Err
        
100     Call SetUserValueByID(ID, "votes_amount", Encuestas)
        
        Exit Sub

SaveVotoDatabase_Err:
102    ' Call TraceError(Err.Number, Err.Description, "modDatabase.SaveVotoDatabase", Erl)

        
End Sub


Public Function GetUserGuildIndexDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildIndexDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_index"), 0)

        Exit Function

ErrorHandler:
102    ' Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildMemberDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildMemberDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_member_history"), vbNullString)

        Exit Function

ErrorHandler:
102     'Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildAspirantDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildAspirantDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_aspirant_index"), 0)

        Exit Function

ErrorHandler:
102    ' Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildPedidosDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildPedidosDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_requests_history"), vbNullString)

        Exit Function

ErrorHandler:
102     'Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(UserName As String, Reason As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_rejected_because", Reason)

        Exit Sub
ErrorHandler:
102     'Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_index", GuildIndex)

        Exit Sub
ErrorHandler:
102    ' Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_aspirant_index", AspirantIndex)

        Exit Sub
ErrorHandler:
102    ' Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_member_history", guilds)

        Exit Sub
ErrorHandler:
102    ' Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_requests_history", Pedidos)

        Exit Sub
ErrorHandler:
102    ' Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim gName       As String

        Dim Miembro     As String

        Dim GuildActual As Integer

        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE UPPER(name) = ?;", UCase$(UserName))

102     If RS Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "Pj Inexistente", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Get the character's current guild
106     GuildActual = SanitizeNullValue(RS!Guild_Index, 0)

108     If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
110         gName = "<" & GuildName(GuildActual) & ">"
        Else
112         gName = "Ninguno"

        End If

        'Get previous guilds
114     Miembro = SanitizeNullValue(RS!guild_member_history, vbNullString)

116     If Len(Miembro) > 400 Then
118         Miembro = ".." & Right$(Miembro, 400)

        End If

120     Call WriteCharacterInfo(UserIndex, UserName, RS!race_id, RS!class_id, RS!genre_id, RS!level, RS!Gold, RS!bank_gold, SanitizeNullValue(RS!guild_requests_history, vbNullString), gName, Miembro, RS!pertenece_real, RS!pertenece_caos, RS!ciudadanos_matados, RS!criminales_matados)

        Exit Sub
ErrorHandler:
122    ' Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function SanitizeNullValue(ByVal value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(value), defaultValue, value)
        
        Exit Function

SanitizeNullValue_Err:
102  '   Call TraceError(Err.Number, Err.Description, "modDatabase.SanitizeNullValue", Erl)

        
End Function


