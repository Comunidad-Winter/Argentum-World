Attribute VB_Name = "modDBManager"
Option Explicit

Private DBClient As Network.Client
Private DBWriter As Network.Writer
Private DBConnected As Boolean
Private Enum DBCommands
    LoadCharacter = 1
    SaveCharacter = 2
    CreateAccount = 3
    BanCharacter = 4
    UnbanCharacter = 5
    SilenceChar = 6
    AddPenalty = 7
    AlterName = 8
    LoginAccount = 9
    GetUserPenalties = 10
    TransferGold = 11
    DepositGold = 12
    WithdrawGold = 13
    KickFaction = 14
    BankGold = 15
    CharacterEvent = 16
    GetLastIp = 17
    
    ' guilds
    ListGuilds = 18
    GetGuild = 19
    SetGuild = 20
    AddMemberRequest = 21
    ListMemberRequests = 22
    EventMemberRequest = 23
    RemoveMember = 24
    ListMembers = 25
    EventElection = 26
    GetVotes = 27
    ListPropositions = 28
    ListRelations = 29
    NewGuildProposition = 30
End Enum

Public Enum e_EventType
    PickGold = 1
    DropGold = 2
    PickItem = 3
    DropItem = 4
    TransferGold = 5
    BuyItem = 6
    SellItem = 7
    P2PCommerce = 8
    KillNpc = 9
    ForgeItem = 10
    LevelUp = 11
    BurnSword = 12
    KillUser = 13
End Enum
Public Sub DBInitialize()
    Set DBWriter = New Network.Writer
End Sub

Public Sub DBClear()
    DBConnected = False
    Call DBWriter.Clear
End Sub

Public Function DBIsConnected() As Boolean
    DBIsConnected = DBConnected
End Function


Public Sub DBConnect(ByVal Address As String, ByVal Service As String)
    On Error GoTo Handler:

    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    
    Call DBInitialize
    
    Set DBClient = New Network.Client
    Call DBClient.Attach(AddressOf OnDBClientConnect, AddressOf OnDBClientClose, AddressOf OnDBClientSend, AddressOf OnDBClientRecv)
    Call DBClient.Connect(Address, Service)

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en DBConnect. ", Erl)
End Sub

Public Sub DBDisconnect()
    On Error GoTo Handler:
    If Not DBClient Is Nothing Then
        Call DBClient.Close(True)
    End If
    
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en DBDisconnect. ", Erl)
End Sub

Public Sub DBPoll()

    On Error GoTo Handler:

    If (DBClient Is Nothing) Then
        Exit Sub
    End If
    
    Call DBClient.Flush
    Call DBClient.Poll

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en DBPoll. ", Erl)
End Sub

Public Sub DBSend(ByVal Buffer As Network.Writer)
    On Error GoTo Handler:
    
    If (DBConnected) Then
        Call DBClient.Send(True, Buffer)
    End If
    
    Call Buffer.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en DBSend. ", Erl)
End Sub

    
Private Sub OnDBClientConnect()
    On Error GoTo Handler:
    Debug.Print ("Conecto al DBManager")

    DBConnected = True
    
    Exit Sub
    
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en OnDBClientConnect. ", Erl)
End Sub

Private Sub OnDBClientClose(ByVal code As Long)
On Error GoTo Handler:
    Debug.Print ("Closed " & code)
    Call DBClear


    Exit Sub
    
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en OnDBClientClose. ", Erl)
End Sub

Private Sub OnDBClientSend(ByVal message As Network.Reader)
On Error GoTo Handler:
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en OnDBClientSend. ", Erl)
End Sub

Private Sub OnDBClientRecv(ByVal Reader As Network.Reader)
On Error GoTo Handler:

    Dim Command As Integer
    Dim UserIndex As Integer
    Dim Validator As Long
    
    Command = Reader.ReadInt16
    UserIndex = Reader.ReadInt16()
    Validator = Reader.ReadInt32()
    If UserIndex < 0 Or UserIndex > 10000 Then
        Err.raise 1, , "Invalid UserIndex out of range"
        Exit Sub
    End If
    If UserIndex > 0 Then
        If UserList(UserIndex).Validator <> Validator Then
            Err.raise 1, , "Invalid Validator. UserIndex: " & UserIndex & ", Validator:" & UserList(UserIndex).Validator & ", Received: " & Validator
            Exit Sub
        End If
    End If
    
    Select Case Command
        Case DBCommands.LoadCharacter
            Call ReceiveLoadCharacter(UserIndex, Reader)
        Case DBCommands.SaveCharacter
            Call ReceiveSaveCharacter(UserIndex, Reader)
        Case DBCommands.BanCharacter
            Call ReceiveBanCharacter(UserIndex, Reader)
        Case DBCommands.UnbanCharacter
            Call ReceiveUnbanCharacter(UserIndex, Reader)
        Case DBCommands.AlterName
            Call ReceiveAlterName(UserIndex, Reader)
        Case DBCommands.LoginAccount
            Call ReceiveLoginAccount(UserIndex, Reader)
        Case DBCommands.TransferGold
            Call ReceiveTransferGold(UserIndex, Reader)
        Case DBCommands.KickFaction
            Call ReceiveKickFaction(UserIndex, Reader)
        Case DBCommands.GetUserPenalties
            Call ReceiveGetUserPenalties(UserIndex, Reader)
        Case DBCommands.WithdrawGold
            Call ReceiveWithdrawGold(UserIndex, Reader)
        Case DBCommands.DepositGold
            Call ReceiveDepositGold(UserIndex, Reader)
        Case DBCommands.BankGold
            Call ReceiveBankGold(UserIndex, Reader)
        Case DBCommands.CreateAccount
            Call ReceiveCreateAccount(UserIndex, Reader)
        Case DBCommands.GetLastIp
            Call ReceiveGetLastIP(UserIndex, Reader)
            
        ' guilds
        Case DBCommands.ListGuilds
            Call ReceiveListGuilds(UserIndex, Reader)
'        Case DBCommands.GetGuild
'            Call ReceiveGetGuild(UserIndex, Reader)
'        Case DBCommands.SetGuild
'            Call ReceiveSetGuild(UserIndex, Reader)
'        Case DBCommands.AddMemberRequest
'            Call ReceiveAddMemberRequest(UserIndex, Reader)
'        Case DBCommands.ListMemberRequests
'            Call ReceiveListMemberRequests(UserIndex, Reader)
'        Case DBCommands.EventMemberRequest
'            Call ReceiveEventMemberRequest(UserIndex, Reader)
'        Case DBCommands.RemoveMember
'            Call ReceiveRemoveMember(UserIndex, Reader)
'        Case DBCommands.ListMembers
'            Call ReceiveListMembers(UserIndex, Reader)
'        Case DBCommands.EventElection
'            Call ReceiveEventElection(UserIndex, Reader)
'        Case DBCommands.GetVotes
'            Call ReceiveGetVotes(UserIndex, Reader)
'        Case DBCommands.ListPropositions
'            Call ReceiveListPropositions(UserIndex, Reader)
'        Case DBCommands.ListRelations
'            Call ReceiveListRelations(UserIndex, Reader)
'        Case DBCommands.NewGuildProposition
'            Call ReceiveNewGuildProposition(UserIndex, Reader)
    End Select

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en OnDBClientRecv. ", Erl)
End Sub

Public Sub SendLoadCharacter(ByVal UserIndex As Integer, ByVal AccountId As Integer, ByVal Name As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.LoadCharacter)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (AccountId)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteString8 (UserList(UserIndex).IP)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendLoadCharacter. ", Erl)
End Sub
Public Sub SendBanCharacter(ByVal UserIndex As Integer, ByVal Name As String, ByVal Reason As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.BanCharacter)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).ID)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteString8 (Reason)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendBanCharacter. ", Erl)
End Sub
Public Sub SendSilenceChar(ByVal UserIndex As Integer, ByVal Name As String, ByVal Time As Integer)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.SilenceChar)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).ID)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteInt16 (Time)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendSilenceChar. ", Erl)
End Sub

Public Sub SendAddPenalty(ByVal UserIndex As Integer, ByVal Name As String, ByVal jailTime As Integer, ByVal Reason As String)
    On Error GoTo Handler:
    
    Dim typePe As Byte
    If jailTime = 0 Then
        typePe = 1
    ElseIf jailTime > 0 Then
        typePe = 3
    End If
    
    DBWriter.WriteInt16 (DBCommands.AddPenalty)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).ID)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteString8 (Reason)
    DBWriter.WriteInt8 (typePe)
    DBWriter.WriteInt16 (jailTime)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendAddPenalty. ", Erl)
End Sub

Public Sub SendCharacterEvent(ByVal UserIndex As Integer, ByVal TypeId As Integer, ByVal Description As String, ByVal EntityId As Long, ByVal Amount As Long, ByVal value As Long)
    On Error GoTo Handler:
            
    DBWriter.WriteInt16 (DBCommands.CharacterEvent)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).ID)
    DBWriter.WriteString8 (UserList(UserIndex).Name)
    DBWriter.WriteInt8 (TypeId)
    DBWriter.WriteString8 (Description)
    DBWriter.WriteInt32 (EntityId)
    DBWriter.WriteInt32 (Amount)
    DBWriter.WriteInt32 (value)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendAddCharacterEvent. ", Erl)
End Sub

Public Sub SendTransferGold(ByVal UserIndex As Integer, ByVal Name As String, ByVal Amount As Long)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.TransferGold)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).AccountId)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteInt32 (Amount)
    DBWriter.WriteInt16 (UserList(UserIndex).flags.TargetNPC)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendTransferGold. ", Erl)
End Sub

Public Sub SendDepositGold(ByVal UserIndex As Integer, ByVal Amount As Long)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.DepositGold)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).AccountId)
    DBWriter.WriteInt32 (Amount)
    DBWriter.WriteInt16 (UserList(UserIndex).flags.TargetNPC)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendDepositGold. ", Erl)
End Sub

Public Sub SendWithdrawGold(ByVal UserIndex As Integer, ByVal Amount As Long)
    On Error GoTo Handler:
        
    DBWriter.WriteInt16 (DBCommands.WithdrawGold)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).AccountId)
    DBWriter.WriteInt32 (Amount)
    DBWriter.WriteInt16 (UserList(UserIndex).flags.TargetNPC)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendWithdrawGold. ", Erl)
End Sub

Public Sub SendBankGold(ByVal UserIndex As Integer, Optional ByVal InitBank As Byte = 0)
    On Error GoTo Handler:
        
    DBWriter.WriteInt16 (DBCommands.BankGold)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).AccountId)
    DBWriter.WriteInt16 (UserList(UserIndex).flags.TargetNPC)
    DBWriter.WriteInt8 (InitBank)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendBankGold. ", Erl)
End Sub

Public Sub SendGetLastIP(ByVal UserIndex As Integer, ByVal Name As String)
    On Error GoTo Handler:
        
    DBWriter.WriteInt16 (DBCommands.GetLastIp)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteString8 (Name)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendGetLastIP. ", Erl)
End Sub

Public Sub SendKickFaction(ByVal UserIndex As Integer, ByVal Name As String, ByVal Faction As Byte)
    On Error GoTo Handler:
        
    DBWriter.WriteInt16 (DBCommands.KickFaction)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteInt8 (Faction)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendKickFaction. ", Erl)
End Sub

Public Sub SendUnbanCharacter(ByVal UserIndex As Integer, ByVal Name As String, ByVal Reason As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.UnbanCharacter)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).ID)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteString8 (Reason)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendUnbanCharacter. ", Erl)
End Sub

Public Sub SendLoginAccount(ByVal UserIndex As Integer, ByVal Email As String, ByVal Password As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.LoginAccount)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteString8 (Email)
    DBWriter.WriteString8 (Password)
    DBWriter.WriteString8 (UserList(UserIndex).IP)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendLoginAccount. ", Erl)
End Sub

Public Sub SendCreateAccount(ByVal UserIndex As Integer, ByVal Email As String, ByVal Password As String, ByVal FirstName As String, ByVal LastName As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.CreateAccount)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteString8 (Email)
    DBWriter.WriteString8 (Password)
    DBWriter.WriteString8 (FirstName)
    DBWriter.WriteString8 (LastName)
    DBWriter.WriteString8 (UserList(UserIndex).IP)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendCreateAccount. ", Erl)
End Sub

Public Sub SendGetUserPenalties(ByVal UserIndex As Integer, ByVal UserName As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.GetUserPenalties)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteString8 (UserName)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendGetUserPenalties. ", Erl)
End Sub

Public Sub SendAlterName(ByVal UserIndex As Integer, ByVal Name As String, ByVal NewName As String)
    On Error GoTo Handler:
    
    DBWriter.WriteInt16 (DBCommands.AlterName)
    DBWriter.WriteInt16 (UserIndex)
    DBWriter.WriteInt32 (UserList(UserIndex).Validator)
    DBWriter.WriteInt32 (UserList(UserIndex).ID)
    DBWriter.WriteString8 (Name)
    DBWriter.WriteString8 (NewName)
    Call DBClient.Send(False, DBWriter)
    DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendAlterName. ", Erl)
End Sub

Private Sub ReceiveTransferGold(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim UserName As String
    Dim TargetUI As Integer
    Dim tUI As Integer
    Dim Amount As Long
    Dim Saldo As Long

    Status = Reader.ReadInt8()
    
    TargetUI = UserList(UserIndex).flags.TargetNPC

    If Status = 1 Then
    
        UserName = Reader.ReadString8()
        Amount = Reader.ReadInt32()
        Saldo = Reader.ReadInt32()
        TargetUI = Reader.ReadInt16()
        'TODO: Registrar operacion en los logs.
    
        tUI = NameIndex(UserName)
        If tUI > 0 Then Call WriteConsoleMsg(tUI, "El usuario " & UserList(UserIndex).Name & " te ha enviado " & Amount & " monedas de oro a tu cuenta de banco.", e_FontTypeNames.FONTTYPE_INFO)
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "¡El envío se ha realizado con éxito! Te queda un saldo de " & Saldo & " monedas. Gracias por utilizar los servicios de Finanzas Goliath", NpcList(TargetUI).Char.charindex, vbWhite)
                
    ElseIf Status = 2 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "No se ha encontrado al usuario en nuestros registros..", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 3 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "No tienes suficiente oro para realizar la operación.", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 4 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "¡No puedo enviarte oro a vos mismo!", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveAlterName. ", Erl)
End Sub

Private Sub ReceiveBankGold(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim TargetUI As Integer
    Dim Amount As Long
    Dim Saldo As Long
    Dim InitBank As Byte

    Status = Reader.ReadInt8()
    
    TargetUI = UserList(UserIndex).flags.TargetNPC

    If Status = 1 Then
    
        Saldo = Reader.ReadInt32()
        TargetUI = Reader.ReadInt16()
        InitBank = Reader.ReadInt8()
        
        If InitBank = 1 Then
            Call WriteGoliathInit(UserIndex, Saldo)
        Else
            If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "Tenés " & Saldo & " monedas de oro en tu cuenta.", NpcList(TargetUI).Char.charindex, vbWhite)
        End If
    ElseIf Status = 2 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "No se ha encontrado al usuario en nuestros registros.", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveBankGold. ", Erl)
End Sub


Private Sub ReceiveWithdrawGold(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim TargetUI As Integer
    Dim Amount As Long
    Dim Saldo As Long

    Status = Reader.ReadInt8()
    
    TargetUI = UserList(UserIndex).flags.TargetNPC

    If Status = 1 Then
    
        Amount = Reader.ReadInt32()
        Saldo = Reader.ReadInt32()
        TargetUI = Reader.ReadInt16()
        'TODO: Registrar operacion en los logs.
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Amount
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGld(UserIndex, Saldo)
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "Tenés " & Saldo & " monedas de oro en tu cuenta.", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 2 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "No se ha encontrado al usuario en nuestros registros.", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 3 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveWithdrawGold. ", Erl)
End Sub

Private Sub ReceiveDepositGold(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim TargetUI As Integer
    Dim Amount As Long
    Dim Saldo As Long

    Status = Reader.ReadInt8()
    
    TargetUI = UserList(UserIndex).flags.TargetNPC

    If Status = 1 Then
    
        Amount = Reader.ReadInt32()
        Saldo = Reader.ReadInt32()
        TargetUI = Reader.ReadInt16()
        'TODO: Registrar operacion en los logs.
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "Tenés " & Saldo & " monedas de oro en tu cuenta.", NpcList(TargetUI).Char.charindex, vbWhite)
        Call WriteUpdateBankGld(UserIndex, Saldo)
    ElseIf Status = 2 Then
        If TargetUI > 0 Then Call WriteChatOverHead(UserIndex, "No se ha encontrado al usuario en nuestros registros.", NpcList(TargetUI).Char.charindex, vbWhite)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveWithdrawGold. ", Erl)
End Sub

Private Sub ReceiveGetUserPenalties(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim txt As String
    Dim Amount As Long

    Status = Reader.ReadInt8()
    

    If Status = 1 Then
            
        Amount = Reader.ReadInt16()
        
        For i = 1 To Amount
            txt = Reader.ReadString8()
            Call WriteConsoleMsg(UserIndex, i & ". " & txt, e_FontTypeNames.FONTTYPE_INFO)
        Next i
        
    ElseIf Status = 2 Then
        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveGetUserPenalties. ", Erl)
End Sub

Private Sub ReceiveKickFaction(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim UserName As String
    Dim TargetUI As Integer
    Dim tUI As Integer
    Dim Amount As Long

    Status = Reader.ReadInt8()
    
    TargetUI = UserList(UserIndex).flags.TargetNPC
    
    If Status = 1 Then
    
        UserName = Reader.ReadString8()

        Call WriteConsoleMsg(UserIndex, "Usuario " & UserName & " expulsado correctamente.", e_FontTypeNames.FONTTYPE_INFO)
        
    ElseIf Status = 2 Then
        Call WriteConsoleMsg(UserIndex, "No existe el personaje.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 3 Then
        Call WriteConsoleMsg(UserIndex, "El personaje no pertenece a la facción.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveKickFaction. ", Erl)
End Sub

Private Sub ReceiveGetLastIP(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim Status As Byte
    Dim CantIps As Integer


    Status = Reader.ReadInt8()
    
    
    If Status = 1 Then
    
        CantIps = Reader.ReadInt16()
        
        
        Call WriteConsoleMsg(UserIndex, "Las últimas ips para el personaje son: ", e_FontTypeNames.FONTTYPE_INFO)
        For i = 1 To CantIps
            Call WriteConsoleMsg(UserIndex, Reader.ReadString8(), e_FontTypeNames.FONTTYPE_INFO)
        Next i
        
        
    ElseIf Status = 2 Then
        Call WriteConsoleMsg(UserIndex, "No existe el personaje.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveGetLastIP. ", Erl)
End Sub

Private Sub ReceiveCreateAccount(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer

    Dim Status As Byte
    Dim UserName As String
    Dim AccountId As Integer

    Status = Reader.ReadInt8()
    
    If UserIndex = 0 Then Exit Sub
    
    If Status = 1 Then 'Cuenta creada
        UserName = Reader.ReadString8()
        AccountId = Reader.ReadInt32()
        
        UserList(UserIndex).AccountId = AccountId

        Dim Personajes() As t_PersonajeCuenta
        Call WriteAccountCharacterList(UserIndex, Personajes, 0)
        
        'TODO: Agregar protecciones para evitar spams.
    ElseIf Status = 2 Then
        Call WriteErrorMsg(UserIndex, "Ya hay una cuenta asociada con ese email.")
        Call CloseSocket(UserIndex)
    ElseIf Status = 255 Then
        Call WriteErrorMsg(UserIndex, "Ocurrió un error al realizar la operación.")
        Call CloseSocket(UserIndex)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveCreateAccount. ", Erl)
End Sub

Private Sub ReceiveAlterName(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim tmpLength As Integer
    Dim CharacterId As Long
    Dim Status As Byte
    Dim UserName As String
    Dim NewName As String
    Dim TargetUI As Integer

    
    Status = Reader.ReadInt8()
    
    If UserIndex = 0 Then Exit Sub
    
    If Status = 1 Then 'Baneado
        UserName = Reader.ReadString8()
        NewName = Reader.ReadString8()
        
        Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", e_FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración Â» " & UserList(UserIndex).Name & " cambió el nombre del usuario """ & UserName & """ por """ & NewName & """.", e_FontTypeNames.FONTTYPE_GM))
        Call LogGM(UserList(UserIndex).Name, "Ha cambiado de nombre al usuario """ & UserName & """. Ahora se llama """ & NewName & """.")
        
        TargetUI = NameIndex(UserName)
        If TargetUI > 0 Then
            UserList(TargetUI).Name = NewName
            Call RefreshCharStatus(TargetUI)
        End If

    ElseIf Status = 2 Then
        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_TALK)
    ElseIf Status = 3 Then
        Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveAlterName. ", Erl)
End Sub

Private Sub ReceiveLoginAccount(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim tmpLength As Integer
    Dim CharacterId As Long
    Dim Status As Byte
    Dim UserName As String
    Dim NewName As String
    Dim TargetUI As Integer

    
    Status = Reader.ReadInt8()
    
    If Status = 1 Then
        UserList(UserIndex).AccountId = Reader.ReadInt32()
        
        Dim Personajes(1 To 10) As t_PersonajeCuenta
        Dim Count As Byte
        
        Count = Reader.ReadInt8()
        
        For i = 1 To Count
            Personajes(i).nombre = Reader.ReadString8()
            Personajes(i).Cabeza = Reader.ReadInt16()
            Personajes(i).clase = Reader.ReadInt16()
            Personajes(i).cuerpo = Reader.ReadInt16()
            Personajes(i).Mapa = Reader.ReadInt16()
            Personajes(i).PosX = Reader.ReadInt16()
            Personajes(i).PosY = Reader.ReadInt16()
            Personajes(i).nivel = Reader.ReadInt16()
            Personajes(i).Status = Reader.ReadInt16()
            Personajes(i).Casco = Reader.ReadInt16()
            Personajes(i).Escudo = Reader.ReadInt16()
            Personajes(i).Arma = Reader.ReadInt16()
            Personajes(i).ClanIndex = Reader.ReadInt16()
            Personajes(i).Privs = UserDarPrivilegioLevel(Personajes(i).nombre)
        

        Next i
        
        Call WriteAccountCharacterList(UserIndex, Personajes, Count)
        
    ElseIf Status = 2 Then
        Call WriteShowMessageBox(UserIndex, "Usuario o Contraseña erronea.")
        Call CloseSocket(UserIndex)
        Exit Sub
    ElseIf Status = 255 Then
        Call WriteShowMessageBox(UserIndex, "Ocurrió un error al realizar la operación.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
   
    
    Exit Sub
Handler:
    Call WriteShowMessageBox(UserIndex, "Error desconocido.")
    Call CloseSocket(UserIndex)
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveLoginAccount. ", Erl)
End Sub

Private Sub ReceiveBanCharacter(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim tmpLength As Integer
    Dim CharacterId As Long
    Dim Status As Byte
    Dim Name As String
    Dim Reason As String

    
    Status = Reader.ReadInt8()
    
    If UserIndex = 0 Then Exit Sub
    
    If Status = 1 Then 'Baneado
        Name = Reader.ReadString8()
        Reason = Reader.ReadString8()
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor Â» " & UserList(UserIndex).Name & " ha baneado a " & Name & " debido a: " & Reason & ".", e_FontTypeNames.FONTTYPE_SERVER))
    ElseIf Status = 2 Then
        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_TALK)
    ElseIf Status = 3 Then
        Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra baneado.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveBanCharacter. ", Erl)
End Sub

Private Sub ReceiveUnbanCharacter(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim tmpLength As Integer
    Dim CharacterId As Long
    Dim Status As Byte
    Dim Name As String
    Dim Reason As String


    Status = Reader.ReadInt8()
    
    If UserIndex = 0 Then Exit Sub
    
    If Status = 1 Then 'Unbaneado
        Name = Reader.ReadString8()
        Reason = Reader.ReadString8()
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor Â» " & UserList(UserIndex).Name & " ha desbaneado a " & Name & " debido a: " & Reason & ".", e_FontTypeNames.FONTTYPE_SERVER))
    ElseIf Status = 2 Then
        Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_TALK)
    ElseIf Status = 3 Then
        Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra baneado.", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf Status = 255 Then
        Call WriteConsoleMsg(UserIndex, "Ocurrió un error al realizar la operación.", e_FontTypeNames.FONTTYPE_INFO)
    End If

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveUnbanCharacter. ", Erl)
End Sub

Public Sub SendSaveCharacter(ByVal UserIndex As Integer, ByVal Logout As Boolean, Optional ByVal NewUser As Boolean = False)
    On Error GoTo Handler:

    Dim LoopC As Integer
    Dim Counter As Integer
    
    With UserList(UserIndex)
        DBWriter.WriteInt16 (DBCommands.SaveCharacter)
        DBWriter.WriteInt16 (UserIndex)
        DBWriter.WriteInt32 (.Validator)
        DBWriter.WriteInt32 (.AccountId)
        DBWriter.WriteInt32 (.ID)
        DBWriter.WriteInt32 (.LoginId)
        If Logout Then
            DBWriter.WriteInt8 (1)
        ElseIf NewUser Then
            DBWriter.WriteInt8 (2)
        Else
            DBWriter.WriteInt8 (0)
        End If
        
        DBWriter.WriteString8 (.IP)
        DBWriter.WriteString8 (.Name)
        
        DBWriter.WriteInt8 (.Stats.ELV)
        DBWriter.WriteInt32 (.Stats.Exp)
        DBWriter.WriteInt8 (.genero)
        DBWriter.WriteInt8 (.raza)
        DBWriter.WriteInt8 (.clase)
        DBWriter.WriteInt8 (.Hogar)
        DBWriter.WriteString8 (.Desc)
        DBWriter.WriteInt32 (.Stats.GLD)
        DBWriter.WriteInt16 (.Stats.SkillPts)
        DBWriter.WriteInt16 (.Pos.Map)
        DBWriter.WriteInt16 (.Pos.X)
        DBWriter.WriteInt16 (.Pos.Y)
        DBWriter.WriteInt16 (.OrigChar.Body)
        DBWriter.WriteInt16 (.OrigChar.Head)
        DBWriter.WriteInt16 (.Char.Body)
        DBWriter.WriteInt16 (.Char.WeaponAnim)
        DBWriter.WriteInt16 (.Char.CascoAnim)
        DBWriter.WriteInt16 (.Char.ShieldAnim)
        DBWriter.WriteInt16 (.Invent.BarcoObjIndex)
        DBWriter.WriteInt8 (.Char.Heading)
        DBWriter.WriteInt16 (.Stats.MinHp)
        DBWriter.WriteInt16 (.Stats.MaxHp)
        DBWriter.WriteInt16 (.Stats.MinMAN)
        DBWriter.WriteInt16 (.Stats.MaxMAN)
        DBWriter.WriteInt16 (.Stats.MinSta)
        DBWriter.WriteInt16 (.Stats.MaxSta)
        DBWriter.WriteInt16 (.Stats.MinHam)
        DBWriter.WriteInt16 (.Stats.MaxHam)
        DBWriter.WriteInt16 (.Stats.MinAGU)
        DBWriter.WriteInt16 (.Stats.MaxAGU)
        DBWriter.WriteInt16 (.Stats.MinHIT)
        DBWriter.WriteInt16 (.Stats.MaxHit)
        DBWriter.WriteInt8 (.Stats.InventLevel)
        DBWriter.WriteInt8 (.Stats.VaultLevel)
        DBWriter.WriteInt16 (.GuildIndex)
        DBWriter.WriteInt16 (.flags.ReturnPos.Map)
        DBWriter.WriteInt16 (.flags.ReturnPos.X)
        DBWriter.WriteInt16 (.flags.ReturnPos.Y)
        DBWriter.WriteInt8 (.Faccion.Status)
        DBWriter.WriteInt8 (.Faccion.RecompensasReal)
        DBWriter.WriteInt32 (.Faccion.MatadosIngreso)
        DBWriter.WriteInt32 (.Stats.NPCsMatados)
        DBWriter.WriteInt32 (.Stats.UsuariosMatados)
        DBWriter.WriteInt32 (.Stats.ciudadanosMatados)
        DBWriter.WriteInt32 (.Stats.CriminalesMatados)
        DBWriter.WriteInt32 (.Stats.MuertesPorNpcs)
        DBWriter.WriteInt32 (.Stats.MuertesPorUsers)
        DBWriter.WriteInt32 (.Stats.MuertesTotales)
        DBWriter.WriteInt32 (.Stats.pasos)
        DBWriter.WriteInt8 (.flags.Envenenado)
        DBWriter.WriteInt8 (.flags.Incinerado)
        DBWriter.WriteInt8 (.flags.Navegando)
        DBWriter.WriteInt8 (.flags.Paralizado)
        DBWriter.WriteInt8 (.flags.Silenciado)
        DBWriter.WriteInt8 (.flags.Montado)
        DBWriter.WriteInt8 (.ChatGlobal)
        DBWriter.WriteInt8 (.ChatCombate)
        DBWriter.WriteInt32 (.Stats.PuntosPesca)
        DBWriter.WriteInt32 (.Stats.ELO)
        DBWriter.WriteInt32 (.Stats.TiempoJugado)
        DBWriter.WriteInt16 (.Counters.Pena)
        DBWriter.WriteInt16 (.flags.MinutosRestantes)
        
        ' ************************** User spells *********************************
        For LoopC = 1 To MAXUSERHECHIZOS
            If .Stats.UserHechizos(LoopC) > 0 Then
                Counter = Counter + 1
            End If
        Next LoopC
        
        DBWriter.WriteInt8 (Counter)
        For LoopC = 1 To MAXUSERHECHIZOS
            If .Stats.UserHechizos(LoopC) > 0 Then
                DBWriter.WriteInt8 (LoopC)
                DBWriter.WriteInt8 (.Stats.UserHechizos(LoopC))
            End If
        Next LoopC
    
            
        ' ************************** User inventory *********************************
        DBWriter.WriteInt8 (MAX_INVENTORY_SLOTS)
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            DBWriter.WriteInt16 (.Invent.Object(LoopC).ObjIndex)
            DBWriter.WriteInt16 (.Invent.Object(LoopC).Amount)
            DBWriter.WriteInt8 (.Invent.Object(LoopC).Equipped)
        Next LoopC

        ' ************************** User bank inventory *********************************
        
        DBWriter.WriteInt8 (MAX_BANCOINVENTORY_SLOTS)
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            DBWriter.WriteInt16 (.BancoInvent.Object(LoopC).ObjIndex)
            DBWriter.WriteInt16 (.BancoInvent.Object(LoopC).Amount)
        Next LoopC
        

        ' ************************** User skills *********************************
        DBWriter.WriteInt8 (NUMSKILLS)
        For LoopC = 1 To NUMSKILLS
            DBWriter.WriteInt8 (.Stats.UserSkills(LoopC))
            DBWriter.WriteInt8 (.Stats.UserSkillsAssigned(LoopC))
        Next LoopC
        



        ' ************************** User quests *********************************
        Counter = 0
        For LoopC = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                Counter = Counter + 1
            End If
        Next LoopC
        DBWriter.WriteInt8 (Counter)
        
        Dim Tmp As Integer, LoopK As Long
        Dim TempStr As String
        
        For LoopC = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs
                TempStr = ""
                If Tmp Then
                
                    For LoopK = 1 To Tmp
                        TempStr = TempStr & CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                        
                        If LoopK < Tmp Then
                            TempStr = TempStr & "-"
                        End If
                    
                    Next LoopK
                    
                    
                End If
                TempStr = TempStr & "|"
                Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                        
                For LoopK = 1 To Tmp
                
                    TempStr = TempStr & CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                    
                    If LoopK < Tmp Then
                        TempStr = TempStr & "-"
                    End If
                
                Next LoopK
                
                DBWriter.WriteInt16 (.QuestStats.Quests(LoopC).QuestIndex)
                DBWriter.WriteString8 (TempStr)
            End If
        Next LoopC
                        
        ' ************************** User completed quests *********************************
        DBWriter.WriteInt16 (.QuestStats.NumQuestsDone)
        If .QuestStats.NumQuestsDone > 0 Then
            For LoopC = 1 To .QuestStats.NumQuestsDone
                DBWriter.WriteInt16 (.QuestStats.QuestsDone(LoopC))
                
            Next LoopC
        End If


    End With
    Call DBClient.Send(False, DBWriter)
    Call DBWriter.Clear
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en SendSaveCharacter. ", Erl)
End Sub
Private Sub ReceiveSaveCharacter(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim tmpLength As Integer
    Dim AccountId As Long
    Dim CharacterId As Long
    Dim Status As Byte
    
    Status = Reader.ReadInt8()
    
    If Status = 3 Then
        Call WriteShowMessageBox(UserIndex, "El nombre de personaje ya existe.")
        Call CloseSocket(UserIndex)
        Exit Sub
    ElseIf Status = 4 Then
        Call WriteShowMessageBox(UserIndex, "Has llegado al limite de personajes por cuenta.")
        Call CloseSocket(UserIndex)
        Exit Sub
    ElseIf Status = 255 Then
        Call WriteShowMessageBox(UserIndex, "Ocurrió un error al realizar la operación.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    
    AccountId = Reader.ReadInt32()
    CharacterId = Reader.ReadInt32()
    

    
    If Status = 1 And UserIndex > 0 Then
        'Call SaveCreditsDatabase(userindex)
        If Not dcnUsersLastLogout.Exists(UCase(UserList(UserIndex).Name)) Then
            Call dcnUsersLastLogout.Add(UCase(UserList(UserIndex).Name), GetTickCount())
        End If
    End If
    
    If Status = 2 Then
        Call ConnectUser(UserIndex, UserList(UserIndex).Name, False)
    End If

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveSaveCharacter. ", Erl)
End Sub
Private Sub ReceiveLoadCharacter(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:
    
    Dim i As Integer
    Dim tmpLength As Integer
    Dim Slot As Integer
    Dim Locked As Byte
    Dim Banned As Byte
    Dim LoopC As Integer
    Dim data As String
    Dim ObjIndex As Integer
    Dim Status As Byte
    
    If UserIndex = 0 Then Exit Sub
    
    
    With UserList(UserIndex)
        
        
        Status = Reader.ReadInt8()
    
     
        If Status = 2 Then 'No existe..
            Call WriteShowMessageBox(UserIndex, "El personaje no existe.")
            Call CloseSocket(UserIndex)
            Exit Sub
        ElseIf Status = 255 Then 'Error
            Call WriteShowMessageBox(UserIndex, "Ocurrió un error desconocido.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        .ID = Reader.ReadInt32()
        .AccountId = Reader.ReadInt32()
        .Name = Reader.ReadString8()
        
        Locked = Reader.ReadInt8()
        Banned = Reader.ReadInt8()
        
        
        If (Locked) Then
             Call WriteShowMessageBox(UserIndex, "El personaje que estás intentando loguear se encuentra en venta, para desbloquearlo deberás hacerlo desde la página web.")
        
             Call CloseSocket(UserIndex)
             Exit Sub
         End If
        
        If (Banned) Then
             Dim BanNick     As String
             Dim BaneoMotivo As String
             Dim Expires As String
             Dim Created As String
             
             BaneoMotivo = Reader.ReadString8()
             BanNick = Reader.ReadString8()
             Created = Reader.ReadInt32()
             Expires = Reader.ReadInt32()
             
             If LenB(BanNick) = 0 Then BanNick = "*Error en la base de datos*"
             If LenB(BaneoMotivo) = 0 Then BaneoMotivo = "*No se registra el motivo del baneo.*"
             If Expires = 0 Then
                Expires = "El baneo es definitivo."
             Else
                Expires = "El baneo será hasta el " & Expires
             End If
         
             Call WriteShowMessageBox(UserIndex, "Se te ha prohibido la entrada al juego el " & Created & " debido a " & BaneoMotivo & ". Esta decisión fue tomada por " & BanNick & "." & Expires & ".")
         
             Call CloseSocket(UserIndex)
             Exit Sub
         End If

        .LoginId = Reader.ReadInt32()
        .Stats.ELV = Reader.ReadInt8()
        .Stats.Exp = Reader.ReadInt32()
        .genero = Reader.ReadInt8()
        .raza = Reader.ReadInt8()
        .clase = Reader.ReadInt8()
        .Hogar = Reader.ReadInt8()
        .Desc = Reader.ReadString8()
        .Stats.GLD = Reader.ReadInt32()
        .Stats.SkillPts = Reader.ReadInt16()
        .Pos.Map = Reader.ReadInt16()
        .Pos.X = Reader.ReadInt16()
        .Pos.Y = Reader.ReadInt16()
        If .Pos.Map > DebugMaps Or .Pos.Map < 1 Then
            .Pos.Map = CityUllathorpe.Map
            .Pos.X = CityUllathorpe.X
            .Pos.Y = CityUllathorpe.Y
        End If
        .ZonaId = ZonaByPos(.Pos)
        
        .Stats.Creditos = 0
        
            
'136         .MENSAJEINFORMACION = RS!message_info


'230         .flags.Pareja = RS!spouse
'232         .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)

        
        .OrigChar.Body = Reader.ReadInt16()
        .OrigChar.Head = Reader.ReadInt16()
        .Invent.BarcoObjIndex = Reader.ReadInt16()
        .OrigChar.Heading = Reader.ReadInt8()
        
        
        .Stats.MinHp = Reader.ReadInt16()
        .Stats.MaxHp = Reader.ReadInt16()
        .Stats.MinMAN = Reader.ReadInt16()
        .Stats.MaxMAN = Reader.ReadInt16()
        .Stats.MinSta = Reader.ReadInt16()
        .Stats.MaxSta = Reader.ReadInt16()
        .Stats.MinHam = Reader.ReadInt16()
        .Stats.MaxHam = Reader.ReadInt16()
        .Stats.MinAGU = Reader.ReadInt16()
        .Stats.MaxAGU = Reader.ReadInt16()
        .Stats.MinHIT = Reader.ReadInt16()
        .Stats.MaxHit = Reader.ReadInt16()
        .Stats.InventLevel = Reader.ReadInt8()
        
        .flags.Escondido = False
        If .Stats.MinHp <= 0 Then
            .flags.Muerto = 1
        End If
        
        .Stats.VaultLevel = Reader.ReadInt8()
        
        .GuildIndex = Reader.ReadInt16()
        .flags.ReturnPos.Map = Reader.ReadInt16()
        .flags.ReturnPos.X = Reader.ReadInt16()
        .flags.ReturnPos.Y = Reader.ReadInt16()

        .Faccion.Status = Reader.ReadInt8()
        .Faccion.RecompensasReal = Reader.ReadInt8()
        .Faccion.MatadosIngreso = Reader.ReadInt32()
        

        .Stats.NPCsMatados = Reader.ReadInt32()
        .Stats.UsuariosMatados = Reader.ReadInt32()
        .Stats.ciudadanosMatados = Reader.ReadInt32()
        .Stats.CriminalesMatados = Reader.ReadInt32()

        .Stats.MuertesPorNpcs = Reader.ReadInt32()
        .Stats.MuertesPorUsers = Reader.ReadInt32()
        .Stats.MuertesTotales = Reader.ReadInt32()
        
        .Stats.pasos = Reader.ReadInt32()
        
        .flags.Envenenado = Reader.ReadInt8()
        .flags.Incinerado = Reader.ReadInt8()
        .flags.Navegando = Reader.ReadInt8()
        .flags.Paralizado = Reader.ReadInt8()
        .flags.Silenciado = Reader.ReadInt8()
        .flags.Montado = Reader.ReadInt8()
        
        .ChatGlobal = Reader.ReadInt8()
        .ChatCombate = Reader.ReadInt8()
        
        .Stats.PuntosPesca = Reader.ReadInt32()
        .Stats.ELO = Reader.ReadInt32()
        
        
        .Stats.TiempoJugado = Reader.ReadInt32()
        .Counters.Pena = Reader.ReadInt16()
        
        .flags.MinutosRestantes = Reader.ReadInt16()
        '.flags.SegundosPasados = RS!silence_elapsed_seconds
    
        'INVENTORY
        
        
        .flags.Desnudo = 1
        
        tmpLength = Reader.ReadInt8()
        Dim obj As t_ObjData
        For i = 1 To tmpLength
            Slot = Reader.ReadInt8()
            ObjIndex = Reader.ReadInt16()
            .Invent.Object(Slot).ObjIndex = ObjIndex
            .Invent.Object(Slot).Amount = Reader.ReadInt16()
            .Invent.Object(Slot).Equipped = Reader.ReadInt8()
            
            If ObjIndex > 0 Then
                If LenB(ObjData(ObjIndex).Name) = 0 Then
                    .Invent.Object(Slot).ObjIndex = 0
                    .Invent.Object(Slot).Amount = 0
                    .Invent.Object(Slot).Equipped = 0
                End If
                If .Invent.Object(Slot).Equipped Then
                    obj = ObjData(ObjIndex)
                    Select Case obj.OBJType
                         Case e_OBJType.otWeapon
                            .OrigChar.WeaponAnim = NingunArma
                            .Invent.WeaponEqpObjIndex = ObjIndex
                            .Invent.WeaponEqpSlot = Slot
                         
                            If .flags.Navegando = 0 Then
                                .OrigChar.WeaponAnim = obj.WeaponAnim
                            End If
               
                         Case e_OBJType.otHerramientas
                            
                            .Invent.HerramientaEqpObjIndex = ObjIndex
                            .Invent.HerramientaEqpSlot = Slot
                      
                            If .flags.Montado = 0 And .flags.Navegando = 0 Then
                               .OrigChar.WeaponAnim = obj.WeaponAnim
                            End If
                
                         Case e_OBJType.otMagicos
                             
                             .Invent.MagicoObjIndex = ObjIndex
                             .Invent.MagicoSlot = Slot
                         
                             Select Case obj.EfectoMagico
                                 Case 1 ' Regenera Stamina
                                     .flags.RegeneracionSta = 1
                                 Case 2 'Modif la fuerza, agilidad, carisma, etc
                                     .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
                                     .Stats.UserAtributos(obj.QueAtributo) = MinimoInt(.Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento, .Stats.UserAtributosBackUP(obj.QueAtributo) * 2)
                         
                                     Call WriteFYA(UserIndex)
                                 Case 3 'Modifica los skills
                                     .Stats.UserSkills(obj.Que_Skill) = .Stats.UserSkills(obj.Que_Skill) + obj.CuantoAumento
                                 Case 4
                                     .flags.RegeneracionHP = 1
                                 Case 5
                                     .flags.RegeneracionMana = 1
                                 Case 6
                                     .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
                                     .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento
                                 Case 9
                                     .flags.NoMagiaEfecto = 1
                                 Case 10
                                     .flags.incinera = 1
                                 Case 11
                                     .flags.Paraliza = 1
                                 Case 12
                                     .flags.CarroMineria = 1
                         
                                 Case 14
                                     '.flags.DañoMagico = obj.CuantoAumento
                         
                                 Case 15 'Pendiete del Sacrificio
                                     .flags.PendienteDelSacrificio = 1
                                 Case 16
                                     .flags.NoPalabrasMagicas = 1
                                 Case 17
                                     .flags.NoDetectable = 1
                            
                                 Case 18 ' Pendiente del Experto
                                     .flags.PendienteDelExperto = 1
                                 Case 19
                                     .flags.Envenena = 1
                                 Case 20 'Anillo ocultismo
                                     .flags.AnilloOcultismo = 1
             
                             End Select
                         Case e_OBJType.otNudillos
                             .Invent.NudilloObjIndex = ObjIndex
                             .Invent.NudilloSlot = Slot
                             
                             If .flags.Montado = 0 And .flags.Navegando = 0 Then
                                .OrigChar.WeaponAnim = obj.WeaponAnim
                             End If
                         Case e_OBJType.otFlechas
                             .Invent.MunicionEqpObjIndex = ObjIndex
                             .Invent.MunicionEqpSlot = Slot
                         Case e_OBJType.otArmadura
                             
                             .Invent.ArmourEqpObjIndex = ObjIndex
                             .Invent.ArmourEqpSlot = Slot
                             If .flags.Montado = 0 And .flags.Navegando = 0 Then
                                 .OrigChar.Body = obj.Ropaje
                             End If
            
                         Case e_OBJType.otCasco
                             .Invent.CascoEqpObjIndex = ObjIndex
                             .Invent.CascoEqpSlot = Slot
                     
                             If .flags.Navegando = 0 Then
                                 .OrigChar.CascoAnim = obj.CascoAnim
                             End If
                         
                         Case e_OBJType.otEscudo
                             
                             .Invent.EscudoEqpObjIndex = ObjIndex
                             .Invent.EscudoEqpSlot = Slot
                             If .flags.Navegando = 0 And .flags.Montado = 0 Then
                                 .OrigChar.ShieldAnim = obj.ShieldAnim
                             End If
                         Case e_OBJType.otDanyoMagico, e_OBJType.otResistencia
                            
                             If ObjData(ObjIndex).OBJType = e_OBJType.otResistencia Then
                                 .Invent.ResistenciaEqpObjIndex = ObjIndex
                                 .Invent.ResistenciaEqpSlot = Slot
                             ElseIf ObjData(ObjIndex).OBJType = e_OBJType.otDanyoMagico Then
                                 .Invent.DañoMagicoEqpObjIndex = ObjIndex
                                 .Invent.DañoMagicoEqpSlot = Slot
                             End If
                    End Select
                End If
            End If
        Next i
            
        'VAULT
        tmpLength = Reader.ReadInt8()
        For i = 1 To tmpLength
        
            Slot = Reader.ReadInt8()

            With .BancoInvent.Object(Slot)
                .ObjIndex = Reader.ReadInt16()
                .Amount = Reader.ReadInt16()
                
                If .ObjIndex <> 0 Then
                    If LenB(ObjData(.ObjIndex).Name) = 0 Then
                        .ObjIndex = 0
                        .Amount = 0
                        .Equipped = 0
                    End If
                End If
            
            End With
        Next i
    
        'SKILLS
        tmpLength = Reader.ReadInt8()
        For i = 1 To tmpLength
            Slot = Reader.ReadInt8()
            .Stats.UserSkills(Slot) = Reader.ReadInt8()
            .Stats.UserSkillsAssigned(Slot) = Reader.ReadInt8()
        Next i
    
        'SPELLS
        tmpLength = Reader.ReadInt8()
        For i = 1 To tmpLength
            Slot = Reader.ReadInt8()
            .Stats.UserHechizos(Slot) = Reader.ReadInt8()
        Next i
    
        'QUESTS COMPLETED
        .QuestStats.NumQuestsDone = Reader.ReadInt16()
         If (.QuestStats.NumQuestsDone > 0) Then
            ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
            For i = 1 To .QuestStats.NumQuestsDone
                .QuestStats.QuestsDone(i) = Reader.ReadInt16()
            Next i
        End If
        tmpLength = Reader.ReadInt16()
        For i = 1 To tmpLength
            
            .QuestStats.Quests(i).QuestIndex = Reader.ReadInt16()
            data = Reader.ReadString8()

            
            If .QuestStats.Quests(i).QuestIndex > 0 Then
                If QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs Then
                    Dim NPCs() As String
                    
                    NPCs = Split(Split(data, "|")(0), "-")
                    ReDim .QuestStats.Quests(i).NPCsKilled(1 To QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs)
                    
                    For LoopC = 1 To QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs
                    .QuestStats.Quests(i).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
                    Next LoopC
                End If
                
                If QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs Then
                
                    Dim NPCsTarget() As String
                    
                    NPCsTarget = Split(Split(data, "|")(1), "-")
                    ReDim .QuestStats.Quests(i).NPCsTarget(1 To QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs)
                    
                    For LoopC = 1 To QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs
                    .QuestStats.Quests(i).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
                    Next LoopC
                
                End If
            
            End If
        Next i
        
        'ACHIVEMENTS
        tmpLength = Reader.ReadInt16()
        For i = 1 To tmpLength
            LoopC = Reader.ReadInt16()
        Next i
    
        'HOUSES
        tmpLength = Reader.ReadInt8()
        For i = 1 To tmpLength
            .keys(LoopC) = Reader.ReadInt16()
        Next i
        
    End With
    
    Call ConnectUser_Complete(UserIndex, UserList(UserIndex).Name, False)
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveLoadCharacter. ", Erl)


'            .LastGuildRejection = SanitizeNullValue(RS!guild_rejected_because, vbNullString)
             
'476         If RS Is Nothing Then Exit Sub
            
'            Dim tipo_usuario_db As Long
'            tipo_usuario_db = RS!is_active_patron
            
'            Select Case tipo_usuario_db
'                Case patron_tier_aventurero
'                    .Stats.tipoUsuario = e_TipoUsuario.tAventurero
'                Case patron_tier_heroe
'                    .Stats.tipoUsuario = e_TipoUsuario.tHeroe
'                Case patron_tier_leyenda
'                    .Stats.tipoUsuario = e_TipoUsuario.tLeyenda
'            End Select
            
    
    
End Sub

Private Sub ReceiveListGuilds(ByVal UserIndex As Integer, ByRef Reader As Network.Reader)
    On Error GoTo Handler:

    If UserIndex = 0 Then Exit Sub
    
    Dim i As Integer
    Dim tmpLength As Integer
    
    tmpLength = Reader.ReadInt16()
    
    'TODO: Falta continuar la implementación
'    For i = 1 To tmpLength
'        guildID = Reader.ReadInt32()
'        guildAt = Reader.ReadString8()
'        GuildName = Reader.ReadString8()
'        guildLevel = Reader.ReadInt()
'        guildFunder = Reader.ReadString8()
'        GuildLeader = Reader.ReadString8()
'        guildAntifaction = Reader.ReadInt16()
'        guildAlignation = Reader.ReadString8()
'        guildMembers = Reader.ReadInt16()
'        guildUrl = Reader.ReadString8()
'        guildDesc = Reader.ReadString8()
'        guildCodex = Reader.ReadString8()
'    Next
    
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Error en ReceiveListGuilds. ", Erl)
    
End Sub

