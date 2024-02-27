Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Mart�n Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Mart�n Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Mart�n Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private LastPacket      As Byte

Private IterationsHID   As Integer

Private Const MAX_ITERATIONS_HID = 200

Private Enum ServerPacketID

    Connected
    logged                  ' LOGGED  0
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquiteToggle
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU   10
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    CharSwing               ' U1
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF 20
    PartySafeOn
    PartySafeOff
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE 30
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+   40
    ShowMessageBox          ' !!
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    fxpiso
    ObjectDelete            ' BO  50
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01 60
    ChangeInventorySlot     ' CSI
    InventoryUnlockSlots
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU 70
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER 80
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR 90
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    ChangeUserTradeSlot     ' COMUSUINV
    UpdateTagAndStatus
    FYA
    CerrarleCliente
    Contadores
    ShowPapiro
    
    'GM messages
    SpawnListt               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    UserOnline '110
    ParticleFX
    ParticleFXToFloor
    ParticleFXWithDestino
    ParticleFXWithDestinoXY
    Hora
    Light
    AuraToChar
    SpeedToChar
    LightToFloor
    NieveToggle
    NieblaToggle
    Goliath
    TextOverChar
    TextOverTile
    TextCharDrop
    FlashScreen
    AlquimistaObj
    ShowAlquimiaForm
    SastreObj
    ShowSastreForm ' 126
    VelocidadToggle
    MacroTrabajoToggle
    BindKeys
    ShowfrmLogear
    ShowFrmMapa
    InmovilizadoOK
    BarFx
    LocaleMsg
    ShowPregunta
    DatosGrupo
    ubicacion
    ArmaMov
    EscudoMov
    ViajarForm
    NadarToggle
    ShowFundarClanForm
    CharUpdateHP
    CharUpdateMAN
    PosLLamadaDeClan
    QuestDetails
    QuestListSend
    NpcQuestListSend
    UpdateNPCSimbolo
    ClanSeguro
    Intervals
    UpdateUserKey
    UpdateRM
    UpdateDM
    SeguroResu
    Stopped
    InvasionInfo
    CommerceRecieveChatMessage
    DoAnimation
    OpenCrafting
    CraftingItem
    CraftingCatalyst
    CraftingResult
    ForceUpdate
    GuardNotice
    AnswerReset
    ObjQuestListSend
    UpdateBankGld
    PelearConPezEspecial
    Privilegios
    ShopInit
    UpdateShopClienteCredits
    UpdateFlag
    CharAtaca
    NotificarClienteSeguido
    RecievePosSeguimiento
    CancelarSeguimiento
    GetInventarioHechizos
    NotificarClienteCasteo
    SendFollowingCharindex
    ForceCharMoveSiguiendo
    PosUpdateUserChar
    PosUpdateChar
    PlayWaveStep
    ShopPjsInit
    AccountCharacterList
    ComboCooldown
    [PacketCount]
End Enum

Public Enum ClientPacketID
    
    CreateAccount
    LoginAccount
    '--------------------
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPA�AR
    PetLeave                '/LIBERAR
    GrupoMsg                '/GrupoMsg
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    punishments             '/PENAS
    Gamble                  '/APOSTAR
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    PartySafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    
    'GM messages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    Invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    banip                   '/BANIP
    UnBanIp                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    Participar              '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    chatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    
    'Nuevas Ladder
    SetSpeed                '/SPEED
    GlobalMessage           '/CONSOLA
    GlobalOnOff
    UseKey
    Day
    SetTime
    DonateGold              '/DONAR
    Promedio                '/PROMEDIO
    GiveItem                '/DAR
    OfertaInicial
    OfertaDeSubasta
    QuestionGM
    CuentaRegresiva
    PossUser
    Duel
    AcceptDuel
    CancelDuel
    QuitDuel
    NieveToggle
    NieblaToggle
    TransFerGold
    Moveitem
    Genio
    Casarse
    CraftAlquimista
    FlagTrabajar
    CraftSastre
    MensajeUser
    TraerBoveda
    CompletarAccion
    InvitarGrupo
    ResponderPregunta
    RequestGrupo
    AbandonarGrupo
    HecharDeGrupo
    MacroPossent
    SubastaInfo
    BanCuenta
    UnbanCuenta
    CerrarCliente
    EventoInfo
    CrearEvento
    BanTemporal
    CancelarExit
    CrearTorneo
    ComenzarTorneo
    CancelarTorneo
    BusquedaTesoro
    CompletarViaje
    BovedaMoveItem
    QuieroFundarClan
    llamadadeclan
    MarcaDeClanPack
    MarcaDeGMPack
    Quest
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    SeguroClan
    Home                    '/HOGAR
    Consulta                '/CONSULTA
    GetMapInfo              '/MAPINFO
    FinEvento
    SeguroResu
    CuentaExtractItem
    CuentaDeposit
    CreateEvent
    CommerceSendChatMessage
    LogMacroClickHechizo
    AddItemCrafting
    RemoveItemCrafting
    AddCatalyst
    RemoveCatalyst
    CraftItem
    CloseCrafting
    MoveCraftItem
    PetLeaveAll
    ResetChar              '/RESET NICK
    ResetearPersonaje
    DeleteItem
    FinalizarPescaEspecial
    RomperCania
    UseItemU
    RepeatMacro
    BuyShopItem
    PerdonFaccion              '/PERDONFACCION NAME
    IniciarCaptura           '/EVENTOCAPTURA PARTICIPANTES CANTIDAD_RONDAS NIVEL_MINIMO PRECIO
    ParticiparCaptura        '/PARTICIPARCAPTURA
    CancelarCaptura          '/CANCELARCAPTURA
    SeguirMouse
    SendPosSeguimiento
    NotifyInventarioHechizos
    PublicarPersonajeMAO
    
    DeleteCharacter
    [PacketCount]
End Enum

Public ServerPacketName() As String
Public ClientPacketName() As String

Private Reader As Network.Reader

Public Sub InitPacketNames()

ReDim ServerPacketName(ServerPacketID.PacketCount)
ReDim ClientPacketName(ClientPacketID.PacketCount)

ServerPacketName(ServerPacketID.Connected) = "Connected"
ServerPacketName(ServerPacketID.logged) = "logged"                  ' LOGGED  0
ServerPacketName(ServerPacketID.RemoveDialogs) = "RemoveDialogs"           ' QTDL
ServerPacketName(ServerPacketID.RemoveCharDialog) = "RemoveCharDialog"        ' QDL
ServerPacketName(ServerPacketID.NavigateToggle) = "NavigateToggle"          ' NAVEG
ServerPacketName(ServerPacketID.EquiteToggle) = "EquiteToggle"
ServerPacketName(ServerPacketID.Disconnect) = "Disconnect"              ' FINOK
ServerPacketName(ServerPacketID.CommerceEnd) = "CommerceEnd"             ' FINCOMOK
ServerPacketName(ServerPacketID.BankEnd) = "BankEnd"                 ' FINBANOK
ServerPacketName(ServerPacketID.CommerceInit) = "CommerceInit"            ' INITCOM
ServerPacketName(ServerPacketID.BankInit) = "BankInit"                ' INITBANCO
ServerPacketName(ServerPacketID.UserCommerceInit) = "UserCommerceInit"        ' INITCOMUSU   10
ServerPacketName(ServerPacketID.UserCommerceEnd) = "UserCommerceEnd"         ' FINCOMUSUOK
ServerPacketName(ServerPacketID.ShowBlacksmithForm) = "ShowBlacksmithForm"      ' SFH
ServerPacketName(ServerPacketID.ShowCarpenterForm) = "ShowCarpenterForm"       ' SFC
ServerPacketName(ServerPacketID.NPCKillUser) = "NPCKillUser"             ' 6
ServerPacketName(ServerPacketID.BlockedWithShieldUser) = "BlockedWithShieldUser"   ' 7
ServerPacketName(ServerPacketID.BlockedWithShieldOther) = "BlockedWithShieldOther"  ' 8
ServerPacketName(ServerPacketID.CharSwing) = "CharSwing"               ' U1
ServerPacketName(ServerPacketID.SafeModeOn) = "SafeModeOn"              ' SEGON
ServerPacketName(ServerPacketID.SafeModeOff) = "SafeModeOff"             ' SEGOFF 20
ServerPacketName(ServerPacketID.PartySafeOn) = "PartySafeOn"
ServerPacketName(ServerPacketID.PartySafeOff) = "PartySafeOff"
ServerPacketName(ServerPacketID.CantUseWhileMeditating) = "CantUseWhileMeditating"  ' M!
ServerPacketName(ServerPacketID.UpdateSta) = "UpdateSta"               ' ASS
ServerPacketName(ServerPacketID.UpdateMana) = "UpdateMana"              ' ASM
ServerPacketName(ServerPacketID.UpdateHP) = "UpdateHP"                ' ASH
ServerPacketName(ServerPacketID.UpdateGold) = "UpdateGold"              ' ASG
ServerPacketName(ServerPacketID.UpdateExp) = "UpdateExp"               ' ASE 30
ServerPacketName(ServerPacketID.ChangeMap) = "ChangeMap"               ' CM
ServerPacketName(ServerPacketID.PosUpdate) = "PosUpdate"               ' PU
ServerPacketName(ServerPacketID.NPCHitUser) = "NPCHitUser"              ' N2
ServerPacketName(ServerPacketID.UserHitNPC) = "UserHitNPC"              ' U2
ServerPacketName(ServerPacketID.UserAttackedSwing) = "UserAttackedSwing"       ' U3
ServerPacketName(ServerPacketID.UserHittedByUser) = "UserHittedByUser"        ' N4
ServerPacketName(ServerPacketID.UserHittedUser) = "UserHittedUser"          ' N5
ServerPacketName(ServerPacketID.ChatOverHead) = "ChatOverHead"            ' ||
ServerPacketName(ServerPacketID.ConsoleMsg) = "ConsoleMsg"              ' || - Beware!! its the same as above, but it was properly splitted
ServerPacketName(ServerPacketID.GuildChat) = "GuildChat"               ' |+   40
ServerPacketName(ServerPacketID.ShowMessageBox) = "ShowMessageBox"          ' !!
ServerPacketName(ServerPacketID.CharacterCreate) = "CharacterCreate"         ' CC
ServerPacketName(ServerPacketID.CharacterRemove) = "CharacterRemove"         ' BP
ServerPacketName(ServerPacketID.CharacterMove) = "CharacterMove"           ' MP, +, * and _ '
ServerPacketName(ServerPacketID.UserIndexInServer) = "UserIndexInServer"       ' IU
ServerPacketName(ServerPacketID.UserCharIndexInServer) = "UserCharIndexInServer"   ' IP
ServerPacketName(ServerPacketID.ForceCharMove) = "ForceCharMove"
ServerPacketName(ServerPacketID.CharacterChange) = "CharacterChange"         ' CP
ServerPacketName(ServerPacketID.ObjectCreate) = "ObjectCreate"            ' HO
ServerPacketName(ServerPacketID.fxpiso) = "fxpiso"
ServerPacketName(ServerPacketID.ObjectDelete) = "ObjectDelete"            ' BO  50
ServerPacketName(ServerPacketID.BlockPosition) = "BlockPosition"           ' BQ
ServerPacketName(ServerPacketID.PlayMIDI) = "PlayMIDI"                ' TM
ServerPacketName(ServerPacketID.PlayWave) = "PlayWave"                ' TW
ServerPacketName(ServerPacketID.guildList) = "guildList"               ' GL
ServerPacketName(ServerPacketID.AreaChanged) = "AreaChanged"             ' CA
ServerPacketName(ServerPacketID.PauseToggle) = "PauseToggle"             ' BKW
ServerPacketName(ServerPacketID.RainToggle) = "RainToggle"              ' LLU
ServerPacketName(ServerPacketID.CreateFX) = "CreateFX"                ' CFX
ServerPacketName(ServerPacketID.UpdateUserStats) = "UpdateUserStats"         ' EST
ServerPacketName(ServerPacketID.WorkRequestTarget) = "WorkRequestTarget"       ' T01 60
ServerPacketName(ServerPacketID.ChangeInventorySlot) = "ChangeInventorySlot"     ' CSI
ServerPacketName(ServerPacketID.InventoryUnlockSlots) = "InventoryUnlockSlots"
ServerPacketName(ServerPacketID.ChangeBankSlot) = "ChangeBankSlot"          ' SBO
ServerPacketName(ServerPacketID.ChangeSpellSlot) = "ChangeSpellSlot"         ' SHS
ServerPacketName(ServerPacketID.Atributes) = "Atributes"               ' ATR
ServerPacketName(ServerPacketID.BlacksmithWeapons) = "BlacksmithWeapons"       ' LAH
ServerPacketName(ServerPacketID.BlacksmithArmors) = "BlacksmithArmors"        ' LAR
ServerPacketName(ServerPacketID.CarpenterObjects) = "CarpenterObjects"        ' OBR
ServerPacketName(ServerPacketID.RestOK) = "RestOK"                  ' DOK
ServerPacketName(ServerPacketID.ErrorMsg) = "ErrorMsg"                ' ERR
ServerPacketName(ServerPacketID.Blind) = "Blind"                   ' CEGU 70
ServerPacketName(ServerPacketID.Dumb) = "Dumb"                    ' DUMB
ServerPacketName(ServerPacketID.ShowSignal) = "ShowSignal"              ' MCAR
ServerPacketName(ServerPacketID.ChangeNPCInventorySlot) = "ChangeNPCInventorySlot"  ' NPCI
ServerPacketName(ServerPacketID.UpdateHungerAndThirst) = "UpdateHungerAndThirst"   ' EHYS
ServerPacketName(ServerPacketID.MiniStats) = "MiniStats"               ' MEST
ServerPacketName(ServerPacketID.LevelUp) = "LevelUp"                 ' SUNI
ServerPacketName(ServerPacketID.AddForumMsg) = "AddForumMsg"             ' FMSG
ServerPacketName(ServerPacketID.ShowForumForm) = "ShowForumForm"           ' MFOR
ServerPacketName(ServerPacketID.SetInvisible) = "SetInvisible"            ' NOVER 80
ServerPacketName(ServerPacketID.MeditateToggle) = "MeditateToggle"          ' MEDOK
ServerPacketName(ServerPacketID.BlindNoMore) = "BlindNoMore"             ' NSEGUE
ServerPacketName(ServerPacketID.DumbNoMore) = "DumbNoMore"              ' NESTUP
ServerPacketName(ServerPacketID.SendSkills) = "SendSkills"              ' SKILLS
ServerPacketName(ServerPacketID.TrainerCreatureList) = "TrainerCreatureList"     ' LSTCRI
ServerPacketName(ServerPacketID.guildNews) = "guildNews"               ' GUILDNE
ServerPacketName(ServerPacketID.OfferDetails) = "OfferDetails"            ' PEACEDE & ALLIEDE
ServerPacketName(ServerPacketID.AlianceProposalsList) = "AlianceProposalsList"    ' ALLIEPR
ServerPacketName(ServerPacketID.PeaceProposalsList) = "PeaceProposalsList"      ' PEACEPR 90
ServerPacketName(ServerPacketID.CharacterInfo) = "CharacterInfo"           ' CHRINFO
ServerPacketName(ServerPacketID.GuildLeaderInfo) = "GuildLeaderInfo"         ' LEADERI
ServerPacketName(ServerPacketID.GuildDetails) = "GuildDetails"            ' CLANDET
ServerPacketName(ServerPacketID.ShowGuildFundationForm) = "ShowGuildFundationForm"  ' SHOWFUN
ServerPacketName(ServerPacketID.ParalizeOK) = "ParalizeOK"              ' PARADOK
ServerPacketName(ServerPacketID.ShowUserRequest) = "ShowUserRequest"         ' PETICIO
ServerPacketName(ServerPacketID.ChangeUserTradeSlot) = "ChangeUserTradeSlot"     ' COMUSUINV
ServerPacketName(ServerPacketID.UpdateTagAndStatus) = "UpdateTagAndStatus"
ServerPacketName(ServerPacketID.FYA) = "FYA"
ServerPacketName(ServerPacketID.CerrarleCliente) = "CerrarleCliente"
ServerPacketName(ServerPacketID.Contadores) = "Contadores"
ServerPacketName(ServerPacketID.ShowPapiro) = "ShowPapiro"
ServerPacketName(ServerPacketID.SpawnListt) = "SpawnListt"               ' SPL
ServerPacketName(ServerPacketID.ShowSOSForm) = "ShowSOSForm"             ' MSOS
ServerPacketName(ServerPacketID.ShowMOTDEditionForm) = "ShowMOTDEditionForm"     ' ZMOTD
ServerPacketName(ServerPacketID.ShowGMPanelForm) = "ShowGMPanelForm"         ' ABPANEL
ServerPacketName(ServerPacketID.UserNameList) = "UserNameList"            ' LISTUSU
ServerPacketName(ServerPacketID.UserOnline) = "UserOnline" '110
ServerPacketName(ServerPacketID.ParticleFX) = "ParticleFX"
ServerPacketName(ServerPacketID.ParticleFXToFloor) = "ParticleFXToFloor"
ServerPacketName(ServerPacketID.ParticleFXWithDestino) = "ParticleFXWithDestino"
ServerPacketName(ServerPacketID.ParticleFXWithDestinoXY) = "ParticleFXWithDestinoXY"
ServerPacketName(ServerPacketID.Hora) = "Hora"
ServerPacketName(ServerPacketID.Light) = "Light"
ServerPacketName(ServerPacketID.AuraToChar) = "AuraToChar"
ServerPacketName(ServerPacketID.SpeedToChar) = "SpeedToChar"
ServerPacketName(ServerPacketID.LightToFloor) = "LightToFloor"
ServerPacketName(ServerPacketID.NieveToggle) = "NieveToggle"
ServerPacketName(ServerPacketID.NieblaToggle) = "NieblaToggle"
ServerPacketName(ServerPacketID.Goliath) = "Goliath"
ServerPacketName(ServerPacketID.TextOverChar) = "TextOverChar"
ServerPacketName(ServerPacketID.TextOverTile) = "TextOverTile"
ServerPacketName(ServerPacketID.TextCharDrop) = "TextCharDrop"
ServerPacketName(ServerPacketID.FlashScreen) = "FlashScreen"
ServerPacketName(ServerPacketID.AlquimistaObj) = "AlquimistaObj"
ServerPacketName(ServerPacketID.ShowAlquimiaForm) = "ShowAlquimiaForm"
ServerPacketName(ServerPacketID.SastreObj) = "SastreObj"
ServerPacketName(ServerPacketID.ShowSastreForm) = "ShowSastreForm" ' 126
ServerPacketName(ServerPacketID.VelocidadToggle) = "VelocidadToggle"
ServerPacketName(ServerPacketID.MacroTrabajoToggle) = "MacroTrabajoToggle"
ServerPacketName(ServerPacketID.BindKeys) = "BindKeys"
ServerPacketName(ServerPacketID.ShowfrmLogear) = "ShowfrmLogear"
ServerPacketName(ServerPacketID.ShowFrmMapa) = "ShowFrmMapa"
ServerPacketName(ServerPacketID.InmovilizadoOK) = "InmovilizadoOK"
ServerPacketName(ServerPacketID.BarFx) = "BarFx"
ServerPacketName(ServerPacketID.LocaleMsg) = "LocaleMsg"
ServerPacketName(ServerPacketID.ShowPregunta) = "ShowPregunta"
ServerPacketName(ServerPacketID.DatosGrupo) = "DatosGrupo"
ServerPacketName(ServerPacketID.ubicacion) = "ubicacion"
ServerPacketName(ServerPacketID.ArmaMov) = "ArmaMov"
ServerPacketName(ServerPacketID.EscudoMov) = "EscudoMov"
ServerPacketName(ServerPacketID.ViajarForm) = "ViajarForm"
ServerPacketName(ServerPacketID.NadarToggle) = "NadarToggle"
ServerPacketName(ServerPacketID.ShowFundarClanForm) = "ShowFundarClanForm"
ServerPacketName(ServerPacketID.CharUpdateHP) = "CharUpdateHP"
ServerPacketName(ServerPacketID.CharUpdateMAN) = "CharUpdateMAN"
ServerPacketName(ServerPacketID.PosLLamadaDeClan) = "PosLLamadaDeClan"
ServerPacketName(ServerPacketID.QuestDetails) = "QuestDetails"
ServerPacketName(ServerPacketID.QuestListSend) = "QuestListSend"
ServerPacketName(ServerPacketID.NpcQuestListSend) = "NpcQuestListSend"
ServerPacketName(ServerPacketID.UpdateNPCSimbolo) = "UpdateNPCSimbolo"
ServerPacketName(ServerPacketID.ClanSeguro) = "ClanSeguro"
ServerPacketName(ServerPacketID.Intervals) = "Intervals"
ServerPacketName(ServerPacketID.UpdateUserKey) = "UpdateUserKey"
ServerPacketName(ServerPacketID.UpdateRM) = "UpdateRM"
ServerPacketName(ServerPacketID.UpdateDM) = "UpdateDM"
ServerPacketName(ServerPacketID.SeguroResu) = "SeguroResu"
ServerPacketName(ServerPacketID.Stopped) = "Stopped"
ServerPacketName(ServerPacketID.InvasionInfo) = "InvasionInfo"
ServerPacketName(ServerPacketID.CommerceRecieveChatMessage) = "CommerceRecieveChatMessage"
ServerPacketName(ServerPacketID.DoAnimation) = "DoAnimation"
ServerPacketName(ServerPacketID.OpenCrafting) = "OpenCrafting"
ServerPacketName(ServerPacketID.CraftingItem) = "CraftingItem"
ServerPacketName(ServerPacketID.CraftingCatalyst) = "CraftingCatalyst"
ServerPacketName(ServerPacketID.CraftingResult) = "CraftingResult"
ServerPacketName(ServerPacketID.ForceUpdate) = "ForceUpdate"
ServerPacketName(ServerPacketID.GuardNotice) = "GuardNotice"
ServerPacketName(ServerPacketID.AnswerReset) = "AnswerReset"
ServerPacketName(ServerPacketID.ObjQuestListSend) = "ObjQuestListSend"
ServerPacketName(ServerPacketID.UpdateBankGld) = "UpdateBankGld"
ServerPacketName(ServerPacketID.PelearConPezEspecial) = "PelearConPezEspecial"
ServerPacketName(ServerPacketID.Privilegios) = "Privilegios"
ServerPacketName(ServerPacketID.ShopInit) = "ShopInit"
ServerPacketName(ServerPacketID.UpdateShopClienteCredits) = "UpdateShopClienteCredits"
ServerPacketName(ServerPacketID.UpdateFlag) = "UpdateFlag"
ServerPacketName(ServerPacketID.CharAtaca) = "CharAtaca"
ServerPacketName(ServerPacketID.NotificarClienteSeguido) = "NotificarClienteSeguido"
ServerPacketName(ServerPacketID.RecievePosSeguimiento) = "RecievePosSeguimiento"
ServerPacketName(ServerPacketID.CancelarSeguimiento) = "CancelarSeguimiento"
ServerPacketName(ServerPacketID.GetInventarioHechizos) = "GetInventarioHechizos"
ServerPacketName(ServerPacketID.NotificarClienteCasteo) = "NotificarClienteCasteo"
ServerPacketName(ServerPacketID.SendFollowingCharindex) = "SendFollowingCharindex"
ServerPacketName(ServerPacketID.ForceCharMoveSiguiendo) = "ForceCharMoveSiguiendo"
ServerPacketName(ServerPacketID.PosUpdateUserChar) = "PosUpdateUserChar"
ServerPacketName(ServerPacketID.PosUpdateChar) = "PosUpdateChar"
ServerPacketName(ServerPacketID.PlayWaveStep) = "PlayWaveStep"
ServerPacketName(ServerPacketID.ShopPjsInit) = "ShopPjsInit"
ServerPacketName(ServerPacketID.AccountCharacterList) = "AccountCharacterList"


ClientPacketName(ClientPacketID.CraftCarpenter) = "CraftCarpenter"          'CNC
ClientPacketName(ClientPacketID.WorkLeftClick) = "WorkLeftClick"           'WLC
ClientPacketName(ClientPacketID.CreateNewGuild) = "CreateNewGuild"          'CIG
ClientPacketName(ClientPacketID.SpellInfo) = "SpellInfo"               'INFS
ClientPacketName(ClientPacketID.EquipItem) = "EquipItem"               'EQUI
ClientPacketName(ClientPacketID.ChangeHeading) = "ChangeHeading"           'CHEA
ClientPacketName(ClientPacketID.ModifySkills) = "ModifySkills"            'SKSE
ClientPacketName(ClientPacketID.Train) = "Train"                   'ENTR
ClientPacketName(ClientPacketID.CommerceBuy) = "CommerceBuy"             'COMP
ClientPacketName(ClientPacketID.BankExtractItem) = "BankExtractItem"         'RETI
ClientPacketName(ClientPacketID.CommerceSell) = "CommerceSell"            'VEND
ClientPacketName(ClientPacketID.BankDeposit) = "BankDeposit"             'DEPO
ClientPacketName(ClientPacketID.ForumPost) = "ForumPost"               'DEMSG
ClientPacketName(ClientPacketID.MoveSpell) = "MoveSpell"               'DESPHE
ClientPacketName(ClientPacketID.ClanCodexUpdate) = "ClanCodexUpdate"         'DESCOD
ClientPacketName(ClientPacketID.UserCommerceOffer) = "UserCommerceOffer"       'OFRECER
ClientPacketName(ClientPacketID.GuildAcceptPeace) = "GuildAcceptPeace"        'ACEPPEAT
ClientPacketName(ClientPacketID.GuildRejectAlliance) = "GuildRejectAlliance"     'RECPALIA
ClientPacketName(ClientPacketID.GuildRejectPeace) = "GuildRejectPeace"        'RECPPEAT
ClientPacketName(ClientPacketID.GuildAcceptAlliance) = "GuildAcceptAlliance"     'ACEPALIA
ClientPacketName(ClientPacketID.GuildOfferPeace) = "GuildOfferPeace"         'PEACEOFF
ClientPacketName(ClientPacketID.GuildOfferAlliance) = "GuildOfferAlliance"      'ALLIEOFF
ClientPacketName(ClientPacketID.GuildAllianceDetails) = "GuildAllianceDetails"    'ALLIEDET
ClientPacketName(ClientPacketID.GuildPeaceDetails) = "GuildPeaceDetails"       'PEACEDET
ClientPacketName(ClientPacketID.GuildRequestJoinerInfo) = "GuildRequestJoinerInfo"  'ENVCOMEN
ClientPacketName(ClientPacketID.GuildAlliancePropList) = "GuildAlliancePropList"   'ENVALPRO
ClientPacketName(ClientPacketID.GuildPeacePropList) = "GuildPeacePropList"      'ENVPROPP
ClientPacketName(ClientPacketID.GuildDeclareWar) = "GuildDeclareWar"         'DECGUERR
ClientPacketName(ClientPacketID.GuildNewWebsite) = "GuildNewWebsite"         'NEWWEBSI
ClientPacketName(ClientPacketID.GuildAcceptNewMember) = "GuildAcceptNewMember"    'ACEPTARI
ClientPacketName(ClientPacketID.GuildRejectNewMember) = "GuildRejectNewMember"    'RECHAZAR
ClientPacketName(ClientPacketID.GuildKickMember) = "GuildKickMember"         'ECHARCLA
ClientPacketName(ClientPacketID.GuildUpdateNews) = "GuildUpdateNews"         'ACTGNEWS
ClientPacketName(ClientPacketID.GuildMemberInfo) = "GuildMemberInfo"         '1HRINFO<
ClientPacketName(ClientPacketID.GuildOpenElections) = "GuildOpenElections"      'ABREELEC
ClientPacketName(ClientPacketID.GuildRequestMembership) = "GuildRequestMembership"  'SOLICITUD
ClientPacketName(ClientPacketID.GuildRequestDetails) = "GuildRequestDetails"     'CLANDETAILS
ClientPacketName(ClientPacketID.Online) = "Online"                  '/ONLINE
ClientPacketName(ClientPacketID.Quit) = "Quit"                    '/SALIR
ClientPacketName(ClientPacketID.GuildLeave) = "GuildLeave"              '/SALIRCLAN
ClientPacketName(ClientPacketID.RequestAccountState) = "RequestAccountState"     '/BALANCE
ClientPacketName(ClientPacketID.PetStand) = "PetStand"                '/QUIETO
ClientPacketName(ClientPacketID.PetFollow) = "PetFollow"               '/ACOMPA�AR
ClientPacketName(ClientPacketID.PetLeave) = "PetLeave"                '/LIBERAR
ClientPacketName(ClientPacketID.GrupoMsg) = "GrupoMsg"                '/GrupoMsg
ClientPacketName(ClientPacketID.TrainList) = "TrainList"               '/ENTRENAR
ClientPacketName(ClientPacketID.Rest) = "Rest"                    '/DESCANSAR
ClientPacketName(ClientPacketID.Meditate) = "Meditate"                '/MEDITAR
ClientPacketName(ClientPacketID.Resucitate) = "Resucitate"              '/RESUCITAR
ClientPacketName(ClientPacketID.Heal) = "Heal"                    '/CURAR
ClientPacketName(ClientPacketID.Help) = "Help"                    '/AYUDA
ClientPacketName(ClientPacketID.RequestStats) = "RequestStats"            '/EST
ClientPacketName(ClientPacketID.CommerceStart) = "CommerceStart"           '/COMERCIAR
ClientPacketName(ClientPacketID.BankStart) = "BankStart"               '/BOVEDA
ClientPacketName(ClientPacketID.Enlist) = "Enlist"                  '/ENLISTAR
ClientPacketName(ClientPacketID.Information) = "Information"             '/INFORMACION
ClientPacketName(ClientPacketID.Reward) = "Reward"                  '/RECOMPENSA
ClientPacketName(ClientPacketID.RequestMOTD) = "RequestMOTD"             '/MOTD
ClientPacketName(ClientPacketID.UpTime) = "UpTime"                  '/UPTIME
ClientPacketName(ClientPacketID.GuildMessage) = "GuildMessage"            '/CMSG
ClientPacketName(ClientPacketID.CentinelReport) = "CentinelReport"          '/CENTINELA
ClientPacketName(ClientPacketID.GuildOnline) = "GuildOnline"             '/ONLINECLAN
ClientPacketName(ClientPacketID.CouncilMessage) = "CouncilMessage"          '/BMSG
ClientPacketName(ClientPacketID.RoleMasterRequest) = "RoleMasterRequest"       '/ROL
ClientPacketName(ClientPacketID.ChangeDescription) = "ChangeDescription"       '/DESC
ClientPacketName(ClientPacketID.GuildVote) = "GuildVote"               '/VOTO
ClientPacketName(ClientPacketID.punishments) = "punishments"             '/PENAS
ClientPacketName(ClientPacketID.Gamble) = "Gamble"                  '/APOSTAR
ClientPacketName(ClientPacketID.LeaveFaction) = "LeaveFaction"            '/RETIRAR ( with no arguments )
ClientPacketName(ClientPacketID.BankExtractGold) = "BankExtractGold"         '/RETIRAR ( with arguments )
ClientPacketName(ClientPacketID.BankDepositGold) = "BankDepositGold"         '/DEPOSITAR
ClientPacketName(ClientPacketID.Denounce) = "Denounce"                '/DENUNCIAR
ClientPacketName(ClientPacketID.LoginExistingChar) = "LoginExistingChar"       'OLOGIN
ClientPacketName(ClientPacketID.LoginNewChar) = "LoginNewChar"            'NLOGIN
ClientPacketName(ClientPacketID.Talk) = "Talk"                    ';
ClientPacketName(ClientPacketID.Yell) = "Yell"                    '-
ClientPacketName(ClientPacketID.Whisper) = "Whisper"                 '\
ClientPacketName(ClientPacketID.Walk) = "Walk"                    'M
ClientPacketName(ClientPacketID.RequestPositionUpdate) = "RequestPositionUpdate"   'RPU
ClientPacketName(ClientPacketID.Attack) = "Attack"                  'AT
ClientPacketName(ClientPacketID.PickUp) = "PickUp"                  'AG
ClientPacketName(ClientPacketID.SafeToggle) = "SafeToggle"              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
ClientPacketName(ClientPacketID.PartySafeToggle) = "PartySafeToggle"
ClientPacketName(ClientPacketID.RequestGuildLeaderInfo) = "RequestGuildLeaderInfo"  'GLINFO
ClientPacketName(ClientPacketID.RequestAtributes) = "RequestAtributes"        'ATR
ClientPacketName(ClientPacketID.RequestSkills) = "RequestSkills"           'ESKI
ClientPacketName(ClientPacketID.RequestMiniStats) = "RequestMiniStats"        'FEST
ClientPacketName(ClientPacketID.CommerceEnd) = "CommerceEnd"             'FINCOM
ClientPacketName(ClientPacketID.UserCommerceEnd) = "UserCommerceEnd"         'FINCOMUSU
ClientPacketName(ClientPacketID.BankEnd) = "BankEnd"                 'FINBAN
ClientPacketName(ClientPacketID.UserCommerceOk) = "UserCommerceOk"          'COMUSUOK
ClientPacketName(ClientPacketID.UserCommerceReject) = "UserCommerceReject"      'COMUSUNO
ClientPacketName(ClientPacketID.Drop) = "Drop"                    'TI
ClientPacketName(ClientPacketID.CastSpell) = "CastSpell"               'LH
ClientPacketName(ClientPacketID.LeftClick) = "LeftClick"               'LC
ClientPacketName(ClientPacketID.DoubleClick) = "DoubleClick"             'RC
ClientPacketName(ClientPacketID.Work) = "Work"                    'UK
ClientPacketName(ClientPacketID.UseSpellMacro) = "UseSpellMacro"           'UMH
ClientPacketName(ClientPacketID.UseItem) = "UseItem"                 'USA
ClientPacketName(ClientPacketID.CraftBlacksmith) = "CraftBlacksmith"         'CNS
ClientPacketName(ClientPacketID.GMMessage) = "GMMessage"               '/GMSG
ClientPacketName(ClientPacketID.showName) = "showName"                '/SHOWNAME
ClientPacketName(ClientPacketID.OnlineRoyalArmy) = "OnlineRoyalArmy"         '/ONLINEREAL
ClientPacketName(ClientPacketID.OnlineChaosLegion) = "OnlineChaosLegion"       '/ONLINECAOS
ClientPacketName(ClientPacketID.GoNearby) = "GoNearby"                '/IRCERCA
ClientPacketName(ClientPacketID.comment) = "comment"                 '/REM
ClientPacketName(ClientPacketID.serverTime) = "serverTime"              '/HORA
ClientPacketName(ClientPacketID.Where) = "Where"                   '/DONDE
ClientPacketName(ClientPacketID.CreaturesInMap) = "CreaturesInMap"          '/NENE
ClientPacketName(ClientPacketID.WarpMeToTarget) = "WarpMeToTarget"          '/TELEPLOC
ClientPacketName(ClientPacketID.WarpChar) = "WarpChar"                '/TELEP
ClientPacketName(ClientPacketID.Silence) = "Silence"                 '/SILENCIAR
ClientPacketName(ClientPacketID.SOSShowList) = "SOSShowList"             '/SHOW SOS
ClientPacketName(ClientPacketID.SOSRemove) = "SOSRemove"               'SOSDONE
ClientPacketName(ClientPacketID.GoToChar) = "GoToChar"                '/IRA
ClientPacketName(ClientPacketID.Invisible) = "Invisible"               '/INVISIBLE
ClientPacketName(ClientPacketID.GMPanel) = "GMPanel"                 '/PANELGM
ClientPacketName(ClientPacketID.RequestUserList) = "RequestUserList"         'LISTUSU
ClientPacketName(ClientPacketID.Working) = "Working"                 '/TRABAJANDO
ClientPacketName(ClientPacketID.Hiding) = "Hiding"                  '/OCULTANDO
ClientPacketName(ClientPacketID.Jail) = "Jail"                    '/CARCEL
ClientPacketName(ClientPacketID.KillNPC) = "KillNPC"                 '/RMATA
ClientPacketName(ClientPacketID.WarnUser) = "WarnUser"                '/ADVERTENCIA
ClientPacketName(ClientPacketID.EditChar) = "EditChar"                '/MOD
ClientPacketName(ClientPacketID.RequestCharInfo) = "RequestCharInfo"         '/INFO
ClientPacketName(ClientPacketID.RequestCharStats) = "RequestCharStats"        '/STAT
ClientPacketName(ClientPacketID.RequestCharGold) = "RequestCharGold"         '/BAL
ClientPacketName(ClientPacketID.RequestCharInventory) = "RequestCharInventory"    '/INV
ClientPacketName(ClientPacketID.RequestCharBank) = "RequestCharBank"         '/BOV
ClientPacketName(ClientPacketID.RequestCharSkills) = "RequestCharSkills"       '/SKILLS
ClientPacketName(ClientPacketID.ReviveChar) = "ReviveChar"              '/REVIVIR
ClientPacketName(ClientPacketID.OnlineGM) = "OnlineGM"                '/ONLINEGM
ClientPacketName(ClientPacketID.OnlineMap) = "OnlineMap"               '/ONLINEMAP
ClientPacketName(ClientPacketID.Forgive) = "Forgive"                 '/PERDON
ClientPacketName(ClientPacketID.Kick) = "Kick"                    '/ECHAR
ClientPacketName(ClientPacketID.Execute) = "Execute"                 '/EJECUTAR
ClientPacketName(ClientPacketID.BanChar) = "BanChar"                 '/BAN
ClientPacketName(ClientPacketID.UnbanChar) = "UnbanChar"               '/UNBAN
ClientPacketName(ClientPacketID.NPCFollow) = "NPCFollow"               '/SEGUIR
ClientPacketName(ClientPacketID.SummonChar) = "SummonChar"              '/SUM
ClientPacketName(ClientPacketID.SpawnListRequest) = "SpawnListRequest"        '/CC
ClientPacketName(ClientPacketID.SpawnCreature) = "SpawnCreature"           'SPA
ClientPacketName(ClientPacketID.ResetNPCInventory) = "ResetNPCInventory"       '/RESETINV
ClientPacketName(ClientPacketID.CleanWorld) = "CleanWorld"              '/LIMPIAR
ClientPacketName(ClientPacketID.ServerMessage) = "ServerMessage"           '/RMSG
ClientPacketName(ClientPacketID.NickToIP) = "NickToIP"                '/NICK2IP
ClientPacketName(ClientPacketID.IPToNick) = "IPToNick"                '/IP2NICK
ClientPacketName(ClientPacketID.GuildOnlineMembers) = "GuildOnlineMembers"      '/ONCLAN
ClientPacketName(ClientPacketID.TeleportCreate) = "TeleportCreate"          '/CT
ClientPacketName(ClientPacketID.TeleportDestroy) = "TeleportDestroy"         '/DT
ClientPacketName(ClientPacketID.RainToggle) = "RainToggle"              '/LLUVIA
ClientPacketName(ClientPacketID.SetCharDescription) = "SetCharDescription"      '/SETDESC
ClientPacketName(ClientPacketID.ForceWAVEToMap) = "ForceWAVEToMap"          '/FORCEWAVMAP
ClientPacketName(ClientPacketID.RoyalArmyMessage) = "RoyalArmyMessage"        '/REALMSG
ClientPacketName(ClientPacketID.ChaosLegionMessage) = "ChaosLegionMessage"      '/CAOSMSG
ClientPacketName(ClientPacketID.TalkAsNPC) = "TalkAsNPC"               '/TALKAS
ClientPacketName(ClientPacketID.DestroyAllItemsInArea) = "DestroyAllItemsInArea"   '/MASSDEST
ClientPacketName(ClientPacketID.AcceptRoyalCouncilMember) = "AcceptRoyalCouncilMember" '/ACEPTCONSE
ClientPacketName(ClientPacketID.AcceptChaosCouncilMember) = "AcceptChaosCouncilMember" '/ACEPTCONSECAOS
ClientPacketName(ClientPacketID.ItemsInTheFloor) = "ItemsInTheFloor"         '/PISO
ClientPacketName(ClientPacketID.MakeDumb) = "MakeDumb"                '/ESTUPIDO
ClientPacketName(ClientPacketID.MakeDumbNoMore) = "MakeDumbNoMore"          '/NOESTUPIDO
ClientPacketName(ClientPacketID.CouncilKick) = "CouncilKick"             '/KICKCONSE
ClientPacketName(ClientPacketID.SetTrigger) = "SetTrigger"              '/TRIGGER
ClientPacketName(ClientPacketID.AskTrigger) = "AskTrigger"              '/TRIGGER with no args
ClientPacketName(ClientPacketID.BannedIPList) = "BannedIPList"            '/BANIPLIST
ClientPacketName(ClientPacketID.BannedIPReload) = "BannedIPReload"          '/BANIPRELOAD
ClientPacketName(ClientPacketID.GuildMemberList) = "GuildMemberList"         '/MIEMBROSCLAN
ClientPacketName(ClientPacketID.GuildBan) = "GuildBan"                '/BANCLAN
ClientPacketName(ClientPacketID.banip) = "banip"                   '/BANIP
ClientPacketName(ClientPacketID.UnBanIp) = "UnBanIp"                 '/UNBANIP
ClientPacketName(ClientPacketID.CreateItem) = "CreateItem"              '/CI
ClientPacketName(ClientPacketID.DestroyItems) = "DestroyItems"            '/DEST
ClientPacketName(ClientPacketID.ChaosLegionKick) = "ChaosLegionKick"         '/NOCAOS
ClientPacketName(ClientPacketID.RoyalArmyKick) = "RoyalArmyKick"           '/NOREAL
ClientPacketName(ClientPacketID.ForceMIDIAll) = "ForceMIDIAll"            '/FORCEMIDI
ClientPacketName(ClientPacketID.ForceWAVEAll) = "ForceWAVEAll"            '/FORCEWAV
ClientPacketName(ClientPacketID.RemovePunishment) = "RemovePunishment"        '/BORRARPENA
ClientPacketName(ClientPacketID.TileBlockedToggle) = "TileBlockedToggle"       '/BLOQ
ClientPacketName(ClientPacketID.KillNPCNoRespawn) = "KillNPCNoRespawn"        '/MATA
ClientPacketName(ClientPacketID.KillAllNearbyNPCs) = "KillAllNearbyNPCs"       '/MASSKILL
ClientPacketName(ClientPacketID.LastIP) = "LastIP"                  '/LASTIP
ClientPacketName(ClientPacketID.ChangeMOTD) = "ChangeMOTD"              '/MOTDCAMBIA
ClientPacketName(ClientPacketID.SetMOTD) = "SetMOTD"                 'ZMOTD
ClientPacketName(ClientPacketID.SystemMessage) = "SystemMessage"           '/SMSG
ClientPacketName(ClientPacketID.CreateNPC) = "CreateNPC"               '/ACC
ClientPacketName(ClientPacketID.CreateNPCWithRespawn) = "CreateNPCWithRespawn"    '/RACC
ClientPacketName(ClientPacketID.ImperialArmour) = "ImperialArmour"          '/AI1 - 4
ClientPacketName(ClientPacketID.ChaosArmour) = "ChaosArmour"             '/AC1 - 4
ClientPacketName(ClientPacketID.NavigateToggle) = "NavigateToggle"          '/NAVE
ClientPacketName(ClientPacketID.ServerOpenToUsersToggle) = "ServerOpenToUsersToggle" '/HABILITAR
ClientPacketName(ClientPacketID.Participar) = "Participar"              '/APAGAR
ClientPacketName(ClientPacketID.TurnCriminal) = "TurnCriminal"            '/CONDEN
ClientPacketName(ClientPacketID.ResetFactions) = "ResetFactions"           '/RAJAR
ClientPacketName(ClientPacketID.RemoveCharFromGuild) = "RemoveCharFromGuild"     '/RAJARCLAN
ClientPacketName(ClientPacketID.AlterName) = "AlterName"               '/ANAME
ClientPacketName(ClientPacketID.DoBackUp) = "DoBackUp"                '/DOBACKUP
ClientPacketName(ClientPacketID.ShowGuildMessages) = "ShowGuildMessages"       '/SHOWCMSG
ClientPacketName(ClientPacketID.ChangeMapInfoPK) = "ChangeMapInfoPK"         '/MODMAPINFO PK
ClientPacketName(ClientPacketID.ChangeMapInfoBackup) = "ChangeMapInfoBackup"     '/MODMAPINFO BACKUP
ClientPacketName(ClientPacketID.ChangeMapInfoRestricted) = "ChangeMapInfoRestricted" '/MODMAPINFO RESTRINGIR
ClientPacketName(ClientPacketID.ChangeMapInfoNoMagic) = "ChangeMapInfoNoMagic"    '/MODMAPINFO MAGIASINEFECTO
ClientPacketName(ClientPacketID.ChangeMapInfoNoInvi) = "ChangeMapInfoNoInvi"     '/MODMAPINFO INVISINEFECTO
ClientPacketName(ClientPacketID.ChangeMapInfoNoResu) = "ChangeMapInfoNoResu"     '/MODMAPINFO RESUSINEFECTO
ClientPacketName(ClientPacketID.ChangeMapInfoLand) = "ChangeMapInfoLand"       '/MODMAPINFO TERRENO
ClientPacketName(ClientPacketID.ChangeMapInfoZone) = "ChangeMapInfoZone"       '/MODMAPINFO ZONA
ClientPacketName(ClientPacketID.SaveChars) = "SaveChars"               '/GRABAR
ClientPacketName(ClientPacketID.CleanSOS) = "CleanSOS"                '/BORRAR SOS
ClientPacketName(ClientPacketID.ShowServerForm) = "ShowServerForm"          '/SHOW INT
ClientPacketName(ClientPacketID.night) = "night"                   '/NOCHE
ClientPacketName(ClientPacketID.KickAllChars) = "KickAllChars"            '/ECHARTODOSPJS
ClientPacketName(ClientPacketID.ReloadNPCs) = "ReloadNPCs"              '/RELOADNPCS
ClientPacketName(ClientPacketID.ReloadServerIni) = "ReloadServerIni"         '/RELOADSINI
ClientPacketName(ClientPacketID.ReloadSpells) = "ReloadSpells"            '/RELOADHECHIZOS
ClientPacketName(ClientPacketID.ReloadObjects) = "ReloadObjects"           '/RELOADOBJ
ClientPacketName(ClientPacketID.chatColor) = "chatColor"               '/CHATCOLOR
ClientPacketName(ClientPacketID.Ignored) = "Ignored"                 '/IGNORADO
ClientPacketName(ClientPacketID.CheckSlot) = "CheckSlot"               '/SLOT
ClientPacketName(ClientPacketID.SetSpeed) = "SetSpeed"                '/SPEED
ClientPacketName(ClientPacketID.GlobalMessage) = "GlobalMessage"           '/CONSOLA
ClientPacketName(ClientPacketID.GlobalOnOff) = "GlobalOnOff"
ClientPacketName(ClientPacketID.UseKey) = "UseKey"
ClientPacketName(ClientPacketID.Day) = "Day"
ClientPacketName(ClientPacketID.SetTime) = "SetTime"
ClientPacketName(ClientPacketID.DonateGold) = "DonateGold"              '/DONAR
ClientPacketName(ClientPacketID.Promedio) = "Promedio"                '/PROMEDIO
ClientPacketName(ClientPacketID.GiveItem) = "GiveItem"                '/DAR
ClientPacketName(ClientPacketID.OfertaInicial) = "OfertaInicial"
ClientPacketName(ClientPacketID.OfertaDeSubasta) = "OfertaDeSubasta"
ClientPacketName(ClientPacketID.QuestionGM) = "QuestionGM"
ClientPacketName(ClientPacketID.CuentaRegresiva) = "CuentaRegresiva"
ClientPacketName(ClientPacketID.PossUser) = "PossUser"
ClientPacketName(ClientPacketID.Duel) = "Duel"
ClientPacketName(ClientPacketID.AcceptDuel) = "AcceptDuel"
ClientPacketName(ClientPacketID.CancelDuel) = "CancelDuel"
ClientPacketName(ClientPacketID.QuitDuel) = "QuitDuel"
ClientPacketName(ClientPacketID.NieveToggle) = "NieveToggle"
ClientPacketName(ClientPacketID.NieblaToggle) = "NieblaToggle"
ClientPacketName(ClientPacketID.TransFerGold) = "TransFerGold"
ClientPacketName(ClientPacketID.Moveitem) = "Moveitem"
ClientPacketName(ClientPacketID.Genio) = "Genio"
ClientPacketName(ClientPacketID.Casarse) = "Casarse"
ClientPacketName(ClientPacketID.CraftAlquimista) = "CraftAlquimista"
ClientPacketName(ClientPacketID.FlagTrabajar) = "FlagTrabajar"
ClientPacketName(ClientPacketID.CraftSastre) = "CraftSastre"
ClientPacketName(ClientPacketID.MensajeUser) = "MensajeUser"
ClientPacketName(ClientPacketID.TraerBoveda) = "TraerBoveda"
ClientPacketName(ClientPacketID.CompletarAccion) = "CompletarAccion"
ClientPacketName(ClientPacketID.InvitarGrupo) = "InvitarGrupo"
ClientPacketName(ClientPacketID.ResponderPregunta) = "ResponderPregunta"
ClientPacketName(ClientPacketID.RequestGrupo) = "RequestGrupo"
ClientPacketName(ClientPacketID.AbandonarGrupo) = "AbandonarGrupo"
ClientPacketName(ClientPacketID.HecharDeGrupo) = "HecharDeGrupo"
ClientPacketName(ClientPacketID.MacroPossent) = "MacroPossent"
ClientPacketName(ClientPacketID.SubastaInfo) = "SubastaInfo"
ClientPacketName(ClientPacketID.BanCuenta) = "BanCuenta"
ClientPacketName(ClientPacketID.UnbanCuenta) = "UnbanCuenta"
ClientPacketName(ClientPacketID.CerrarCliente) = "CerrarCliente"
ClientPacketName(ClientPacketID.EventoInfo) = "EventoInfo"
ClientPacketName(ClientPacketID.CrearEvento) = "CrearEvento"
ClientPacketName(ClientPacketID.BanTemporal) = "BanTemporal"
ClientPacketName(ClientPacketID.CancelarExit) = "CancelarExit"
ClientPacketName(ClientPacketID.CrearTorneo) = "CrearTorneo"
ClientPacketName(ClientPacketID.ComenzarTorneo) = "ComenzarTorneo"
ClientPacketName(ClientPacketID.CancelarTorneo) = "CancelarTorneo"
ClientPacketName(ClientPacketID.BusquedaTesoro) = "BusquedaTesoro"
ClientPacketName(ClientPacketID.CompletarViaje) = "CompletarViaje"
ClientPacketName(ClientPacketID.BovedaMoveItem) = "BovedaMoveItem"
ClientPacketName(ClientPacketID.QuieroFundarClan) = "QuieroFundarClan"
ClientPacketName(ClientPacketID.llamadadeclan) = "llamadadeclan"
ClientPacketName(ClientPacketID.MarcaDeClanPack) = "MarcaDeClanPack"
ClientPacketName(ClientPacketID.MarcaDeGMPack) = "MarcaDeGMPack"
ClientPacketName(ClientPacketID.Quest) = "Quest"
ClientPacketName(ClientPacketID.QuestAccept) = "QuestAccept"
ClientPacketName(ClientPacketID.QuestListRequest) = "QuestListRequest"
ClientPacketName(ClientPacketID.QuestDetailsRequest) = "QuestDetailsRequest"
ClientPacketName(ClientPacketID.QuestAbandon) = "QuestAbandon"
ClientPacketName(ClientPacketID.SeguroClan) = "SeguroClan"
ClientPacketName(ClientPacketID.Home) = "Home"                    '/HOGAR
ClientPacketName(ClientPacketID.Consulta) = "Consulta"                '/CONSULTA
ClientPacketName(ClientPacketID.GetMapInfo) = "GetMapInfo"              '/MAPINFO
ClientPacketName(ClientPacketID.FinEvento) = "FinEvento"
ClientPacketName(ClientPacketID.SeguroResu) = "SeguroResu"
ClientPacketName(ClientPacketID.CuentaExtractItem) = "CuentaExtractItem"
ClientPacketName(ClientPacketID.CuentaDeposit) = "CuentaDeposit"
ClientPacketName(ClientPacketID.CreateEvent) = "CreateEvent"
ClientPacketName(ClientPacketID.CommerceSendChatMessage) = "CommerceSendChatMessage"
ClientPacketName(ClientPacketID.LogMacroClickHechizo) = "LogMacroClickHechizo"
ClientPacketName(ClientPacketID.AddItemCrafting) = "AddItemCrafting"
ClientPacketName(ClientPacketID.RemoveItemCrafting) = "RemoveItemCrafting"
ClientPacketName(ClientPacketID.AddCatalyst) = "AddCatalyst"
ClientPacketName(ClientPacketID.RemoveCatalyst) = "RemoveCatalyst"
ClientPacketName(ClientPacketID.CraftItem) = "CraftItem"
ClientPacketName(ClientPacketID.CloseCrafting) = "CloseCrafting"
ClientPacketName(ClientPacketID.MoveCraftItem) = "MoveCraftItem"
ClientPacketName(ClientPacketID.PetLeaveAll) = "PetLeaveAll"
ClientPacketName(ClientPacketID.ResetChar) = "ResetChar"              '/RESET NICK
ClientPacketName(ClientPacketID.ResetearPersonaje) = "ResetearPersonaje"
ClientPacketName(ClientPacketID.DeleteItem) = "DeleteItem"
ClientPacketName(ClientPacketID.FinalizarPescaEspecial) = "FinalizarPescaEspecial"
ClientPacketName(ClientPacketID.RomperCania) = "RomperCania"
ClientPacketName(ClientPacketID.UseItemU) = "UseItemU"
ClientPacketName(ClientPacketID.RepeatMacro) = "RepeatMacro"
ClientPacketName(ClientPacketID.BuyShopItem) = "BuyShopItem"
ClientPacketName(ClientPacketID.PerdonFaccion) = "PerdonFaccion"              '/PERDONFACCION NAME
ClientPacketName(ClientPacketID.IniciarCaptura) = "IniciarCaptura"           '/EVENTOCAPTURA PARTICIPANTES CANTIDAD_RONDAS NIVEL_MINIMO PRECIO
ClientPacketName(ClientPacketID.ParticiparCaptura) = "ParticiparCaptura"        '/PARTICIPARCAPTURA
ClientPacketName(ClientPacketID.CancelarCaptura) = "CancelarCaptura"          '/CANCELARCAPTURA
ClientPacketName(ClientPacketID.SeguirMouse) = "SeguirMouse"
ClientPacketName(ClientPacketID.SendPosSeguimiento) = "SendPosSeguimiento"
ClientPacketName(ClientPacketID.NotifyInventarioHechizos) = "NotifyInventarioHechizos"
ClientPacketName(ClientPacketID.PublicarPersonajeMAO) = "PublicarPersonajeMAO"
ClientPacketName(ClientPacketID.CreateAccount) = "CreateAccount"
ClientPacketName(ClientPacketID.LoginAccount) = "LoginAccount"
ClientPacketName(ClientPacketID.DeleteCharacter) = "DeleteCharacter"
End Sub

''
' Handles incoming data.

Public Function HandleIncomingData(ByVal message As Network.Reader) As Boolean
On Error GoTo HandleIncomingData_Err

    Set Reader = message
    
    Dim PacketId As Long
    PacketId = Reader.ReadInt16
    
    #If LogRecv Then
        If PacketId <> ServerPacketID.CharacterMove Then 'Evito este msg pq te floodea todo
            Debug.Print "[RECV] -> " & ServerPacketName(PacketId) & " " & message.GetAvailable() & " bytes."
        End If
    #End If
    
    Select Case PacketId
        Case ServerPacketID.Connected
            Call HandleConnected
        Case ServerPacketID.logged
            Call HandleLogged
        Case ServerPacketID.RemoveDialogs
            Call HandleRemoveDialogs
        Case ServerPacketID.RemoveCharDialog
            Call HandleRemoveCharDialog
        Case ServerPacketID.NavigateToggle
            Call HandleNavigateToggle
        Case ServerPacketID.EquiteToggle
            Call HandleEquiteToggle
        Case ServerPacketID.Disconnect
            Call HandleDisconnect
        Case ServerPacketID.CommerceEnd
            Call HandleCommerceEnd
        Case ServerPacketID.BankEnd
            Call HandleBankEnd
        Case ServerPacketID.CommerceInit
            Call HandleCommerceInit
        Case ServerPacketID.BankInit
            Call HandleBankInit
        Case ServerPacketID.UserCommerceInit
            Call HandleUserCommerceInit
        Case ServerPacketID.UserCommerceEnd
            Call HandleUserCommerceEnd
        Case ServerPacketID.ShowBlacksmithForm
            Call HandleShowBlacksmithForm
        Case ServerPacketID.ShowCarpenterForm
            Call HandleShowCarpenterForm
        Case ServerPacketID.NPCKillUser
            Call HandleNPCKillUser
        Case ServerPacketID.BlockedWithShieldUser
            Call HandleBlockedWithShieldUser
        Case ServerPacketID.BlockedWithShieldOther
            Call HandleBlockedWithShieldOther
        Case ServerPacketID.CharSwing
            Call HandleCharSwing
        Case ServerPacketID.SafeModeOn
            Call HandleSafeModeOn
        Case ServerPacketID.SafeModeOff
            Call HandleSafeModeOff
        Case ServerPacketID.PartySafeOn
            Call HandlePartySafeOn
        Case ServerPacketID.PartySafeOff
            Call HandlePartySafeOff
        Case ServerPacketID.CantUseWhileMeditating
            Call HandleCantUseWhileMeditating
        Case ServerPacketID.UpdateSta
            Call HandleUpdateSta
        Case ServerPacketID.UpdateMana
            Call HandleUpdateMana
        Case ServerPacketID.UpdateHP
            Call HandleUpdateHP
        Case ServerPacketID.UpdateGold
            Call HandleUpdateGold
        Case ServerPacketID.UpdateExp
            Call HandleUpdateExp
        Case ServerPacketID.ChangeMap
            Call HandleChangeMap
        Case ServerPacketID.PosUpdate
            Call HandlePosUpdate
        Case ServerPacketID.PosUpdateUserChar
            Call HandlePosUpdateUserChar
        Case ServerPacketID.PosUpdateChar
            Call HandlePosUpdateChar
        Case ServerPacketID.NPCHitUser
            Call HandleNPCHitUser
        Case ServerPacketID.UserHitNPC
            Call HandleUserHitNPC
        Case ServerPacketID.UserAttackedSwing
            Call HandleUserAttackedSwing
        Case ServerPacketID.UserHittedByUser
            Call HandleUserHittedByUser
        Case ServerPacketID.UserHittedUser
            Call HandleUserHittedUser
        Case ServerPacketID.ChatOverHead
            Call HandleChatOverHead
        Case ServerPacketID.ConsoleMsg
            Call HandleConsoleMessage
        Case ServerPacketID.GuildChat
            Call HandleGuildChat
        Case ServerPacketID.ShowMessageBox
            Call HandleShowMessageBox
        Case ServerPacketID.CharacterCreate
            Call HandleCharacterCreate
        Case ServerPacketID.UpdateFlag
            Call HandleUpdateFlag
        Case ServerPacketID.CharacterRemove
            Call HandleCharacterRemove
        Case ServerPacketID.CharacterMove
            Call HandleCharacterMove
        Case ServerPacketID.UserIndexInServer
            Call HandleUserIndexInServer
        Case ServerPacketID.UserCharIndexInServer
            Call HandleUserCharIndexInServer
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        Case ServerPacketID.ForceCharMoveSiguiendo
            Call HandleForceCharMoveSiguiendo
        Case ServerPacketID.CharacterChange
            Call HandleCharacterChange
        Case ServerPacketID.ObjectCreate
            Call HandleObjectCreate
        Case ServerPacketID.fxpiso
            Call HandleFxPiso
        Case ServerPacketID.ObjectDelete
            Call HandleObjectDelete
        Case ServerPacketID.BlockPosition
            Call HandleBlockPosition
        Case ServerPacketID.PlayMIDI
            Call HandlePlayMIDI
        Case ServerPacketID.PlayWave
            Call HandlePlayWave
        Case ServerPacketID.PlayWaveStep
            Call HandlePlayWaveStep
        Case ServerPacketID.guildList
            Call HandleGuildList
        Case ServerPacketID.AreaChanged
            Call HandleAreaChanged
        Case ServerPacketID.PauseToggle
            Call HandlePauseToggle
        Case ServerPacketID.RainToggle
            Call HandleRainToggle
        Case ServerPacketID.CreateFX
            Call HandleCreateFX
        Case ServerPacketID.CharAtaca
            Call HandleCharAtaca
        Case ServerPacketID.RecievePosSeguimiento
            Call HandleRecievePosSeguimiento
        Case ServerPacketID.CancelarSeguimiento
            Call HandleCancelarSeguimiento
        Case ServerPacketID.GetInventarioHechizos
            Call HandleGetInventarioHechizos
        Case ServerPacketID.NotificarClienteCasteo
            Call HandleNotificarClienteCasteo
        Case ServerPacketID.SendFollowingCharindex
            Call HandleSendFollowingCharindex
        Case ServerPacketID.NotificarClienteSeguido
            Call HandleNotificarClienteSeguido
        Case ServerPacketID.UpdateUserStats
            Call HandleUpdateUserStats
        Case ServerPacketID.WorkRequestTarget
            Call HandleWorkRequestTarget
        Case ServerPacketID.ChangeInventorySlot
            Call HandleChangeInventorySlot
        Case ServerPacketID.InventoryUnlockSlots
            Call HandleInventoryUnlockSlots
        Case ServerPacketID.ChangeBankSlot
            Call HandleChangeBankSlot
        Case ServerPacketID.ChangeSpellSlot
            Call HandleChangeSpellSlot
        Case ServerPacketID.Atributes
            Call HandleAtributes
        Case ServerPacketID.BlacksmithWeapons
            Call HandleBlacksmithWeapons
        Case ServerPacketID.BlacksmithArmors
            Call HandleBlacksmithArmors
        Case ServerPacketID.CarpenterObjects
            Call HandleCarpenterObjects
        Case ServerPacketID.RestOK
            Call HandleRestOK
        Case ServerPacketID.ErrorMsg
            Call HandleErrorMessage
        Case ServerPacketID.Blind
            Call HandleBlind
        Case ServerPacketID.Dumb
            Call HandleDumb
        Case ServerPacketID.ShowSignal
            Call HandleShowSignal
        Case ServerPacketID.ChangeNPCInventorySlot
            Call HandleChangeNPCInventorySlot
        Case ServerPacketID.UpdateHungerAndThirst
            Call HandleUpdateHungerAndThirst
        Case ServerPacketID.MiniStats
            Call HandleMiniStats
        Case ServerPacketID.LevelUp
            Call HandleLevelUp
        Case ServerPacketID.AddForumMsg
            Call HandleAddForumMessage
        Case ServerPacketID.ShowForumForm
            Call HandleShowForumForm
        Case ServerPacketID.SetInvisible
            Call HandleSetInvisible
        Case ServerPacketID.MeditateToggle
            Call HandleMeditateToggle
        Case ServerPacketID.BlindNoMore
            Call HandleBlindNoMore
        Case ServerPacketID.DumbNoMore
            Call HandleDumbNoMore
        Case ServerPacketID.SendSkills
            Call HandleSendSkills
        Case ServerPacketID.TrainerCreatureList
            Call HandleTrainerCreatureList
        Case ServerPacketID.guildNews
            Call HandleGuildNews
        Case ServerPacketID.OfferDetails
            Call HandleOfferDetails
        Case ServerPacketID.AlianceProposalsList
            Call HandleAlianceProposalsList
        Case ServerPacketID.PeaceProposalsList
            Call HandlePeaceProposalsList
        Case ServerPacketID.CharacterInfo
            Call HandleCharacterInfo
        Case ServerPacketID.GuildLeaderInfo
            Call HandleGuildLeaderInfo
        Case ServerPacketID.GuildDetails
            Call HandleGuildDetails
        Case ServerPacketID.ShowGuildFundationForm
            Call HandleShowGuildFundationForm
        Case ServerPacketID.ParalizeOK
            Call HandleParalizeOK
        Case ServerPacketID.ShowUserRequest
            Call HandleShowUserRequest
        Case ServerPacketID.ChangeUserTradeSlot
            Call HandleChangeUserTradeSlot
        'Case ServerPacketID.SendNight
        '    Call HandleSendNight
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
        Case ServerPacketID.FYA
            Call HandleFYA
        Case ServerPacketID.CerrarleCliente
            Call HandleCerrarleCliente
        Case ServerPacketID.Contadores
            Call HandleContadores
        Case ServerPacketID.ShowPapiro
            Call HandleShowPapiro
        Case ServerPacketID.SpawnListt
            Call HandleSpawnList
        Case ServerPacketID.ShowSOSForm
            Call HandleShowSOSForm
        Case ServerPacketID.ShowMOTDEditionForm
            Call HandleShowMOTDEditionForm
        Case ServerPacketID.ShowGMPanelForm
            Call HandleShowGMPanelForm
        Case ServerPacketID.UserNameList
            Call HandleUserNameList
        Case ServerPacketID.UserOnline
            Call HandleUserOnline
        Case ServerPacketID.ParticleFX
            Call HandleParticleFX
        Case ServerPacketID.ParticleFXToFloor
            Call HandleParticleFXToFloor
        Case ServerPacketID.ParticleFXWithDestino
            Call HandleParticleFXWithDestino
        Case ServerPacketID.ParticleFXWithDestinoXY
            Call HandleParticleFXWithDestinoXY
        Case ServerPacketID.Hora
            Call HandleHora
        Case ServerPacketID.Light
            Call HandleLight
        Case ServerPacketID.AuraToChar
            Call HandleAuraToChar
        Case ServerPacketID.SpeedToChar
            Call HandleSpeedToChar
        Case ServerPacketID.LightToFloor
            Call HandleLightToFloor
        Case ServerPacketID.NieveToggle
            Call HandleNieveToggle
        Case ServerPacketID.NieblaToggle
            Call HandleNieblaToggle
        Case ServerPacketID.Goliath
            Call HandleGoliath
        Case ServerPacketID.TextOverChar
            Call HandleTextOverChar
        Case ServerPacketID.TextOverTile
            Call HandleTextOverTile
        Case ServerPacketID.TextCharDrop
            Call HandleTextCharDrop
        Case ServerPacketID.FlashScreen
            Call HandleFlashScreen
        Case ServerPacketID.AlquimistaObj
            Call HandleAlquimiaObjects
        Case ServerPacketID.ShowAlquimiaForm
            Call HandleShowAlquimiaForm
        Case ServerPacketID.SastreObj
            Call HandleSastreObjects
        Case ServerPacketID.ShowSastreForm
            Call HandleShowSastreForm
        Case ServerPacketID.VelocidadToggle
            Call HandleVelocidadToggle
        Case ServerPacketID.MacroTrabajoToggle
            Call HandleMacroTrabajoToggle
        Case ServerPacketID.BindKeys
            Call HandleBindKeys
        Case ServerPacketID.ShowfrmLogear
            Call HandleShowfrmLogear
        Case ServerPacketID.ShowFrmMapa
            Call HandleShowFrmMapa
        Case ServerPacketID.InmovilizadoOK
            Call HandleInmovilizadoOK
        Case ServerPacketID.BarFx
            Call HandleBarFx
        Case ServerPacketID.LocaleMsg
            Call HandleLocaleMsg
        Case ServerPacketID.ShowPregunta
            Call HandleShowPregunta
        Case ServerPacketID.DatosGrupo
            Call HandleDatosGrupo
        Case ServerPacketID.ubicacion
            Call HandleUbicacion
        Case ServerPacketID.ArmaMov
            Call HandleArmaMov
        Case ServerPacketID.EscudoMov
            Call HandleEscudoMov
        Case ServerPacketID.ViajarForm
            Call HandleViajarForm
        Case ServerPacketID.NadarToggle
            Call HandleNadarToggle
        Case ServerPacketID.ShowFundarClanForm
            Call HandleShowFundarClanForm
        Case ServerPacketID.CharUpdateHP
            Call HandleCharUpdateHP
        Case ServerPacketID.CharUpdateMAN
            Call HandleCharUpdateMAN
        Case ServerPacketID.PosLLamadaDeClan
            Call HandlePosLLamadaDeClan
        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails
        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend
        Case ServerPacketID.NpcQuestListSend
            Call HandleNpcQuestListSend
        Case ServerPacketID.UpdateNPCSimbolo
            Call HandleUpdateNPCSimbolo
        Case ServerPacketID.ClanSeguro
            Call HandleClanSeguro
        Case ServerPacketID.Intervals
            Call HandleIntervals
        Case ServerPacketID.UpdateUserKey
            Call HandleUpdateUserKey
        Case ServerPacketID.UpdateRM
            Call HandleUpdateRM
        Case ServerPacketID.UpdateDM
            Call HandleUpdateDM
        Case ServerPacketID.SeguroResu
            Call HandleSeguroResu
        Case ServerPacketID.Stopped
            Call HandleStopped
        Case ServerPacketID.InvasionInfo
            Call HandleInvasionInfo
        Case ServerPacketID.CommerceRecieveChatMessage
            Call HandleCommerceRecieveChatMessage
        Case ServerPacketID.DoAnimation
            Call HandleDoAnimation
        Case ServerPacketID.OpenCrafting
            Call HandleOpenCrafting
        Case ServerPacketID.CraftingItem
            Call HandleCraftingItem
        Case ServerPacketID.CraftingCatalyst
            Call HandleCraftingCatalyst
        Case ServerPacketID.CraftingResult
            Call HandleCraftingResult
        Case ServerPacketID.ForceUpdate
            Call HandleForceUpdate
        Case ServerPacketID.AnswerReset
            Call HandleAnswerReset
        Case ServerPacketID.ObjQuestListSend
            Call HandleObjQuestListSend
        Case ServerPacketID.UpdateBankGld
            Call HandleUpdateBankGld
        Case ServerPacketID.PelearConPezEspecial
            Call HandlePelearConPezEspecial
        Case ServerPacketID.Privilegios
            Call HandlePrivilegios
        Case ServerPacketID.ShopInit
            Call HandleShopInit
        Case ServerPacketID.ShopPjsInit
            Call HandleShopPjsInit
        Case ServerPacketID.UpdateShopClienteCredits
            Call HandleUpdateShopClienteCredits
        Case ServerPacketID.AccountCharacterList
            Call HandleAccountCharacterList
        Case ServerPacketID.ComboCooldown
            Call HandleComboCooldown
        Case Else
            err.Raise &HDEADBEEF, "Invalid Message"
    End Select
    
    If (message.GetAvailable() > 0) Then
        err.Raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketId & "' se encuentra en mal estado con '" & message.GetAvailable() & "' bytes de mas"
    End If

    HandleIncomingData = True
    
HandleIncomingData_Err:
    
    Set Reader = Nothing

    If err.Number <> 0 Then
        Call RegistrarError(err.Number, err.Description & ". PacketID: " & PacketId, "Protocol.HandleIncomingData", Erl)
       ' Call modNetwork.Disconnect
        
        HandleIncomingData = False
    End If

End Function

''
' Handles the Connected message.

Private Sub HandleConnected()

    frmMain.ShowFPS.enabled = True

    Call Login
    
End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleLogged_Err
 
    ' Variable initialization
    newUser = Reader.ReadBool
    
    UserCiego = False
    EngineRun = True
    UserDescansar = False
    Nombres = True
    Pregunta = False

    frmMain.stabar.visible = True
    
    frmMain.HpBar.visible = True

    If UserMaxMAN <> 0 Then
        frmMain.manabar.visible = True

    End If

    frmMain.hambar.visible = True
    frmMain.AGUbar.visible = True
    frmMain.Hpshp.visible = (UserMinHp > 0)
    frmMain.MANShp.visible = (UserMinMAN > 0)
    frmMain.STAShp.visible = (UserMinSTA > 0)
    frmMain.AGUAsp.visible = (UserMinAGU > 0)
    frmMain.COMIDAsp.visible = (UserMinHAM > 0)
    frmMain.GldLbl.visible = True
    ' frmMain.Label6.Visible = True
    frmMain.Fuerzalbl.visible = True
    frmMain.AgilidadLbl.visible = True
    frmMain.imgDeleteItem.visible = True
    QueRender = 0
     lFrameTimer = 0
     FramesPerSecCounter = 0
    
    frmMain.ImgSegParty = LoadInterface("boton-seguro-party-on.bmp")
    frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
    frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    SeguroParty = True
    SeguroClanX = True
    SeguroResuX = True
    
    'Set connected state
    
    Call SetConnected
    
    'Show tip
    
    Exit Sub

HandleLogged_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleLogged", Erl)
    
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRemoveDialogs_Err

    Call Dialogos.RemoveAllDialogs
    
    Exit Sub

HandleRemoveDialogs_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleRemoveDialogs", Erl)
    
    
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
    
    On Error GoTo HandleRemoveCharDialog_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Call Dialogos.RemoveDialog(Reader.ReadInt16())
    
    Exit Sub

HandleRemoveCharDialog_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleRemoveCharDialog", Erl)
    
    
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNavigateToggle_Err

    UserNavegando = Reader.ReadBool()
    
    Exit Sub

HandleNavigateToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNavigateToggle", Erl)
    
    
End Sub

Private Sub HandleNadarToggle()
    
    On Error GoTo HandleNadarToggle_Err

    UserNadando = Reader.ReadBool()
    UserNadandoTrajeCaucho = Reader.ReadBool()
    
    Exit Sub

HandleNadarToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNadarToggle", Erl)
    
    
End Sub

Private Sub HandleEquiteToggle()
 
    On Error GoTo HandleEquiteToggle_Err
    
    UserMontado = Not UserMontado

    Exit Sub

HandleEquiteToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleEquiteToggle", Erl)
    
    
End Sub

Private Sub HandleVelocidadToggle()
    
    On Error GoTo HandleVelocidadToggle_Err

    If UserCharIndex = 0 Then Exit Sub
    
    charlist(UserCharIndex).Speeding = Reader.ReadReal32()
    
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar / charlist(UserCharIndex).Speeding)
    
    Exit Sub

HandleVelocidadToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleVelocidadToggle", Erl)
    
    
End Sub

Private Sub HandleMacroTrabajoToggle()
    'Activa o Desactiva el macro de trabajo  06/07/2014 Ladder
    
    On Error GoTo HandleMacroTrabajoToggle_Err

    Dim activar As Boolean
    activar = Reader.ReadBool()

    If activar = False Then
    
        Call ResetearUserMacro
        
    Else
    
        Call AddToConsole("Has comenzado a trabajar...", 2, 223, 51, 1, 0)
        
        frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
        frmMain.MacroLadder.enabled = True
        
        UserMacro.Intervalo = IntervaloTrabajoConstruir
        UserMacro.Activado = True
        UserMacro.cantidad = 999
        UserMacro.TIPO = 6
        
        TargetXMacro = tX
        TargetYMacro = tY

    End If
    
    Exit Sub

HandleMacroTrabajoToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleMacroTrabajoToggle", Erl)
    
    
End Sub

''
' Handles the Disconnect message.

Public Sub HandleDisconnect()
    
    On Error GoTo HandleDisconnect_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long
    
    If (Not Reader Is Nothing) Then
    FullLogout = Reader.ReadBool
    End If

    Call WriteVar(RESOURCES_PATH & "/OUTPUT/" & "Configuracion.ini", "OPCIONES", "LastScroll", hlst.Scroll)

    Mod_Declaraciones.Connected = False
    
    Call ResetearUserMacro

    'Close connection
    Call modNetwork.Disconnect
    
    'Hide main form
    'FrmCuenta.Visible = True
    Call resetearCartel
    frmConnect.visible = True
    QueRender = 1
    isLogged = False
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    
    ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

    frmMain.picHechiz.visible = False
    
    frmMain.UpdateLight.enabled = False
    frmMain.UpdateDaytime.enabled = False
    
    frmMain.visible = False
    
    Seguido = False
    
    OpcionMenu = 0

    frmMain.picInv.visible = True
    frmMain.picHechiz.visible = False

    frmMain.cmdlanzar.visible = False
    'frmMain.lblrefuerzolanzar.Visible = False
    frmMain.cmdMoverHechi(0).visible = False
    frmMain.cmdMoverHechi(1).visible = False
    
    QuePesta�aInferior = 0
    frmMain.stabar.visible = True
    frmMain.HpBar.visible = True
    frmMain.manabar.visible = True
    frmMain.hambar.visible = True
    frmMain.AGUbar.visible = True
    frmMain.Hpshp.visible = True
    frmMain.MANShp.visible = True
    frmMain.STAShp.visible = True
    frmMain.AGUAsp.visible = True
    frmMain.COMIDAsp.visible = True
    frmMain.GldLbl.visible = True
    frmMain.Fuerzalbl.visible = True
    frmMain.AgilidadLbl.visible = True
    frmMain.lblWeapon.visible = True
    frmMain.lblShielder.visible = True
    frmMain.lblHelm.visible = True
    frmMain.lblArmor.visible = True
    frmMain.lblResis.visible = True
    frmMain.lbldm.visible = True
    frmMain.ImgSeg.visible = False
    frmMain.ImgSegParty.visible = False
    frmMain.ImgSegClan.visible = False
    frmMain.ImgSegResu.visible = False
    initPacketControl
    'Stop audio
    If Sonido Then
        Sound.Sound_Stop_All
        Sound.Ambient_Stop

    End If

    Call CleanDialogs
    
    'frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form
    UserMap = 1
    MapSize = MapSize1
    
    Entraday = 1
    Entradax = 1
    Call EraseChar(UserCharIndex, True)
    Call SwitchMap(UserMap)
    
    
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    MiCabeza = 0
    UserHogar = 0

    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    For i = 1 To UserInvUnlocked
        frmMain.imgInvLock(i - 1).Picture = Nothing
    Next i

    For i = 1 To MAX_INVENTORY_SLOTS
        Call frmMain.Inventario.ClearSlot(i)
        Call frmBancoObj.InvBankUsu.ClearSlot(i)
        Call frmComerciar.InvComNpc.ClearSlot(i)
        Call frmComerciar.InvComUsu.ClearSlot(i)
        Call frmBancoCuenta.InvBankUsuCuenta.ClearSlot(i)
        Call frmComerciarUsu.InvUser.ClearSlot(i)
        Call frmCrafteo.InvCraftUser.ClearSlot(i)
    Next i

    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        Call frmBancoObj.InvBoveda.ClearSlot(i)
    Next i

    For i = 1 To MAX_KEYS
        Call FrmKeyInv.InvKeys.ClearSlot(i)
    Next i

    For i = 1 To MAX_SLOTS_CRAFTEO
        Call frmCrafteo.InvCraftItems.ClearSlot(i)
    Next i

    Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)
    
    UserInvUnlocked = 0

    Alocados = 0

    'Reset global vars
    UserParalizado = False
    UserSaliendo = False
    UserStopped = False
    UserInmovilizado = False
    pausa = False
    UserMeditar = False
    UserDescansar = False
    UserNavegando = False
    UserMontado = False
    UserNadando = False
    UserNadandoTrajeCaucho = False
    bRain = False
    AlphaNiebla = 30
    frmMain.TimerNiebla.enabled = False
    bNiebla = False
    bNieve = False
    bFogata = False
    SkillPoints = 0
    UserEstado = 0
    
    InviCounter = 0
    DrogaCounter = 0
     
    frmMain.Contadores.enabled = False
    
    InvasionActual = 0
    frmMain.Evento.enabled = False
     
    'Delete all kind of dialogs
    
    'Reset some char variables...
    For i = 1 To LastChar + 1
        charlist(i).Invisible = False
        charlist(i).Arma_Aura = ""
        charlist(i).Body_Aura = ""
        charlist(i).Escudo_Aura = ""
        charlist(i).DM_Aura = ""
        charlist(i).RM_Aura = ""
        charlist(i).Otra_Aura = ""
        charlist(i).Head_Aura = ""
        charlist(i).Speeding = 0
        charlist(i).AuraAngle = 0
    Next i

    For i = 1 To LastChar + 1
        charlist(i).dialog = ""
    Next i
        
    'Unload all forms except frmMain and frmConnect
    Dim Frm As Form
    
    For Each Frm In Forms

        If Frm.Name <> frmMain.Name And Frm.Name <> frmConnect.Name And Frm.Name <> frmMensaje.Name Then
            Unload Frm

        End If

    Next
    
    frmConnect.Show
    FrmLogear.Show , frmConnect

    Exit Sub
HandleDisconnect_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleDisconnect", Erl)
    
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleCommerceEnd_Err

    'Reset vars
    Comerciando = False
    
    'Hide form
    ' Unload frmComerciar
    
    Exit Sub

HandleCommerceEnd_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCommerceEnd", Erl)
    
    
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleBankEnd_Err

    'Unload frmBancoObj
    Comerciando = False
    
    Exit Sub

HandleBankEnd_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBankEnd", Erl)
    
    
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
    
    On Error GoTo HandleCommerceInit_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i       As Long

    Dim NpcName As String

    NpcName = Reader.ReadString8()

    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        With frmMain.Inventario
            Call frmComerciar.InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmComerciar.PicInvUser)
    frmComerciar.Show , frmMain
    frmComerciar.Refresh
    
    Exit Sub

HandleCommerceInit_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCommerceInit", Erl)
    
    
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
    
    On Error GoTo HandleBankInit_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long

    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        With frmMain.Inventario
            Call frmBancoObj.InvBankUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True

    frmBancoObj.lblcosto = PonerPuntos(UserGLD)
    frmBancoObj.Show , frmMain
    frmBancoObj.Refresh
    
    Exit Sub

HandleBankInit_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBankInit", Erl)
    
    
End Sub

Private Sub HandleGoliath()
    
    On Error GoTo HandleGoliathInit_Err

    '***************************************************
    '
    '***************************************************

    Dim UserBoveOro As Long

    Dim UserInvBove As Byte
    
    UserBoveOro = Reader.ReadInt32()
    UserInvBove = Reader.ReadInt8()
    Call frmGoliath.ParseBancoInfo(UserBoveOro, UserInvBove)
    
    Exit Sub

HandleGoliathInit_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleGoliathInit", Erl)
    
    
End Sub

Private Sub HandleShowfrmLogear()
    
    On Error GoTo HandleShowfrmLogear_Err

    '***************************************************
    '
    '***************************************************
    FrmLogear.Show , frmConnect
    
    Exit Sub

HandleShowfrmLogear_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowfrmLogear", Erl)
    
    
End Sub

Private Sub HandleShowFrmMapa()
    
    On Error GoTo HandleShowFrmMapa_Err

    '***************************************************
    '
    '***************************************************
    ExpMult = Reader.ReadInt16()
    OroMult = Reader.ReadInt16()
    
    Call frmMapaGrande.CalcularPosicionMAPA

    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain
    
    Exit Sub

HandleShowFrmMapa_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowFrmMapa", Erl)
    
    
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
    
    On Error GoTo HandleUserCommerceInit_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long
    
    'Clears lists if necessary
    
    'Fill inventory list
    With frmMain.Inventario

        For i = 1 To MAX_INVENTORY_SLOTS
            frmComerciarUsu.InvUser.SetItem i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), 0, 0, 0, 0, .ItemName(i), 0
        Next i

    End With
        
    frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserGLD)
    
    Dim j As Byte

    For j = 1 To 6
        Call frmComerciarUsu.InvOtherSell.SetItem(j, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciarUsu.InvUserSell.SetItem(j, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next j
    
    'Set state and show form
    Comerciando = True
    
    frmComerciarUsu.Show , frmMain
    
    Exit Sub

HandleUserCommerceInit_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserCommerceInit", Erl)
    
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleUserCommerceEnd_Err
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
    
    Exit Sub

HandleUserCommerceEnd_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserCommerceEnd", Erl)
    
    
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleShowBlacksmithForm_Err
    
    If frmMain.macrotrabajo.enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftBlacksmith(MacroBltIndex)
        
    Else
    
        frmHerrero.lstArmas.Clear

        Dim i As Byte

        For i = 0 To UBound(CascosHerrero())

            If CascosHerrero(i).Index = 0 Then Exit For
            Call frmHerrero.lstArmas.AddItem(ObjData(CascosHerrero(i).Index).Name)
        Next i

        frmHerrero.Command3.Picture = LoadInterface("boton-casco-over.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        Call Establecer_Borde(frmHerrero.lstArmas, frmHerrero, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmHerrero.List1, frmHerrero, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmHerrero.List2, frmHerrero, COLOR_AZUL, 0, 0)
        frmHerrero.Show , frmMain

    End If
    
    Exit Sub

HandleShowBlacksmithForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowBlacksmithForm", Erl)
    
    
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleShowCarpenterForm_Err
        
   ' If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
    
        'Call WriteCraftCarpenter(MacroBltIndex)
        
   ' Else
         
        COLOR_AZUL = RGB(0, 0, 0)
    
        ' establece el borde al listbox
        Call Establecer_Borde(frmCarp.lstArmas, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List1, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List2, frmCarp, COLOR_AZUL, 0, 0)
        frmCarp.Show , frmMain

   ' End If
    
    Exit Sub

HandleShowCarpenterForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowCarpenterForm", Erl)
    
    
End Sub

Private Sub HandleShowAlquimiaForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleShowAlquimiaForm_Err
    
    If frmMain.macrotrabajo.enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftAlquimista(MacroBltIndex)
        
    Else
    
        frmAlqui.Picture = LoadInterface("alquimia.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        
        ' establece el borde al listbox
        Call Establecer_Borde(frmAlqui.lstArmas, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List1, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List2, frmAlqui, COLOR_AZUL, 1, 1)

        frmAlqui.Show , frmMain

    End If
    
    Exit Sub

HandleShowAlquimiaForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowAlquimiaForm", Erl)
    
    
End Sub

Private Sub HandleShowSastreForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleShowSastreForm_Err
        
    If frmMain.macrotrabajo.enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftSastre(MacroBltIndex)
        
    Else
    
        COLOR_AZUL = RGB(0, 0, 0)

        ' establece el borde al listbox
        Call Establecer_Borde(FrmSastre.lstArmas, FrmSastre, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(FrmSastre.List1, FrmSastre, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(FrmSastre.List2, FrmSastre, COLOR_AZUL, 1, 1)
        FrmSastre.Picture = LoadInterface("sastreria.bmp")

        Dim i As Byte

        FrmSastre.lstArmas.Clear

        For i = 1 To UBound(SastreRopas())

            If SastreRopas(i).Index = 0 Then Exit For
            FrmSastre.lstArmas.AddItem (ObjData(SastreRopas(i).Index).Name)
        Next i
    
        FrmSastre.Command1.Picture = LoadInterface("sastreria_vestimentahover.bmp")
        FrmSastre.Show , frmMain

    End If
    
    Exit Sub

HandleShowSastreForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowSastreForm", Erl)
    
    
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleNPCKillUser_Err
        
    Call AddToConsole(MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleNPCKillUser_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNPCKillUser", Erl)
    
    
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleBlockedWithShieldUser_Err
        
    Call AddToConsole(MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleBlockedWithShieldUser_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlockedWithShieldUser", Erl)
    
    
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleBlockedWithShieldOther_Err
        
    Call AddToConsole(MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleBlockedWithShieldOther_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlockedWithShieldOther", Erl)
    
    
End Sub

''
' Handles the UserSwing message.

Private Sub HandleCharSwing()
    
    On Error GoTo HandleCharSwing_Err
    
    Dim charindex As Integer

    charindex = Reader.ReadInt16
    
    Dim ShowFX As Boolean

    ShowFX = Reader.ReadBool
    
    Dim ShowText As Boolean

    ShowText = Reader.ReadBool
    
    Dim NotificoTexto As Boolean
    
    NotificoTexto = Reader.ReadBool
        
    With charlist(charindex)

        If ShowText And NotificoTexto Then
            Call SetCharacterDialogFx(charindex, IIf(charindex = UserCharIndex, "Fallas", "Fall�"), RGBA_From_Comp(255, 0, 0))

        End If
        
        Call Sound.Sound_Play(2, False, Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y)) ' Swing
        
        ' If ShowFX And .Invisible = False Then Call SetCharacterFx(charindex, 90, 0)
         
        
    End With
    
    Exit Sub

HandleCharSwing_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharSwing", Erl)
    
    
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleSafeModeOn_Err
        
    Call frmMain.DibujarSeguro
    Call AddToConsole(MENSAJE_SEGURO_ACTIVADO, 65, 190, 156, False, False, False)
    
    Exit Sub

HandleSafeModeOn_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSafeModeOn", Erl)
    
    
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleSafeModeOff_Err
    
    Call frmMain.DesDibujarSeguro
    Call AddToConsole(MENSAJE_SEGURO_DESACTIVADO, 65, 190, 156, False, False, False)
    
    Exit Sub

HandleSafeModeOff_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSafeModeOff", Erl)
    
    
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandlePartySafeOff()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    
    On Error GoTo HandlePartySafeOff_Err
    
    Call frmMain.ControlSeguroParty(False)
    Call AddToConsole(MENSAJE_SEGURO_PARTY_OFF, 250, 250, 0, False, True, False)
    
    Exit Sub

HandlePartySafeOff_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePartySafeOff", Erl)
    
    
End Sub

Private Sub HandleClanSeguro()
    
    On Error GoTo HandleClanSeguro_Err

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    Dim Seguro As Boolean
    
    'Get data and update form
    Seguro = Reader.ReadBool()
    
    If SeguroClanX Then
    
        Call AddToConsole("Seguro de clan desactivado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-off.bmp")
        SeguroClanX = False
        
    Else
        Call AddToConsole("Seguro de clan activado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
        SeguroClanX = True

    End If
    
    Exit Sub

HandleClanSeguro_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleClanSeguro", Erl)
    
    
End Sub

Private Sub HandleIntervals()
    
    On Error GoTo HandleIntervals_Err

    IntervaloArco = Reader.ReadInt32()
    IntervaloCaminar = Reader.ReadInt32()
    IntervaloGolpe = Reader.ReadInt32()
    IntervaloGolpeMagia = Reader.ReadInt32()
    IntervaloMagia = Reader.ReadInt32()
    IntervaloMagiaGolpe = Reader.ReadInt32()
    IntervaloGolpeUsar = Reader.ReadInt32()
    IntervaloTrabajoExtraer = Reader.ReadInt32()
    IntervaloTrabajoConstruir = Reader.ReadInt32()
    IntervaloUsarU = Reader.ReadInt32()
    IntervaloUsarClic = Reader.ReadInt32()
    IntervaloTirar = Reader.ReadInt32()
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, IntervaloGolpe)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, IntervaloUsarU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, IntervaloUsarClic)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, IntervaloMagia)
    Call MainTimer.SetInterval(TimersIndex.Arrows, IntervaloArco)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, IntervaloMagiaGolpe)
    Call MainTimer.SetInterval(TimersIndex.AttackSpell, IntervaloGolpeMagia)
    Call MainTimer.SetInterval(TimersIndex.AttackUse, IntervaloGolpeUsar)
    Call MainTimer.SetInterval(TimersIndex.Drop, IntervaloTirar)
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar)

    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.AttackSpell)
    Call MainTimer.Start(TimersIndex.AttackUse)
    Call MainTimer.Start(TimersIndex.Drop)
    Call MainTimer.Start(TimersIndex.Walk)
    
    Exit Sub

HandleIntervals_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleIntervals", Erl)
    
    
End Sub

Private Sub HandleUpdateUserKey()
    
    On Error GoTo HandleUpdateUserKey_Err
 
    Dim Slot As Integer, Llave As Integer
    
    Slot = Reader.ReadInt16
    Llave = Reader.ReadInt16

    Call FrmKeyInv.InvKeys.SetItem(Slot, Llave, 1, 0, ObjData(Llave).GrhIndex, eObjType.otLlaves, 0, 0, 0, 0, ObjData(Llave).Name, 0)
    
    Exit Sub

HandleUpdateUserKey_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateUserKey", Erl)
    
    
End Sub

Private Sub HandleUpdateDM()
    
    On Error GoTo HandleUpdateDM_Err
 
    Dim Value As Integer

    Value = Reader.ReadInt16

    frmMain.lbldm = "+" & Value & "%"
    
    Exit Sub

HandleUpdateDM_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateDM", Erl)
    
    
End Sub

Private Sub HandleUpdateRM()
    
    On Error GoTo HandleUpdateRM_Err
 
    Dim Value As Integer

    Value = Reader.ReadInt16

    frmMain.lblResis = "+" & Value
    
    Exit Sub

HandleUpdateRM_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateRM", Erl)
    
    
End Sub

' Handles the ResuscitationSafeOn message.
Private Sub HandlePartySafeOn()

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    On Error GoTo HandlePartySafeOn_Err

    Call frmMain.ControlSeguroParty(True)
    Call AddToConsole(MENSAJE_SEGURO_PARTY_ON, 250, 250, 0, False, True, False)
    
    Exit Sub

HandlePartySafeOn_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePartySafeOn", Erl)
    
    
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleCantUseWhileMeditating_Err

    Call AddToConsole(MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
    
    Exit Sub

HandleCantUseWhileMeditating_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCantUseWhileMeditating", Erl)
    
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
    
    On Error GoTo HandleUpdateSta_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Get data and update form
    UserMinSTA = Reader.ReadInt16()
    frmMain.STAShp.Width = UserMinSTA / UserMaxSTA * 97
    frmMain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA

    If QuePesta�aInferior = 0 Then
        frmMain.STAShp.visible = (UserMinSTA > 0)

    End If
    
    Exit Sub

HandleUpdateSta_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateSta", Erl)
    
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
    
    On Error GoTo HandleUpdateMana_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Dim OldMana As Integer
    OldMana = UserMinMAN
    
    'Get data and update form
    UserMinMAN = Reader.ReadInt16()
    
    If UserMeditar And UserMinMAN - OldMana > 0 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Has ganado " & UserMinMAN - OldMana & " de man�.", .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    If UserMaxMAN > 0 Then
        frmMain.MANShp.Width = UserMinMAN / UserMaxMAN * 243
        frmMain.manabar.Caption = UserMinMAN & " / " & UserMaxMAN

        If QuePesta�aInferior = 0 Then
            frmMain.MANShp.visible = (UserMinMAN > 0)
            frmMain.manabar.visible = True

        End If

    Else
        frmMain.MANShp.Width = 0
        frmMain.manabar.visible = False
        frmMain.MANShp.visible = False

    End If
    
    Exit Sub

HandleUpdateMana_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateMana", Erl)
    
    
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
    
    On Error GoTo HandleUpdateHP_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Dim NuevoValor As Long
    NuevoValor = Reader.ReadInt16()
    
    
    'Get data and update form
    UserMinHp = NuevoValor
    frmMain.Hpshp.Width = UserMinHp / UserMaxHp * 243
    frmMain.HpBar.Caption = UserMinHp & " / " & UserMaxHp
    
    If QuePesta�aInferior = 0 Then
        frmMain.Hpshp.visible = (UserMinHp > 0)

    End If
    
    'Velocidad de la musica
    
    'Is the user alive??
    If UserMinHp = 0 Then
        UserEstado = 1
        charlist(UserCharIndex).Invisible = False
        If MostrarTutorial And tutorial_index <= 0 Then
            If tutorial(e_tutorialIndex.TUTORIAL_Muerto).Activo = 1 Then
                tutorial_index = e_tutorialIndex.TUTORIAL_Muerto
                Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).grh, -1, &H164B8A, , , False, 100, 629, 100, 685, 640, 530, 50, 100)
            End If
        End If
        DrogaCounter = 0
        Call deleteCharIndexs
    Else
        UserEstado = 0

    End If
    
    Exit Sub

HandleUpdateHP_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateHP", Erl)
    
    
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
    
    On Error GoTo HandleUpdateGold_Err

    '***************************************************
    'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    '- 08/14/07: Added GldLbl color variation depending on User Gold and Level
    '***************************************************

    'Get data and update form
    UserGLD = Reader.ReadInt32()
    OroPorNivel = Reader.ReadInt32()
    
    frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
    
    'If UserGLD > UserLvl * OroPorNivel Then
    If UserGLD <= 100000 Then
        frmMain.GldLbl.ForeColor = vbRed
    Else
        frmMain.GldLbl.ForeColor = &H80FFFF
    End If
    
    Exit Sub

HandleUpdateGold_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateGold", Erl)
    
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
    
    On Error GoTo HandleUpdateExp_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Get data and update form
    UserExp = Reader.ReadInt32()

    If UserPasarNivel > 0 Then
        frmMain.EXPBAR.Width = UserExp / UserPasarNivel * 235
        frmMain.lblPorcLvl.Caption = Round(UserExp * (100 / UserPasarNivel), 2) & "%"
        frmMain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
        
    Else
        frmMain.EXPBAR.Width = 235
        frmMain.lblPorcLvl.Caption = "�Nivel m�ximo!"
        frmMain.exp.Caption = "�Nivel m�ximo!"

    End If
    
    Exit Sub

HandleUpdateExp_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateExp", Erl)
    
    
End Sub

Private Sub HandleChangeMap()
    On Error GoTo HandleChangeMap_Err
    UserMap = Reader.ReadInt16()
    
    
    'Hay que resetear el mapa para que lo vuelva a cargar
    Dim x As Integer
    Dim y As Integer
    For x = 1 To MAP_MAX_X
        For y = 1 To MAP_MAX_Y
            MapData(x, y).x = 0
            MapData(x, y).y = 0
        Next y
    Next x
    
 
    If bRain Then
        If Not MapDat.Lluvia Then
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    If frmComerciar.visible Then Unload frmComerciar
    If frmBancoObj.visible Then Unload frmBancoObj
    If frmEstadisticas.visible Then Unload frmEstadisticas
    If frmStatistics.visible Then Unload frmStatistics
    If frmHerrero.visible Then Unload frmHerrero
    If FrmSastre.visible Then Unload FrmSastre
    If frmAlqui.visible Then Unload frmAlqui
    If frmCarp.visible Then Unload frmCarp
    If FrmGrupo.visible Then Unload FrmGrupo
    If frmGoliath.visible Then Unload frmGoliath
    If FrmViajes.visible Then Unload FrmViajes
    If frmCantidad.visible Then Unload frmCantidad
    Call SwitchMap(UserMap)
    If UserMap = 1 Then
        MapSize = MapSize1
    Else
        MapSize = MapSize2
    End If
    Exit Sub

HandleChangeMap_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChangeMap", Erl)
    
    
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
    
    On Error GoTo HandlePosUpdate_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Remove char from old position
    If MapData(rrX(UserPos.x), rrY(UserPos.y)).charindex = UserCharIndex Then
        MapData(rrX(UserPos.x), rrY(UserPos.y)).charindex = 0

    End If
    
    'Set new pos
    UserPos.x = Reader.ReadInt16()
    UserPos.y = Reader.ReadInt16()
    checkZona
    'Set char
    MapData(rrX(UserPos.x), rrY(UserPos.y)).charindex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos
        
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
                
    'Update pos label and minimap
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y


    
    
    Call RefreshAllChars
    
    Exit Sub

HandlePosUpdate_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePosUpdate", Erl)
    
    
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdateUserChar()
    
    On Error GoTo HandlePosUpdateUserChar_Err

    
    Dim temp_x As Integer, temp_y As Integer
    
    
    temp_x = UserPos.x
    temp_y = UserPos.y
    'Set new pos
    UserPos.x = Reader.ReadInt16()
    UserPos.y = Reader.ReadInt16()
    checkZona
    Dim charindex As Integer
    charindex = Reader.ReadInt16()
    
        'Remove char from old position
    If MapData(rrX(temp_x), rrY(temp_y)).charindex = charindex Then
        MapData(rrX(temp_x), rrY(temp_y)).charindex = 0
    End If
    'Set char
    MapData(rrX(UserPos.x), rrY(UserPos.y)).charindex = charindex
    charlist(charindex).Pos = UserPos
        
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
                

    Call RefreshAllChars
    
    Exit Sub

HandlePosUpdateUserChar_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePosUpdateUserChar", Erl)
        
End Sub
''
' Handles the NPCHitUser message.

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdateChar()
    
    On Error GoTo HandlePosUpdateChar_Err

    Dim charindex As Integer
    Dim x As Integer, y As Integer
    
    'Set new pos
    charindex = Reader.ReadInt16()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If charindex = 0 Then Exit Sub
    
    If charlist(charindex).Pos.x > 0 And charlist(charindex).Pos.y > 0 Then
    
        If MapData(rrX(charlist(charindex).Pos.x), rrY(charlist(charindex).Pos.y)).charindex = charindex Then
            MapData(rrX(charlist(charindex).Pos.x), rrY(charlist(charindex).Pos.y)).charindex = 0
        End If
        
        MapData(rrX(x), rrY(y)).charindex = charindex
        charlist(charindex).Pos.x = x
        charlist(charindex).Pos.y = y
    
    End If
    
    Exit Sub

HandlePosUpdateChar_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePosUpdateChar", Erl)
        
End Sub
''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
    
    On Error GoTo HandleNPCHitUser_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Lugar As Byte, Da�oStr As String
    
    Lugar = Reader.ReadInt8()

    Da�oStr = PonerPuntos(Reader.ReadInt16)

    Select Case Lugar

        Case bCabeza
            Call AddToConsole(MENSAJE_GOLPE_CABEZA & Da�oStr, 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddToConsole(MENSAJE_GOLPE_BRAZO_IZQ & Da�oStr, 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddToConsole(MENSAJE_GOLPE_BRAZO_DER & Da�oStr, 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddToConsole(MENSAJE_GOLPE_PIERNA_IZQ & Da�oStr, 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddToConsole(MENSAJE_GOLPE_PIERNA_DER & Da�oStr, 255, 0, 0, True, False, False)

        Case bTorso
            Call AddToConsole(MENSAJE_GOLPE_TORSO & Da�oStr, 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleNPCHitUser_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNPCHitUser", Erl)
    
    
End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
    
    On Error GoTo HandleUserHitNPC_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Call AddToConsole(MENSAJE_GOLPE_CRIATURA_1 & PonerPuntos(Reader.ReadInt32()) & MENSAJE_2, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleUserHitNPC_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserHitNPC", Erl)
    
    
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
    
    On Error GoTo HandleUserAttackedSwing_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Call AddToConsole(MENSAJE_1 & charlist(Reader.ReadInt16()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleUserAttackedSwing_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserAttackedSwing", Erl)
    
    
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
    
    On Error GoTo HandleUserHittedByUser_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim attacker As String
    Dim intt     As Integer
    
    intt = Reader.ReadInt16()
    
    Dim Pos As String

    Pos = InStr(charlist(intt).nombre, "<")
    
    If Pos = 0 Then Pos = Len(charlist(intt).nombre) + 2
    
    attacker = Left$(charlist(intt).nombre, Pos - 2)
    
    Dim Lugar As Byte
    Lugar = Reader.ReadInt8
    
    Dim Da�oStr As String
    Da�oStr = PonerPuntos(Reader.ReadInt16())
    
    Select Case Lugar

        Case bCabeza
            Call AddToConsole(attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddToConsole(attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddToConsole(attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddToConsole(attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddToConsole(attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bTorso
            Call AddToConsole(attacker & MENSAJE_RECIVE_IMPACTO_TORSO & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleUserHittedByUser_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserHittedByUser", Erl)
    
    
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
    
    On Error GoTo HandleUserHittedUser_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim victim As String
    
    Dim intt   As Integer
    
    intt = Reader.ReadInt16()
    'attacker = charlist().Nombre
    
    Dim Pos As String

    Pos = InStr(charlist(intt).nombre, "<")
    
    If Pos = 0 Then Pos = Len(charlist(intt).nombre) + 2
    
    victim = Left$(charlist(intt).nombre, Pos - 2)
    
    Dim Lugar As Byte
    Lugar = Reader.ReadInt8()
    
    Dim Da�oStr As String
    Da�oStr = PonerPuntos(Reader.ReadInt16())
    
    Select Case Lugar

        Case bCabeza
            Call AddToConsole(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddToConsole(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddToConsole(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddToConsole(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddToConsole(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bTorso
            Call AddToConsole(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_TORSO & Da�oStr & MENSAJE_2, 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleUserHittedUser_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserHittedUser", Erl)
    
    
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    Dim chat       As String

    Dim charindex  As Integer

    Dim r          As Byte

    Dim G          As Byte

    Dim B          As Byte

    Dim colortexto As Long

    Dim QueEs      As String

    Dim EsSpell    As Boolean
    
    Dim x As Integer, y As Integer
    chat = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    
    r = Reader.ReadInt8()
    G = Reader.ReadInt8()
    B = Reader.ReadInt8()
    
    colortexto = vbColor_2_Long(Reader.ReadInt32())
    EsSpell = Reader.ReadBool()
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(rrX(x), rrY(y)).charindex = charindex
            End If
        End With
    End If
    
    'Optimizacion de protocolo por Ladder
    QueEs = ReadField(1, chat, Asc("*"))
    
    Dim copiar As Boolean

    copiar = True
    
    Dim duracion As Integer

    duracion = 250
    
    Dim Text As String
    Text = ReadField(2, chat, Asc("*"))
    
    Select Case QueEs
        Case "NPCDESC"
            
            chat = NpcData(Text).desc
            copiar = False
                        
            If npcs_en_render And tutorial_index <= 0 Then
                Dim icon As Long
                icon = HeadData(NpcData(Text).Head).Head(3).GrhIndex
                
                'Si icon es 0 quiere decir que no tiene cabeza, por ende renderizo body
                If icon = 0 Then
                    icon = GrhData(BodyData(NpcData(Text).Body).Walk(3).GrhIndex).Frames(1)
                    Call mostrarCartel(Split(NpcData(Text).Name, " <")(0), NpcData(Text).desc, icon, 200 + 30 * Len(chat), &H164B8A, , , True, 100, 629, 100, 585, 20, 650, 50, 80)
                Else
                    Call mostrarCartel(Split(NpcData(Text).Name, " <")(0), NpcData(Text).desc, icon, 200 + 30 * Len(chat), &H164B8A, , , True, 100, 629, 100, 685, -20, 589, 128, 128)
                End If
                
            End If
            
        Case "PMAG"
            chat = HechizoData(ReadField(2, chat, Asc("*"))).PalabrasMagicas
            If charlist(UserCharIndex).Muerto = True Then chat = ""
            copiar = False
            duracion = 20
            
        Case "QUESTFIN"
            chat = QuestList(ReadField(2, chat, Asc("*"))).DescFinal
            copiar = False
            duracion = 20
            
        Case "QUESTNEXT"
            chat = QuestList(ReadField(2, chat, Asc("*"))).NextQuest
            copiar = False
            duracion = 20
            
            If LenB(chat) = 0 Then
                chat = "Ya has completado esa misi�n para m�."

            End If
            
        Case "NOCONSOLA" ' El chat no sale en la consola
            chat = ReadField(2, chat, Asc("*"))
            copiar = False
            duracion = 20
        
    End Select
   'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(charindex).active = 1 Then
        Call Char_Dialog_Set(charindex, chat, colortexto, duracion, 30, 1, EsSpell)
    End If
    
    If charlist(charindex).esNpc = False Then
        If CopiarDialogoAConsola = 1 And copiar Then
    
            Call WriteChatOverHeadInConsole(charindex, chat, r, G, B)

        End If

    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChatOverHead", Erl)
    

End Sub

Private Sub HandleTextOverChar()

    On Error GoTo ErrHandler
    
    Dim chat      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    chat = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    
    Color = Reader.ReadInt32()
    
    Call SetCharacterDialogFx(charindex, chat, RGBA_From_vbColor(Color))

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleTextOverChar", Erl)
    

End Sub

Private Sub HandleTextOverTile()

    On Error GoTo ErrHandler
    
    Dim Text  As String

    Dim x     As Integer, y As Integer

    Dim Color As Long
    
    Text = Reader.ReadString8()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    Color = Reader.ReadInt32()
    
    Exit Sub
    
    If InMapBounds(x, y) Then
    
        With MapData(rrX(x), rrY(y))
            Dim Index As Integer
            
            If UBound(.DialogEffects) = 0 Then
                ReDim .DialogEffects(1 To 1)
                
                Index = 1
            Else

                For Index = 1 To UBound(.DialogEffects)

                    If .DialogEffects(Index).Text = vbNullString Then
                        Exit For

                    End If

                Next
                
                If Index > UBound(.DialogEffects) Then
                    ReDim .DialogEffects(1 To UBound(.DialogEffects) + 1)

                End If

            End If
            
            With .DialogEffects(Index)
            
                .Color = RGBA_From_vbColor(Color)
                .Start = FrameTime
                .Text = Text
                .offset.x = 0
                .offset.y = 0
            
            End With

        End With
        
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleTextOverTile", Erl)
    

End Sub

Private Sub HandleTextCharDrop()

    On Error GoTo ErrHandler
    
    Dim Text      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    Text = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    Color = Reader.ReadInt32()
    
    If charindex = 0 Then Exit Sub
    
    Exit Sub

    Dim x As Integer, y As Integer, OffsetX As Integer, OffsetY As Integer
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        OffsetX = .MoveOffsetX + .Body.HeadOffset.x
        OffsetY = .MoveOffsetY + .Body.HeadOffset.y

    End With
    
    If InMapBounds(x, y) Then
    
        With MapData(rrX(x), rrY(y))
            Dim Index As Integer
            
            If UBound(.DialogEffects) = 0 Then
                ReDim .DialogEffects(1 To 1)
                
                Index = 1
            Else

                For Index = 1 To UBound(.DialogEffects)

                    If .DialogEffects(Index).Text = vbNullString Then
                        Exit For

                    End If

                Next
                
                If Index > UBound(.DialogEffects) Then
                    ReDim .DialogEffects(1 To UBound(.DialogEffects) + 1)

                End If

            End If
            
            With .DialogEffects(Index)
            
                .Color = RGBA_From_vbColor(Color)
                .Start = FrameTime
                .Text = Text
                .offset.x = OffsetX
                .offset.y = OffsetY
            
            End With

        End With
        
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleTextCharDrop", Erl)
    

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim chat      As String
    Dim FontIndex As Integer
    Dim str       As String
    Dim r         As Byte
    Dim G         As Byte
    Dim B         As Byte
    Dim QueEs     As String
    Dim NpcName   As String
    Dim objname   As String
    Dim Hechizo   As Byte
    Dim UserName  As String
    Dim Valor     As String

    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()
    
    If ChatGlobal = 0 And FontIndex = FontTypeNames.FONTTYPE_GLOBAL Then Exit Sub

    QueEs = ReadField(1, chat, Asc("*"))

    Select Case QueEs

        Case "NPCNAME"
            NpcName = NpcData(ReadField(2, chat, Asc("*"))).Name
            chat = NpcName & ReadField(3, chat, Asc("*"))

        Case "O" 'OBJETO
            objname = ObjData(ReadField(2, chat, Asc("*"))).Name
            chat = objname & ReadField(3, chat, Asc("*"))

        Case "HECINF"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = "------------< Informaci�n del hechizo >------------" & vbCrLf & "Nombre: " & HechizoData(Hechizo).nombre & vbCrLf & "Descripci�n: " & HechizoData(Hechizo).desc & vbCrLf & "Skill requerido: " & HechizoData(Hechizo).MinSkill & " de magia." & vbCrLf & "Mana necesario: " & HechizoData(Hechizo).ManaRequerido & " puntos." & vbCrLf & "Stamina necesaria: " & HechizoData(Hechizo).StaRequerido & " puntos."

        Case "ProMSG"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = HechizoData(Hechizo).PropioMsg

        Case "HecMSG"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = HechizoData(Hechizo).HechizeroMsg & " la criatura."

        Case "HecMSGU"
            Hechizo = ReadField(2, chat, Asc("*"))
            UserName = ReadField(3, chat, Asc("*"))
            chat = HechizoData(Hechizo).HechizeroMsg & " " & UserName & "."
                
        Case "HecMSGA"
            Hechizo = ReadField(2, chat, Asc("*"))
            UserName = ReadField(3, chat, Asc("*"))
            chat = UserName & " " & HechizoData(Hechizo).TargetMsg
                
        Case "EXP"
            Valor = ReadField(2, chat, Asc("*"))
            'chat = "Has ganado " & valor & " puntos de experiencia."
        
        Case "ID"

            Dim id    As Integer
            Dim extra As String

            id = ReadField(2, chat, Asc("*"))
            extra = ReadField(3, chat, Asc("*"))
                
            chat = Locale_Parse_ServerMessage(id, extra)
           
    End Select
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            G = 255
        Else
            G = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            B = 255
        Else
            B = Val(str)

        End If
            
        Call AddToConsole(Left$(chat, InStr(1, chat, "~") - 1), r, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    
    Else

        With FontTypes(FontIndex)
            Call AddToConsole(chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleConsoleMessage", Erl)
    

End Sub

Private Sub HandleLocaleMsg()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim chat      As String

    Dim FontIndex As Integer

    Dim str       As String

    Dim r         As Byte

    Dim G         As Byte

    Dim B         As Byte

    Dim QueEs     As String

    Dim NpcName   As String

    Dim objname   As String

    Dim Hechizo   As Byte

    Dim UserName  As String

    Dim Valor     As String

    Dim id        As Integer

    id = Reader.ReadInt16()
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()

    chat = Locale_Parse_ServerMessage(id, chat)
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            G = 255
        Else
            G = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            B = 255
        Else
            B = Val(str)

        End If
            
        Call AddToConsole(Left$(chat, InStr(1, chat, "~") - 1), r, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontIndex)
            Call AddToConsole(chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleLocaleMsg", Erl)
    
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08 (NicoNZ)
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim chat As String

    Dim status As Byte
    
    Dim str  As String

    Dim r    As Byte

    Dim G    As Byte

    Dim B    As Byte

    Dim tmp  As Integer

    Dim Cont As Integer
    
    status = Reader.ReadInt8()
    chat = Reader.ReadString8()
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        Call AddToConsole(chat, .red, .green, .blue, .bold, .italic, , , status)
    End With


    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleGuildChat", Erl)
    

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim mensaje As String

    mensaje = Reader.ReadString8()

    Select Case QueRender

        Case 0
            frmMensaje.msg.Caption = mensaje
            frmMensaje.Show , frmMain

        Case 1
            Call Sound.Sound_Play(SND_EXCLAMACION)
            Call TextoAlAsistente(mensaje)
            Call Long_2_RGBAList(textcolorAsistente, -1)

        Case 2
            frmMensaje.Show
            frmMensaje.msg.Caption = mensaje
        
        Case 3
            frmMensaje.Show , frmConnect
            frmMensaje.msg.Caption = mensaje

    End Select
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowMessageBox", Erl)
    

End Sub


''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
    
    On Error GoTo HandleUserIndexInServer_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    userIndex = Reader.ReadInt16()
    
    Exit Sub

HandleUserIndexInServer_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserIndexInServer", Erl)
    
    
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
    
    On Error GoTo HandleUserCharIndexInServer_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserCharIndex = Reader.ReadInt16()
    'Debug.Print "UserCharIndex " & UserCharIndex
    UserPos = charlist(UserCharIndex).Pos
    
    
    Call RefreshMap
    
    'MapData(rX(UserPos.x), rY(UserPos.y)).charindex = UserCharIndex
    
    
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
    
    lastMove = FrameTime
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If
    
    Call checkZona
    If frmMapaGrande.visible Then
        Call frmMapaGrande.ActualizarPosicionMapa
    End If
    
    Exit Sub

HandleUserCharIndexInServer_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserCharIndexInServer", Erl)
    
    
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim charindex     As Integer

    Dim Body          As Integer

    Dim Head          As Integer

    Dim Heading       As E_Heading

    Dim x As Integer

    Dim y As Integer

    Dim weapon        As Integer

    Dim shield        As Integer

    Dim helmet        As Integer

    Dim privs         As Integer

    Dim AuraParticula As Byte

    Dim ParticulaFx   As Byte

    Dim appear        As Byte

    Dim group_index   As Integer
    
    charindex = Reader.ReadInt16()

    Body = Reader.ReadInt16()
    Head = Reader.ReadInt16()
    Heading = Reader.ReadInt8()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    weapon = Reader.ReadInt16()
    shield = Reader.ReadInt16()
    helmet = Reader.ReadInt16()
    
    With charlist(charindex)
        'Call SetCharacterFx(charindex, Reader.ReadInt16(), Reader.ReadInt16())
        .FxIndex = Reader.ReadInt16
        
        Reader.ReadInt16 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)

        End If
        
        Dim NombreYClan As String
        NombreYClan = Reader.ReadString8()
     
   '
    
         
        Dim Pos As Integer
        Pos = InStr(NombreYClan, "<")

        If Pos = 0 Then Pos = InStr(NombreYClan, "[")
        If Pos = 0 Then Pos = Len(NombreYClan) + 2
        
        .nombre = Left$(NombreYClan, Pos - 2)
        .clan = mid$(NombreYClan, Pos)
        
        .status = Reader.ReadInt8()
        
        privs = Reader.ReadInt8()
        ParticulaFx = Reader.ReadInt8()
        .Head_Aura = Reader.ReadString8()
        .Arma_Aura = Reader.ReadString8()
        .Body_Aura = Reader.ReadString8()
        .DM_Aura = Reader.ReadString8()
        .RM_Aura = Reader.ReadString8()
        .Otra_Aura = Reader.ReadString8()
        .Escudo_Aura = Reader.ReadString8()
        .Speeding = Reader.ReadReal32()
        .Invisible = False
        
        Dim FlagNpc As Byte
        FlagNpc = Reader.ReadInt8()
        
        .esNpc = FlagNpc > 0
        .EsMascota = FlagNpc = 2
        
        .appear = Reader.ReadInt8()
        appear = .appear
        .group_index = Reader.ReadInt16()
        .clan_index = Reader.ReadInt16()
        .clan_nivel = Reader.ReadInt8()
        .UserMinHp = Reader.ReadInt32()
        .UserMaxHp = Reader.ReadInt32()
        .UserMinMAN = Reader.ReadInt32()
        .UserMaxMAN = Reader.ReadInt32()
        .simbolo = Reader.ReadInt8()
         Dim flags As Byte
        
        flags = Reader.ReadInt8()
        
                
        .Idle = flags And &O1
        
        .Navegando = flags And &O2
        .tipoUsuario = Reader.ReadInt8()
        .teamCaptura = Reader.ReadInt8()
        .banderaIndex = Reader.ReadInt8()
        .AnimAtaque1 = Reader.ReadInt16()
        
        
        If (.Pos.x <> 0 And .Pos.y <> 0) Then
            If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then
                'Erase the old character from map
                MapData(rrX(charlist(charindex).Pos.x), rrY(charlist(charindex).Pos.y)).charindex = 0

            End If

        End If

        If privs <> 0 Then
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0

        End If

        .Muerto = (Body = CASPER_BODY_IDLE)
        '.AlphaPJ = 255
    
        Call MakeChar(charindex, Body, Head, Heading, x, y, weapon, shield, helmet, ParticulaFx, appear)
        'Debug.Print "name: " & charlist(charindex).nombre; "|charindex: " & charindex, x, y
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).started = FrameTime

        End If
        
    End With
    
    Call RefreshAllChars
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharacterCreate", Erl)
    

End Sub


Private Sub HandleUpdateFlag()

    On Error GoTo ErrHandler

    Dim charindex As Integer
    Dim flag As Long
    
    
    charindex = Reader.ReadInt16()
    flag = Reader.ReadInt8()
    
    With charlist(charindex)
        .banderaIndex = flag
    End With
    

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleTextOverChar", Erl)
    

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
    
    On Error GoTo HandleCharacterRemove_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex   As Integer
    Dim dbgid As Integer
    
    Dim Desvanecido As Boolean
    Dim fueWarp As Boolean
    charindex = Reader.ReadInt16()
    Desvanecido = Reader.ReadBool()
    fueWarp = Reader.ReadBool()
    If Desvanecido And charlist(charindex).esNpc = True Then
        Call CrearFantasma(charindex)
    End If

    Call EraseChar(charindex)
    Call RefreshAllChars
    
    Exit Sub

HandleCharacterRemove_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharacterRemove", Erl)
    
    
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
    
    On Error GoTo HandleCharacterMove_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer
    Dim x As Integer
    Dim y As Integer
    Dim dir       As Byte
    charindex = Reader.ReadInt16()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    With charlist(charindex)
        
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(charindex)

        End If

    End With
    
    Call Char_Move_by_Pos(charindex, x, y)
    
    Call RefreshAllChars
    
    Exit Sub

HandleCharacterMove_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharacterMove", Erl)
    
    
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    On Error GoTo HandleForceCharMove_Err
    
    Dim direccion As Byte
    direccion = Reader.ReadInt8()
    
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)

    Call Char_Move_by_Head(UserCharIndex, direccion)
    Call MoveScreen(direccion)
        
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If

    If frmMapaGrande.visible Then
        Call frmMapaGrande.ActualizarPosicionMapa
    End If
    
    Call RefreshAllChars
    
    Exit Sub

HandleForceCharMove_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleForceCharMove", Erl)
    
    
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMoveSiguiendo()
    
    On Error GoTo HandleForceCharMoveSiguiendo_Err
    
    Dim direccion As Byte
    direccion = Reader.ReadInt8()
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)
    'Capaz hay que eliminar el char_move_by_head
    
    UserPos.x = charlist(CharindexSeguido).Pos.x
    UserPos.y = charlist(CharindexSeguido).Pos.y
    checkZona
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If
    
    Call Char_Move_by_Head(CharindexSeguido, direccion)
    Call MoveScreen(direccion)
    'Call RefreshAllChars
    
    Exit Sub

HandleForceCharMoveSiguiendo_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleForceCharMoveSiguiendo", Erl)
    
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim charindex As Integer

    Dim TempInt   As Integer

    Dim headIndex As Integer

    charindex = Reader.ReadInt16()
    
    With charlist(charindex)
        TempInt = Reader.ReadInt16()

        If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
            .Body = BodyData(0)
        Else
            .Body = BodyData(TempInt)
            .iBody = TempInt

        End If
        
        headIndex = Reader.ReadInt16()

        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .IHead = 0
            
        Else
            .Head = HeadData(headIndex)
            .IHead = headIndex

        End If

        .Muerto = (.iBody = CASPER_BODY_IDLE)
        
        .Heading = Reader.ReadInt8()
        
        TempInt = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(WeaponAnimData) Then
            .Arma = WeaponAnimData(TempInt)
        End If

        TempInt = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(ShieldAnimData) Then
            .Escudo = ShieldAnimData(TempInt)
        End If
        
        TempInt = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(CascoAnimData) Then
            .Casco = CascoAnimData(TempInt)
        End If
                
        If .Body.HeadOffset.y = -26 Then
            .EsEnano = True
        Else
            .EsEnano = False

        End If
        
        'Call SetCharacterFx(charindex, Reader.ReadInt16(), Reader.ReadInt16())
        .FxIndex = Reader.ReadInt16
        
        Reader.ReadInt16 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)

        End If
        
        Dim flags As Byte
        
        flags = Reader.ReadInt8()
        
        .Idle = flags And &O1
        
        .Navegando = flags And &O2
        
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).started = FrameTime

        End If

    End With
    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharacterChange", Erl)
    
    
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
    
    On Error GoTo HandleObjectCreate_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Integer

    Dim y As Integer

    Dim ObjIndex As Integer
    
    Dim Amount   As Integer

    Dim Color    As RGBA

    Dim Rango    As Byte

    Dim id       As Long
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    ObjIndex = Reader.ReadInt16()
    
    Amount = Reader.ReadInt16
    
    MapData(rrX(x), rrY(y)).ObjGrh.GrhIndex = ObjData(ObjIndex).GrhIndex
    
    MapData(rrX(x), rrY(y)).OBJInfo.ObjIndex = ObjIndex
    
    MapData(rrX(x), rrY(y)).OBJInfo.Amount = Amount
    
    Call InitGrh(MapData(rrX(x), rrY(y)).ObjGrh, MapData(rrX(x), rrY(y)).ObjGrh.GrhIndex)
    
    If ObjData(ObjIndex).CreaLuz <> "" Then
        Call Long_2_RGBA(Color, Val(ReadField(2, ObjData(ObjIndex).CreaLuz, Asc(":"))))
        Rango = Val(ReadField(1, ObjData(ObjIndex).CreaLuz, Asc(":")))
        MapData(rrX(x), rrY(y)).luz.Color = Color
        MapData(rrX(x), rrY(y)).luz.Rango = Rango
        
        If Rango < 100 Then
            id = x & y
            LucesCuadradas.Light_Create x, y, Color, Rango, id
        Else
            LucesRedondas.Create_Light_To_Map x, y, Color, Rango - 99
        End If
        
    End If
        
    If ObjData(ObjIndex).CreaParticulaPiso <> 0 Then
        MapData(rrX(x), rrY(y)).particle_group = 0
        General_Particle_Create ObjData(ObjIndex).CreaParticulaPiso, x, y, -1

    End If
    
    Exit Sub

HandleObjectCreate_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleObjectCreate", Erl)
    
    
End Sub

Private Sub HandleFxPiso()
    
    On Error GoTo HandleFxPiso_Err

    '***************************************************
    'Ladder
    '30/5/10
    '***************************************************
    
    Dim x As Integer

    Dim y As Integer

    Dim fX As Byte

    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    fX = Reader.ReadInt16()
    
    Call SetMapFx(x, y, fX, 0)
    
    Exit Sub

HandleFxPiso_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleFxPiso", Erl)
    
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
    
    On Error GoTo HandleObjectDelete_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Integer

    Dim y As Integer

    Dim id As Long
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If ObjData(MapData(rrX(x), rrY(y)).OBJInfo.ObjIndex).CreaLuz <> "" Then
        id = LucesCuadradas.Light_Find(x & y)
        LucesCuadradas.Light_Remove id
        MapData(rrX(x), rrY(y)).luz.Color = COLOR_EMPTY
        MapData(rrX(x), rrY(y)).luz.Rango = 0
       ' LucesCuadradas.Light_Render_All

    End If
    
    MapData(rrX(x), rrY(y)).ObjGrh.GrhIndex = 0
    MapData(rrX(x), rrY(y)).OBJInfo.ObjIndex = 0
    
    If ObjData(MapData(rrX(x), rrY(y)).OBJInfo.ObjIndex).CreaParticulaPiso <> 0 Then
        Graficos_Particulas.Particle_Group_Remove (MapData(rrX(x), rrY(y)).particle_group)

    End If
    
    Exit Sub

HandleObjectDelete_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleObjectDelete", Erl)
    
    
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
    
    On Error GoTo HandleBlockPosition_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Integer, y As Integer, B As Byte
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    B = Reader.ReadInt8()

    MapData(rrX(x), rrY(y)).Blocked = MapData(rrX(x), rrY(y)).Blocked And Not eBlock.ALL_SIDES
    MapData(rrX(x), rrY(y)).Blocked = MapData(rrX(x), rrY(y)).Blocked Or B
    
    Exit Sub

HandleBlockPosition_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlockPosition", Erl)
    
    
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
    
    On Error GoTo HandlePlayMIDI_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call Reader.ReadInt8   ' File
    Call Reader.ReadInt16  ' Loop
    
    Exit Sub

HandlePlayMIDI_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePlayMIDI", Erl)
    
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
    
    On Error GoTo HandlePlayWave_Err

    '***************************************************
    'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
        
    Dim wave As Integer
    Dim srcX As Integer
    Dim srcY As Integer
    Dim cancelLastWave As Byte
    
    wave = Reader.ReadInt16()
    srcX = Reader.ReadInt16()
    srcY = Reader.ReadInt16()
    cancelLastWave = Reader.ReadInt8()
    
    If wave = 400 And MapDat.Niebla = 0 Then Exit Sub
    If wave = 401 And MapDat.Niebla = 0 Then Exit Sub
    If wave = 402 And MapDat.Niebla = 0 Then Exit Sub
    If wave = 403 And MapDat.Niebla = 0 Then Exit Sub
    If wave = 404 And MapDat.Niebla = 0 Then Exit Sub
    
    If cancelLastWave Then
        Call Sound.Sound_Stop(CStr(wave))
        If cancelLastWave = 2 Then Exit Sub
    End If
    
    If srcX = 0 Or srcY = 0 Then
        Call Sound.Sound_Play(CStr(wave), False, 0, 0)
    Else

        If Not EstaEnArea(srcX, srcY) Then
        Else
            Call Sound.Sound_Play(CStr(wave), False, Sound.Calculate_Volume(srcX, srcY), Sound.Calculate_Pan(srcX, srcY))
        End If

    End If
    
    ' Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
    
    Exit Sub

HandlePlayWave_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePlayWave", Erl)
    
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWaveStep()
    
    On Error GoTo HandlePlayWaveStep_Err
        
    Dim grh As Long
    Dim distance As Byte
    Dim balance As Integer
    Dim step As Boolean
    
    grh = Reader.ReadInt32()
    distance = Reader.ReadInt8()
    balance = Reader.ReadInt16()
    step = Reader.ReadBool()
    
    Call DoPasosInvi(grh, distance, balance, step)
    
    
    Exit Sub

HandlePlayWaveStep_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePlayWaveStep", Erl)
    
    
End Sub

Private Sub HandlePosLLamadaDeClan()
    
    On Error GoTo HandlePosLLamadaDeClan_Err

    '***************************************************
    'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
        
    Dim map  As Integer

    Dim srcX As Integer

    Dim srcY As Integer
    
    map = Reader.ReadInt16()
    srcX = Reader.ReadInt16()
    srcY = Reader.ReadInt16()

    Dim idmap As Integer

    
    LLamadaDeclanMapa = map
    idmap = ObtenerIdMapaDeLlamadaDeClan(map)

    Dim x As Long

    Dim y As Long
    
    x = (idmap - 1) Mod 14
    y = Int((idmap - 1) / 14)

    'frmMapaGrande.lblAllies.Top = Y * 32
    'frmMapaGrande.lblAllies.Left = X * 32

    frmMapaGrande.llamadadeclan.Top = y * 32 + (srcX / 4.5)
    frmMapaGrande.llamadadeclan.Left = x * 32 + (srcY / 4.5)

    frmMapaGrande.llamadadeclan.visible = True

    frmMain.LlamaDeclan.enabled = True

    frmMapaGrande.Shape2.visible = True

    frmMapaGrande.Shape2.Top = y * 32
    frmMapaGrande.Shape2.Left = x * 32

    LLamadaDeclanx = srcX
    LLamadaDeclany = srcY

    HayLLamadaDeclan = True
    
    ' Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
    
    Exit Sub

HandlePosLLamadaDeClan_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePosLLamadaDeClan", Erl)
    
    
End Sub

Private Sub HandleCharUpdateHP()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim charindex As Integer

    Dim minhp     As Long

    Dim maxhp     As Long
    
    charindex = Reader.ReadInt16()
    minhp = Reader.ReadInt32()
    maxhp = Reader.ReadInt32()

    charlist(charindex).UserMinHp = minhp
    charlist(charindex).UserMaxHp = maxhp
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharUpdateHP", Erl)
    
    
End Sub

Private Sub HandleCharUpdateMAN()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim charindex As Integer

    Dim minman     As Long

    Dim maxman     As Long
    
    charindex = Reader.ReadInt16()
    minman = Reader.ReadInt32()
    maxman = Reader.ReadInt32()

    charlist(charindex).UserMinMAN = minman
    charlist(charindex).UserMaxMAN = maxman
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharUpdateMAN", Erl)
    
    
End Sub

Private Sub HandleArmaMov()
    
    On Error GoTo HandleArmaMov_Err

    '***************************************************

    Dim charindex As Integer

   charindex = Reader.ReadInt16()

    With charlist(charindex)

        If Not .Moving Then
            .MovArmaEscudo = True
            .Arma.WeaponWalk(.Heading).started = FrameTime
            .Arma.WeaponWalk(.Heading).Loops = 0

        End If

    End With
    
    Exit Sub

HandleArmaMov_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleArmaMov", Erl)
    
    
End Sub

Private Sub HandleEscudoMov()
    
    On Error GoTo HandleEscudoMov_Err

    '***************************************************

    Dim charindex As Integer

    charindex = Reader.ReadInt16()

    With charlist(charindex)

        If Not .Moving Then
            .MovArmaEscudo = True
            .Escudo.ShieldWalk(.Heading).started = FrameTime
            .Escudo.ShieldWalk(.Heading).Loops = 0

        End If

    End With
    
    Exit Sub

HandleEscudoMov_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleEscudoMov", Erl)
    
    
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    'Clear guild's list
    frmGuildAdm.guildslist.Clear
    
    Dim guildsStr As String
    guildsStr = Reader.ReadString8()
    
    If Len(guildsStr) > 0 Then

        Dim guilds() As String
        guilds = Split(guildsStr, SEPARATOR)
        
        ReDim ClanesList(0 To UBound(guilds())) As Tclan
        
        ListaClanes = True
        
        Dim i As Long

        For i = 0 To UBound(guilds())
            ClanesList(i).nombre = ReadField(1, guilds(i), Asc("-"))
            ClanesList(i).Alineacion = Val(ReadField(2, guilds(i), Asc("-")))
            ClanesList(i).indice = i
        Next i
        
        For i = 0 To UBound(guilds())
            'If ClanesList(i).Alineacion = 0 Then
            Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
            'End If
        Next i

    End If
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    Call Establecer_Borde(frmGuildAdm.guildslist, frmGuildAdm, COLOR_AZUL, 0, 0)

    Call frmGuildAdm.Show(vbModeless, frmMain)
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleGuildList", Erl)
    

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
    
    On Error GoTo HandleAreaChanged_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Integer

    Dim y As Integer
    
    Dim Heading As Byte
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    Heading = Reader.ReadInt8()
        
    Call CambioDeArea(x, y, Heading)
    
    Exit Sub

HandleAreaChanged_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAreaChanged", Erl)
    
    
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandlePauseToggle_Err
    
    pausa = Not pausa
    
    Exit Sub

HandlePauseToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePauseToggle", Erl)
    
    
End Sub

Private Sub HandleRainToggle()
    '**
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '**
    'Remove packet ID

    On Error GoTo HandleRainToggle_Err


    bRain = Reader.ReadBool

    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub


    If Not bRain Then
        If MapDat.Lluvia Then

            If bTecho Then
                Call Sound.Sound_Play(192)
            Else
                Call Sound.Sound_Play(195)

            End If

            Call Sound.Ambient_Stop

            Call Graficos_Particulas.Engine_MeteoParticle_Set(-1)

        End If

    Else

        If MapDat.Lluvia Then

            Call Graficos_Particulas.Engine_MeteoParticle_Set(Particula_Lluvia)

        End If

        ' Call Audio.StopWave(AmbientalesBufferIndex)
    End If


    Exit Sub

HandleRainToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleRainToggle", Erl)


End Sub


''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
    
    On Error GoTo HandleCreateFX_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer

    Dim fX        As Integer

    Dim Loops     As Integer
    
    Dim x As Integer, y As Integer
        
    charindex = Reader.ReadInt16()
    fX = Reader.ReadInt16()
    Loops = Reader.ReadInt16()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(rrX(x), rrY(y)).charindex = charindex
            End If
        End With
    End If
    
    
    If fX = 0 Then
        charlist(charindex).fX.AnimacionContador = 29
        Exit Sub

    End If
    
    Call SetCharacterFx(charindex, fX, Loops)
    
    Exit Sub

HandleCreateFX_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCreateFX", Erl)
    
    
End Sub

''
' Handles the CharAtaca message.

Private Sub HandleCharAtaca()
    
    On Error GoTo HandleCharAtaca_Err
    
    Dim NpcIndex As Integer
    Dim VictimIndex As Integer
    Dim danio     As Long
    Dim AnimAttack As Integer
    
    NpcIndex = Reader.ReadInt16()
    VictimIndex = Reader.ReadInt16()
    danio = Reader.ReadInt32()
    AnimAttack = Reader.ReadInt16()
    
    Dim grh As grh
            
    If AnimAttack > 0 Then
        charlist(NpcIndex).Body = BodyData(AnimAttack)
        charlist(NpcIndex).Body.Walk(charlist(NpcIndex).Heading).started = FrameTime
    End If
    
    'renderizo sangre si est� sin montar ni navegar
    If danio > 0 And charlist(VictimIndex).Navegando = 0 Then Call SetCharacterFx(VictimIndex, 14, 0)
        
    
    If charlist(UserCharIndex).Muerto = False Then
        Call Sound.Sound_Play(CStr(IIf(danio = -1, 2, 10)), False, Sound.Calculate_Volume(charlist(NpcIndex).Pos.x, charlist(NpcIndex).Pos.y), Sound.Calculate_Pan(charlist(NpcIndex).Pos.x, charlist(NpcIndex).Pos.y))
    End If
        
    Exit Sub
    

HandleCharAtaca_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharAtaca", Erl)
    
    
End Sub

''
' Handles the CharAtaca message.

Private Sub HandleNotificarClienteSeguido()
    
    On Error GoTo NotificarClienteSeguido_Err
    
    Seguido = Reader.ReadInt8
    
    Exit Sub
    

NotificarClienteSeguido_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.NotificarClienteSeguido", Erl)
    
    
End Sub
''
' Handles the UpdateUserStats message.
''
' Handles the CharAtaca message.

Private Sub HandleRecievePosSeguimiento()
    
    On Error GoTo RecievePosSeguimiento_Err
    
    Dim PosX As Integer
    Dim PosY As Integer
    
    PosX = Reader.ReadInt16()
    PosY = Reader.ReadInt16()
    
    frmMain.shapexy.Left = PosX - 6
    frmMain.shapexy.Top = PosY - 6
    Exit Sub
    

RecievePosSeguimiento_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.RecievePosSeguimiento", Erl)
    
    
End Sub

Private Sub HandleCancelarSeguimiento()
    
    On Error GoTo CancelarSeguimiento_Err
    
    frmMain.shapexy.Left = 1200
    frmMain.shapexy.Top = 1200
    CharindexSeguido = 0
    OffsetLimitScreen = 32
    Exit Sub
    

CancelarSeguimiento_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.CancelarSeguimiento", Erl)
    
End Sub

Private Sub HandleGetInventarioHechizos()
    
    On Error GoTo GetInventarioHechizos_Err
    
    Dim inventario_o_hechizos As Byte
    Dim hechiSel As Byte
    Dim scrollSel As Byte
    
    inventario_o_hechizos = Reader.ReadInt8()
    hechiSel = Reader.ReadInt8()
    scrollSel = Reader.ReadInt8()
    'Clicke� en inventario
    If inventario_o_hechizos = 1 Then
        Call frmMain.inventoryClick
    'Clicke� en hechizos
    ElseIf inventario_o_hechizos = 2 Then
        Call frmMain.hechizosClick
        hlst.Scroll = scrollSel
        hlst.ListIndex = hechiSel
        
    End If
    Exit Sub
    

GetInventarioHechizos_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.GetInventarioHechizos", Erl)
    
End Sub


Private Sub HandleNotificarClienteCasteo()
    
    On Error GoTo NotificarClienteCasteo_Err
    
    Dim Value As Byte
    
    Value = Reader.ReadInt8()
    'Clicke� en inventario
    If Value = 1 Then
        frmMain.shapexy.BackColor = RGB(0, 170, 0)
    'Clicke� en hechizos
    Else
        frmMain.shapexy.BackColor = RGB(170, 0, 0)
    End If
    Exit Sub
    

NotificarClienteCasteo_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.NotificarClienteCasteo", Erl)
    
End Sub



Private Sub HandleSendFollowingCharindex()
    
    On Error GoTo SendFollowingCharindex_Err
    
    Dim charindex As Integer
    charindex = Reader.ReadInt16()
    UserCharIndex = charindex
    CharindexSeguido = charindex
    OffsetLimitScreen = 31
    Exit Sub
    

SendFollowingCharindex_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.SendFollowingCharindex", Erl)
    
End Sub
''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
    
    On Error GoTo HandleUpdateUserStats_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserMaxHp = Reader.ReadInt16()
    UserMinHp = Reader.ReadInt16()
    UserMaxMAN = Reader.ReadInt16()
    UserMinMAN = Reader.ReadInt16()
    UserMaxSTA = Reader.ReadInt16()
    UserMinSTA = Reader.ReadInt16()
    UserGLD = Reader.ReadInt32()
    OroPorNivel = Reader.ReadInt32()
    UserLvl = Reader.ReadInt8()
    UserPasarNivel = Reader.ReadInt32()
    UserExp = Reader.ReadInt32()
    UserClase = Reader.ReadInt8()
    
    If UserPasarNivel > 0 Then
        frmMain.lblPorcLvl.Caption = Round(UserExp * (100 / UserPasarNivel), 2) & "%"
        frmMain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
        frmMain.EXPBAR.Width = UserExp / UserPasarNivel * 235
    Else
        frmMain.EXPBAR.Width = 235
        frmMain.lblPorcLvl.Caption = "�Nivel m�ximo!" 'nivel maximo
        frmMain.exp.Caption = "�Nivel m�ximo!"

    End If
    
    If UserMaxHp > 0 Then
        frmMain.Hpshp.Width = UserMinHp / UserMaxHp * 243
    Else
        frmMain.Hpshp.Width = 0
    End If

    frmMain.HpBar.Caption = UserMinHp & " / " & UserMaxHp

    If QuePesta�aInferior = 0 Then
        frmMain.Hpshp.visible = (UserMinHp > 0)

    End If

    If UserMaxMAN > 0 Then
        frmMain.MANShp.Width = UserMinMAN / UserMaxMAN * 243
        frmMain.manabar.Caption = UserMinMAN & " / " & UserMaxMAN

        If QuePesta�aInferior = 0 Then
            frmMain.MANShp.visible = (UserMinMAN > 0)
            frmMain.manabar.visible = True

        End If

    Else
        frmMain.manabar.visible = False
        frmMain.MANShp.Width = 0
        frmMain.MANShp.visible = False

    End If
    
    If UserMaxSTA > 0 Then
        frmMain.STAShp.Width = UserMinSTA / UserMaxSTA * 97
    Else
        frmMain.STAShp.Width = 0
    End If

    frmMain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA
    
    If QuePesta�aInferior = 0 Then
        frmMain.STAShp.visible = (UserMinSTA > 0)

    End If
    
    'If UserGLD > UserLvl * OroPorNivel Then
    If UserGLD <= 100000 Then
        frmMain.GldLbl.ForeColor = vbRed
    Else
        frmMain.GldLbl.ForeColor = &H80FFFF
    End If

    frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
    frmMain.lblLvl.Caption = UserLvl
    frmMain.lblClase.Caption = ListaClases(UserClase)
    
    If UserMinHp = 0 Then
        UserEstado = 1
        charlist(UserCharIndex).Invisible = False
        DrogaCounter = 0
    Else
        UserEstado = 0

    End If
    
    Exit Sub

HandleUpdateUserStats_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateUserStats", Erl)
    
    
End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
    
    On Error GoTo HandleWorkRequestTarget_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Dim UsingSkillREcibido As Byte
    
    UsingSkillREcibido = Reader.ReadInt8()
    casteaArea = Reader.ReadBool()
    RadioHechizoArea = Reader.ReadInt8()
    'RadioHechizoArea = RadioHechizoArea / 2
    
    If EstaSiguiendo Then Exit Sub
    If UsingSkillREcibido = 0 Then
        frmMain.MousePointer = 0
        Call FormParser.Parse_Form(frmMain, E_NORMAL)
        UsingSkill = UsingSkillREcibido
        Exit Sub

    End If

    If UsingSkillREcibido = UsingSkill Then Exit Sub
   
    UsingSkill = UsingSkillREcibido
    frmMain.MousePointer = 2
    Select Case UsingSkill

        Case magia
            Call FormParser.Parse_Form(frmMain, E_CAST)
            
            Call AddToConsole(MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
            

        Case Robar
            Call AddToConsole(MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case Herreria
            Call AddToConsole(MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case Proyectiles
            Call AddToConsole(MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_ARROW)

        Case eSkill.Talar, eSkill.Alquimia, eSkill.Carpinteria, eSkill.Herreria, eSkill.Mineria, eSkill.Pescar
            Call AddToConsole("Has click donde deseas trabajar...", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case Grupo
            Call AddToConsole(MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case MarcaDeClan
            Call AddToConsole("Seleccione el personaje que desea marcar..", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case MarcaDeGM
            Call AddToConsole("Seleccione el personaje que desea marcar..", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

    End Select
    
    Exit Sub

HandleWorkRequestTarget_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleWorkRequestTarget", Erl)
    
    
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot        As Byte
    Dim ObjIndex    As Integer
    Dim Name        As String
    Dim Amount      As Integer
    Dim Equipped    As Boolean
    Dim GrhIndex    As Long
    Dim ObjType     As Byte
    Dim MaxHit      As Integer
    Dim MinHit      As Integer
    Dim MaxDef      As Integer
    Dim MinDef      As Integer
    Dim Value       As Single
    Dim podrausarlo As Byte

    Slot = Reader.ReadInt8()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16()
    Equipped = Reader.ReadBool()
    Value = Reader.ReadReal32()
    podrausarlo = Reader.ReadInt8()

    Name = ObjData(ObjIndex).Name
    GrhIndex = ObjData(ObjIndex).GrhIndex
    ObjType = ObjData(ObjIndex).ObjType
    MaxHit = ObjData(ObjIndex).MaxHit
    MinHit = ObjData(ObjIndex).MinHit
    MaxDef = ObjData(ObjIndex).MaxDef
    MinDef = ObjData(ObjIndex).MinDef

    If Equipped Then

        Select Case ObjType

            Case eObjType.otWeapon
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot

            Case eObjType.otNudillos
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot

            Case eObjType.otArmadura
                frmMain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = Slot

            Case eObjType.otESCUDO
                frmMain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = Slot

            Case eObjType.otCASCO
                frmMain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = Slot

        End Select
        
    Else

        Select Case Slot

            Case UserWeaponEqpSlot
                frmMain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0

            Case UserArmourEqpSlot
                frmMain.lblArmor = "0/0"
                UserArmourEqpSlot = 0

            Case UserHelmEqpSlot
                frmMain.lblShielder = "0/0"
                UserHelmEqpSlot = 0

            Case UserShieldEqpSlot
                frmMain.lblHelm = "0/0"
                UserShieldEqpSlot = 0

        End Select

    End If

    Call frmMain.Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    If frmComerciar.visible Then
        Call frmComerciar.InvComUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    ElseIf frmBancoObj.visible Then
        Call frmBancoObj.InvBankUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
        
    ElseIf frmBancoCuenta.visible Then
        Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    
    ElseIf frmCrafteo.visible Then
        Call frmCrafteo.InvCraftUser.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChangeInventorySlot", Erl)
    

End Sub

' Handles the InventoryUnlockSlots message.
Private Sub HandleInventoryUnlockSlots()
    '***************************************************
    'Author: Ruthnar
    'Last Modification: 30/09/20
    '
    '***************************************************
    
    On Error GoTo HandleInventoryUnlockSlots_Err
    
    Dim i As Integer
    
    UserInvUnlocked = Reader.ReadInt8
    
    For i = 1 To UserInvUnlocked
    
        frmMain.imgInvLock(i - 1).Picture = LoadInterface("inventoryunlocked.bmp")
    
    Next i

    Exit Sub

HandleInventoryUnlockSlots_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleInventoryUnlockSlots", Erl)
    
    
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot As Byte
    Dim BankSlot As Inventory
    
    With BankSlot
        Slot = Reader.ReadInt8()
        .ObjIndex = Reader.ReadInt16()
        .Amount = Reader.ReadInt16()
        .Valor = Reader.ReadInt32()
        .PuedeUsar = Reader.ReadInt8()
        
        If .ObjIndex > 0 Then
            .Name = ObjData(.ObjIndex).Name
            .GrhIndex = ObjData(.ObjIndex).GrhIndex
            .ObjType = ObjData(.ObjIndex).ObjType
            .MaxHit = ObjData(.ObjIndex).MaxHit
            .MinHit = ObjData(.ObjIndex).MinHit
            .Def = ObjData(.ObjIndex).MaxDef
        End If
        
        Call frmBancoObj.InvBoveda.SetItem(Slot, .ObjIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)

    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChangeBankSlot", Erl)
    

End Sub

''
' Handles the ChangeSpellSlot message

Private Sub HandleChangeSpellSlot()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot     As Byte

    Dim Index    As Byte

    Dim cooldown As Integer

    Slot = Reader.ReadInt8()
    
    UserHechizos(Slot) = Reader.ReadInt16()
    Index = Reader.ReadInt8()

    If Index < 254 Then
    
        If Slot <= hlst.ListCount Then
            hlst.List(Slot - 1) = HechizoData(Index).nombre
        Else
            Call hlst.AddItem(HechizoData(Index).nombre)
            hlst.Scroll = LastScroll
        End If

    Else
    
        If Slot <= hlst.ListCount Then
            hlst.List(Slot - 1) = "(Vacio)"
        Else
            Call hlst.AddItem("(Vacio)")
            hlst.Scroll = LastScroll
        End If
    
    End If
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChangeSpellSlot", Erl)
    

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
    
    On Error GoTo HandleAtributes_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Reader.ReadInt8()
    Next i
    
    'Show them in character creation

    
    If LlegaronStats Then
        frmStatistics.Iniciar_Labels
        frmStatistics.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
        frmStatistics.Show , frmMain
    Else
        LlegaronAtrib = True
    End If
    

    
    Exit Sub

HandleAtributes_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAtributes", Erl)
    
    
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = Reader.ReadInt16()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To count
        ArmasHerrero(i).Index = Reader.ReadInt16()
        ' tmp = ObjData(ArmasHerrero(i).Index).name        'Get the object's name
        ArmasHerrero(i).LHierro = Reader.ReadInt16()  'The iron needed
        ArmasHerrero(i).LPlata = Reader.ReadInt16()    'The silver needed
        ArmasHerrero(i).LOro = Reader.ReadInt16()    'The gold needed
        
        ' Call frmHerrero.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i).Index = 0
    Next i
    
    i = 0
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlacksmithWeapons", Erl)
    

End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = Reader.ReadInt16()
    
    'Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To count
        tmp = Reader.ReadString8()         'Get the object's name
        DefensasHerrero(i).LHierro = Reader.ReadInt16()   'The iron needed
        DefensasHerrero(i).LPlata = Reader.ReadInt16()   'The silver needed
        DefensasHerrero(i).LOro = Reader.ReadInt16()   'The gold needed
        
        ' Call frmHerrero.lstArmaduras.AddItem(tmp)
        DefensasHerrero(i).Index = Reader.ReadInt16()
    Next i
        
    Dim A      As Byte
    Dim e      As Byte
    Dim c      As Byte
    Dim tmpObj As ObjDatas

    A = 0
    e = 0
    c = 0
    
    For i = 1 To UBound(DefensasHerrero())

        If DefensasHerrero(i).Index = 0 Then Exit For
        
        tmpObj = ObjData(DefensasHerrero(i).Index)
        
        If tmpObj.ObjType = 3 Then
           
            ArmadurasHerrero(A).Index = DefensasHerrero(i).Index
            ArmadurasHerrero(A).LHierro = DefensasHerrero(i).LHierro
            ArmadurasHerrero(A).LPlata = DefensasHerrero(i).LPlata
            ArmadurasHerrero(A).LOro = DefensasHerrero(i).LOro
            A = A + 1

        End If
        
        ' Escudos (16), Objetos Magicos (21) y Anillos (35) van en la misma lista
        If tmpObj.ObjType = 16 Or tmpObj.ObjType = 35 Or tmpObj.ObjType = 21 Or tmpObj.ObjType = 100 Then
            EscudosHerrero(e).Index = DefensasHerrero(i).Index
            EscudosHerrero(e).LHierro = DefensasHerrero(i).LHierro
            EscudosHerrero(e).LPlata = DefensasHerrero(i).LPlata
            EscudosHerrero(e).LOro = DefensasHerrero(i).LOro
            e = e + 1

        End If

        If tmpObj.ObjType = 17 Then
            CascosHerrero(c).Index = DefensasHerrero(i).Index
            CascosHerrero(c).LHierro = DefensasHerrero(i).LHierro
            CascosHerrero(c).LPlata = DefensasHerrero(i).LPlata
            CascosHerrero(c).LOro = DefensasHerrero(i).LOro
            c = c + 1

        End If

    Next i
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlacksmithArmors", Erl)
    

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = Reader.ReadInt8()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To count
        ObjCarpintero(i) = Reader.ReadInt16()
        
        Call frmCarp.lstArmas.AddItem(ObjData(ObjCarpintero(i)).Name)
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCarpenterObjects", Erl)
    

End Sub

Private Sub HandleSastreObjects()

    '***************************************************
    'Author: Ladder
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = Reader.ReadInt16()
    
    For i = i To UBound(ObjSastre())
        ObjSastre(i).Index = 0
    Next i
    
    i = 0
    
    For i = 1 To count
        ObjSastre(i).Index = Reader.ReadInt16()
        
        ObjSastre(i).PielLobo = ObjData(ObjSastre(i).Index).PielLobo
        ObjSastre(i).PielOsoPardo = ObjData(ObjSastre(i).Index).PielOsoPardo
        ObjSastre(i).PielOsoPolar = ObjData(ObjSastre(i).Index).PielOsoPolar

    Next i
    
    Dim r As Byte

    Dim G As Byte
    
    i = 0
    r = 1
    G = 1
    
    For i = i To UBound(ObjSastre())
    
        If ObjData(ObjSastre(i).Index).ObjType = 3 Or ObjData(ObjSastre(i).Index).ObjType = 100 Then
        
            SastreRopas(r).Index = ObjSastre(i).Index
            SastreRopas(r).PielLobo = ObjSastre(i).PielLobo
            SastreRopas(r).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreRopas(r).PielOsoPolar = ObjSastre(i).PielOsoPolar
            r = r + 1

        End If

        If ObjData(ObjSastre(i).Index).ObjType = 17 Then
            SastreGorros(G).Index = ObjSastre(i).Index
            SastreGorros(G).PielLobo = ObjSastre(i).PielLobo
            SastreGorros(G).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreGorros(G).PielOsoPolar = ObjSastre(i).PielOsoPolar
            G = G + 1

        End If

    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSastreObjects", Erl)
    

End Sub

''
Private Sub HandleAlquimiaObjects()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    Dim Obj   As Integer

    count = Reader.ReadInt16()
    
    Call frmAlqui.lstArmas.Clear
    
    For i = 1 To count
        Obj = Reader.ReadInt16()
        tmp = ObjData(Obj).Name        'Get the object's name

        ObjAlquimista(i) = Obj
        Call frmAlqui.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ObjAlquimista())
        ObjAlquimista(i) = 0
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAlquimiaObjects", Erl)
    

End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRestOK_Err
    
    UserDescansar = Not UserDescansar
    
    Exit Sub

HandleRestOK_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleRestOK", Erl)
    
    
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Call MsgBox(Reader.ReadString8())
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleErrorMessage", Erl)
    

End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlind_Err
    
    UserCiego = True
    
    Call SetRGBA(global_light, 4, 4, 4)
    Call MapUpdateGlobalLight
    
    Exit Sub

HandleBlind_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlind", Erl)
    
    
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleDumb_Err
    
    UserEstupido = True
    
    Exit Sub

HandleDumb_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleDumb", Erl)
    
    
End Sub

''
' Handles the ShowSignal message.
'Optimizacion de protocolo por Ladder

Private Sub HandleShowSignal()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim tmp As String
    Dim grh As Integer

    tmp = ObjData(Reader.ReadInt16()).Texto
    grh = Reader.ReadInt16()
    
    Call InitCartel(tmp, grh)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowSignal", Erl)
    

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot As Byte
    Slot = Reader.ReadInt8()
    
    Dim SlotInv As NpCinV

    With SlotInv
        .ObjIndex = Reader.ReadInt16()
        .Name = ObjData(.ObjIndex).Name
        .Amount = Reader.ReadInt16()
        .Valor = Reader.ReadReal32()
        .GrhIndex = ObjData(.ObjIndex).GrhIndex
        .ObjType = ObjData(.ObjIndex).ObjType
        .MaxHit = ObjData(.ObjIndex).MaxHit
        .MinHit = ObjData(.ObjIndex).MinHit
        .Def = ObjData(.ObjIndex).MaxDef
        .PuedeUsar = Reader.ReadInt8()
        
        Call frmComerciar.InvComNpc.SetItem(Slot, .ObjIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)
        
    End With
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChangeNPCInventorySlot", Erl)
    

End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
    
    On Error GoTo HandleUpdateHungerAndThirst_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserMaxAGU = Reader.ReadInt8()
    UserMinAGU = Reader.ReadInt8()
    UserMaxHAM = Reader.ReadInt8()
    UserMinHAM = Reader.ReadInt8()
    frmMain.AGUAsp.Width = UserMinAGU / UserMaxAGU * 47
    frmMain.COMIDAsp.Width = UserMinHAM / UserMaxHAM * 47
    frmMain.AGUbar.Caption = UserMinAGU '& " / " & UserMaxAGU
    frmMain.hambar.Caption = UserMinHAM ' & " / " & UserMaxHAM
    
    If QuePesta�aInferior = 0 Then
        frmMain.AGUAsp.visible = (UserMinAGU > 0)
        frmMain.COMIDAsp.visible = (UserMinHAM > 0)

    End If

    Exit Sub

HandleUpdateHungerAndThirst_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateHungerAndThirst", Erl)
    
    
End Sub

Private Sub HandleHora()
    '***************************************************
    
    On Error GoTo HandleHora_Err

    HoraMundo = GetTickCount() - Reader.ReadInt32()
    DuracionDia = Reader.ReadInt32()
    
    If Not Connected Then
        Call RevisarHoraMundo(True)

    End If
    
    Exit Sub

HandleHora_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleHora", Erl)
    
    
End Sub
 
Private Sub HandleLight()
    
    On Error GoTo HandleLight_Err
 
    Dim Color As String
    
    Color = Reader.ReadString8()

    'Call SetGlobalLight(Map_light_base)
    
    Exit Sub

HandleLight_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleLight", Erl)
    
    
End Sub
 
Private Sub HandleFYA()
    
    On Error GoTo HandleFYA_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserAtributos(eAtributos.Fuerza) = Reader.ReadInt8()
    UserAtributos(eAtributos.Agilidad) = Reader.ReadInt8()
    
    DrogaCounter = Reader.ReadInt16()
    
    If DrogaCounter > 0 Then
        frmMain.Contadores.enabled = True

    End If
    
    If UserAtributos(eAtributos.Fuerza) >= 35 Then
        frmMain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserAtributos(eAtributos.Fuerza) >= 25 Then
        frmMain.Fuerzalbl.ForeColor = RGB(204, 100, 100)
    Else
        frmMain.Fuerzalbl.ForeColor = vbWhite

    End If
    
    If UserAtributos(eAtributos.Agilidad) >= 35 Then
        frmMain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserAtributos(eAtributos.Agilidad) >= 25 Then
        frmMain.AgilidadLbl.ForeColor = RGB(204, 100, 100)
    Else
        frmMain.AgilidadLbl.ForeColor = vbWhite

    End If

    frmMain.Fuerzalbl.Caption = UserAtributos(eAtributos.Fuerza)
    frmMain.AgilidadLbl.Caption = UserAtributos(eAtributos.Agilidad)
    
    Exit Sub

HandleFYA_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleFYA", Erl)
    
    
End Sub

Private Sub HandleUpdateNPCSimbolo()
    
    On Error GoTo HandleUpdateNPCSimbolo_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim NpcIndex As Integer

    Dim simbolo  As Byte
    
    NpcIndex = Reader.ReadInt16()
    
    simbolo = Reader.ReadInt8()

    charlist(NpcIndex).simbolo = simbolo
    
    Exit Sub

HandleUpdateNPCSimbolo_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateNPCSimbolo", Erl)
    
    
End Sub

Private Sub HandleCerrarleCliente()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCerrarleCliente_Err
    
    EngineRun = False

    Call CloseClient
    
    Exit Sub

HandleCerrarleCliente_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCerrarleCliente", Erl)
    
    
End Sub

Private Sub HandleContadores()
    
    On Error GoTo HandleContadores_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    InviCounter = Reader.ReadInt16()
    DrogaCounter = Reader.ReadInt16()
    

    frmMain.Contadores.enabled = True
    
    Exit Sub

HandleContadores_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleContadores", Erl)
    
    
End Sub

Private Sub HandleShowPapiro()
    On Error GoTo HandleShowPapiro_Err
    
    frmMensajePapiro.Show , frmMain
    
    'incomingdata papiromessage
    Exit Sub

HandleShowPapiro_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowPapiro", Erl)
End Sub


''
' Handles the MiniStats message.
Private Sub HandleFlashScreen()
    
    On Error GoTo HandleEfectToScreen_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Color As Long, duracion As Long, ignorar As Boolean
    
    Color = Reader.ReadInt32()
    duracion = Reader.ReadInt32()
    ignorar = Reader.ReadBool()
    
    Dim r, G, B As Byte

    B = (Color And 16711680) / 65536
    G = (Color And 65280) / 256
    r = Color And 255
    Color = D3DColorARGB(255, r, G, B)

    If Not MapDat.Niebla = 1 And Not ignorar Then
        'Debug.Print "trueno cancelado"
       
        Exit Sub

    End If

    Call EfectoEnPantalla(Color, duracion)
    
    Exit Sub

HandleEfectToScreen_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleEfectToScreen", Erl)
    
    
End Sub

Private Sub HandleMiniStats()
    
    On Error GoTo HandleMiniStats_Err
    
    With UserEstadisticas
        .CiudadanosMatados = Reader.ReadInt32()
        .CriminalesMatados = Reader.ReadInt32()
        .Alineacion = Reader.ReadInt8()
        
        .NpcsMatados = Reader.ReadInt32()
        .Clase = ListaClases(Reader.ReadInt8())
        .PenaCarcel = Reader.ReadInt32()
        .VecesQueMoriste = Reader.ReadInt32()
        .Genero = Reader.ReadInt8()
        .PuntosPesca = Reader.ReadInt32()

        If .Genero = 1 Then
            .Genero = "Hombre"
        Else
            .Genero = "Mujer"

        End If

        .Raza = Reader.ReadInt8()
        .Raza = ListaRazas(.Raza)
    End With
    
    If LlegaronAtrib Then
        frmStatistics.Iniciar_Labels
        frmStatistics.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
        frmStatistics.Show , frmMain
    Else
        LlegaronStats = True
    End If
    
    Exit Sub

HandleMiniStats_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleMiniStats", Erl)
    
    
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
    
    On Error GoTo HandleLevelUp_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    SkillPoints = Reader.ReadInt16()
    
    Exit Sub

HandleLevelUp_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleLevelUp", Erl)
    
    
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim title   As String

    Dim message As String
    
    title = Reader.ReadString8()
    message = Reader.ReadString8()

    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAddForumMessage", Erl)
    

End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowForumForm_Err
    
    ' If Not frmForo.Visible Then
    '   frmForo.Show , frmMain
    ' End If
    
    Exit Sub

HandleShowForumForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowForumForm", Erl)
    
    
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
    
    On Error GoTo HandleSetInvisible_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer
    Dim x As Integer, y As Integer
    charindex = Reader.ReadInt16()
    charlist(charindex).Invisible = Reader.ReadBool()
    charlist(charindex).TimerI = 0
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If x + y > 0 Then
        With charlist(charindex)
            If Not .Invisible And charindex <> UserCharIndex Then
                If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(rrX(x), rrY(y)).charindex = charindex
                If Abs(.MoveOffsetX) > 32 Or Abs(.MoveOffsetY) > 32 Or (.MoveOffsetX <> 0 And .MoveOffsetY <> 0) Then
                    .MoveOffsetX = 0
                    .MoveOffsetY = 0
                End If
            End If
        End With
    End If
    
    Exit Sub

HandleSetInvisible_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSetInvisible", Erl)
    
    
End Sub


''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleMeditateToggle_Err
    
    Dim charindex As Integer, fX As Integer
    
    Dim x As Integer, y As Integer
    
    charindex = Reader.ReadInt16
    fX = Reader.ReadInt16
    x = Reader.ReadInt16
    y = Reader.ReadInt16
    
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(rrX(x), rrY(y)).charindex = charindex
            End If
        End With
    End If
    
    If charindex = UserCharIndex Then
        UserMeditar = (fX <> 0)
        
        If UserMeditar Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Comienzas a meditar.", .red, .green, .blue, .bold, .italic)

            End With

        Else

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Has dejado de meditar.", .red, .green, .blue, .bold, .italic)

            End With

        End If

    End If
    
    With charlist(charindex)

        If fX <> 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)

        End If
        
        .FxIndex = fX
        .fX.Loops = -1
        .fX.AnimacionContador = 0

    End With
    
    Exit Sub

HandleMeditateToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleMeditateToggle", Erl)
    
    
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlindNoMore_Err
    
    UserCiego = False
    
    Call RestaurarLuz
    
    Exit Sub

HandleBlindNoMore_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBlindNoMore", Erl)
    
    
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleDumbNoMore_Err
    
    UserEstupido = False
    
    Exit Sub

HandleDumbNoMore_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleDumbNoMore", Erl)
    
    
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
    
    On Error GoTo HandleSendSkills_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = Reader.ReadInt8()
        UserSkillsAssigned(i) = Reader.ReadInt8()
        'frmEstadisticas.skills(i).Caption = SkillsNames(i)
    Next i

    If LlegaronSkills Then
        Alocados = SkillPoints
        frmEstadisticas.Puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Picture = LoadInterface("ventanaskills.bmp")
        frmEstadisticas.Show , frmMain
        LlegaronSkills = False
    End If
    
    Exit Sub

HandleSendSkills_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSendSkills", Erl)
    
    
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim creatures() As String

    Dim i           As Long
    
    creatures = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i

    frmEntrenador.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleTrainerCreatureList", Erl)
    

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    ' Dim guildList() As String
    Dim List()      As String

    Dim i           As Long
    
    Dim ClanNivel   As Byte

    Dim expacu      As Integer

    Dim ExpNe       As Integer

    Dim guildList() As String
        
    frmGuildNews.news = Reader.ReadString8()
    
    'Get list of existing guilds
    List = Split(Reader.ReadString8(), SEPARATOR)
        
    'Empty the list
    Call frmGuildNews.guildslist.Clear
        
    For i = 0 To UBound(List())
        Call frmGuildNews.guildslist.AddItem(ReadField(1, List(i), Asc("-")))
    Next i
    
    'Get  guilds list member
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    
    Dim cantidad As String

    cantidad = CStr(UBound(guildList()) + 1)
        
    Call frmGuildNews.miembros.Clear
        
    For i = 0 To UBound(guildList())

        If i = 0 Then
            Call frmGuildNews.miembros.AddItem(guildList(i) & "(Lider)")
        Else
            Call frmGuildNews.miembros.AddItem(guildList(i))

        End If

        'Debug.Print guildList(i)
    Next i
    
    ClanNivel = Reader.ReadInt8()
    expacu = Reader.ReadInt16()
    ExpNe = Reader.ReadInt16()
     
    With frmGuildNews
        .Frame4.Caption = "Total: " & cantidad & " miembros" '"Lista de miembros" ' - " & cantidad & " totales"
     
        .expcount.Caption = expacu & "/" & ExpNe
        .EXPBAR.Width = (((expacu + 1 / 100) / (ExpNe + 1 / 100)) * 2370)
        .nivel = "Nivel: " & ClanNivel

        If ExpNe > 0 Then
       
            .porciento.Caption = Round(CDbl(expacu) * CDbl(100) / CDbl(ExpNe), 0) & "%"
        Else
            .porciento.Caption = "�Nivel Maximo!"
            .expcount.Caption = "�Nivel Maximo!"

        End If
        
        '.expne = "Experiencia necesaria: " & expne
        
        Select Case ClanNivel

            Case 1
                .beneficios = "Max miembros: 5"

            Case 2
                .beneficios = "Pedir ayuda (G) / Max miembros: 7"

            Case 3
                .beneficios = "Pedir ayuda (G) / Seguro de clan." & vbCrLf & "Max miembros: 7"

            Case 4
                .beneficios = "Pedir ayuda (G) / Seguro de clan. " & vbCrLf & "Max miembros: 12"

            Case 5
                .beneficios = "Pedir ayuda (G) / Seguro de clan /  Ver vida y mana." & vbCrLf & " Max miembros: 15"
                
            Case 6
                .beneficios = "Pedir ayuda (G) / Seguro de clan / Ver vida y mana/ Verse invisible." & vbCrLf & " Max miembros: 20"
        
        End Select
    
    End With
    
    frmGuildNews.Show vbModeless, frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleGuildNews", Erl)
    

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleOfferDetails", Erl)
    

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAlianceProposalsList", Erl)
    

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePeaceProposalsList", Erl)
    

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    With frmCharInfo

        If .frmType = CharInfoFrmType.frmMembers Then
            .Rechazar.visible = False
            .Aceptar.visible = False
            .Echar.visible = True
            .desc.visible = False
        Else
            .Rechazar.visible = True
            .Aceptar.visible = True
            .Echar.visible = False
            .desc.visible = True

        End If
    
        If Reader.ReadInt8() = 1 Then
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"
        End If
            
        .nombre.Caption = "Nombre: " & Reader.ReadString8()
        .Raza.Caption = "Raza: " & ListaRazas(Reader.ReadInt8())
        .Clase.Caption = "Clase: " & ListaClases(Reader.ReadInt8())

        .nivel.Caption = "Nivel: " & Reader.ReadInt8()
        .oro.Caption = "Oro: " & Reader.ReadInt32()
        .Banco.Caption = "Banco: " & Reader.ReadInt32()
    
        .txtPeticiones.Text = Reader.ReadString8()
        .guildactual.Caption = "Clan: " & Reader.ReadString8()
        .txtMiembro.Text = Reader.ReadString8()
            
        Dim armada As Boolean
    
        Dim caos   As Boolean
            
        armada = Reader.ReadBool()
        caos = Reader.ReadBool()
            
        If armada Then
            .ejercito.Caption = "Ej�rcito: Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Ej�rcito: Legi�n Oscura"
    
        End If
            
        .ciudadanos.Caption = "Ciudadanos asesinados: " & CStr(Reader.ReadInt32())
        .Criminales.Caption = "Criminales asesinados: " & CStr(Reader.ReadInt32())
    
        Call .Show(vbModeless, frmMain)
    
    End With
        
    Exit Sub
    
ErrHandler:
    
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCharacterInfo", Erl)
    

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim str As String
    
    Dim List() As String

    Dim i      As Long
    
    With frmGuildLeader
        'Empty the list
        Call .guildslist.Clear
    
        str = Reader.ReadString8()
    
        If LenB(str) > 0 Then
            'Get list of existing guilds
            List = Split(str, SEPARATOR)

            For i = 0 To UBound(List())
                Call .guildslist.AddItem(ReadField(1, List(i), Asc("-")))
            Next i
        End If
        
        'Empty the list
        Call .members.Clear
        
        str = Reader.ReadString8()
        
        If LenB(str) > 0 Then
            'Get list of guild's members
            List = Split(str, SEPARATOR)
            .miembros.Caption = CStr(UBound(List()) + 1)

            For i = 0 To UBound(List())
                Call .members.AddItem(List(i))
            Next i
        End If
        
        .txtguildnews = Reader.ReadString8()
        
        'Empty the list
        Call .solicitudes.Clear
        
        str = Reader.ReadString8()
        
        If LenB(str) > 0 Then
            'Get list of join requests
            List = Split(str, SEPARATOR)
        
            For i = 0 To UBound(List())
                Call .solicitudes.AddItem(List(i))
            Next i
        End If
        
        Dim expacu As Integer

        Dim ExpNe  As Integer

        Dim nivel  As Byte
         
        nivel = Reader.ReadInt8()
        .nivel = "Nivel: " & nivel
        
        expacu = Reader.ReadInt16()
        ExpNe = Reader.ReadInt16()
        'barra
        .expcount.Caption = expacu & "/" & ExpNe
        
        If ExpNe > 0 Then
            .EXPBAR.Width = expacu / ExpNe * 239
            .porciento.Caption = Round(expacu / ExpNe * 100#, 0) & "%"
        Else
            .EXPBAR.Width = 239
            .porciento.Caption = "�Nivel m�ximo!"
            .expcount.Caption = "�Nivel m�ximo!"
        End If

        Select Case nivel

               Case 1
                .beneficios = "Max miembros: 5"
                .maxMiembros = 5
            Case 2
                .beneficios = "Pedir ayuda (G) / Max miembros: 7"
                .maxMiembros = 7

            Case 3
                .beneficios = "Pedir ayuda (G) / Seguro de clan." & vbCrLf & "Max miembros: 7"
                .maxMiembros = 7

            Case 4
                .beneficios = "Pedir ayuda (G) / Seguro de clan. " & vbCrLf & "Max miembros: 12"
                .maxMiembros = 12

            Case 5
                .beneficios = "Pedir ayuda (G) / Seguro de clan /  Ver vida y mana." & vbCrLf & " Max miembros: 15"
                .maxMiembros = 15
                
            Case 6
                .beneficios = "Pedir ayuda (G) / Seguro de clan / Ver vida y mana/ Verse invisible." & vbCrLf & " Max miembros: 20"
                .maxMiembros = 20
        End Select
        
        .Show , frmMain

    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleGuildLeaderInfo", Erl)
    

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    With frmGuildBrief

        If Not .EsLeader Then

        End If
        
        .nombre.Caption = "Nombre:" & Reader.ReadString8()
        .fundador.Caption = "Fundador:" & Reader.ReadString8()
        .creacion.Caption = "Fecha de creacion:" & Reader.ReadString8()
        .lider.Caption = "L�der:" & Reader.ReadString8()
        .miembros.Caption = "Miembros:" & Reader.ReadInt16()
        
        .lblAlineacion.Caption = "Alineaci�n: " & Reader.ReadString8()
        
        .desc.Text = Reader.ReadString8()
        .nivel.Caption = "Nivel de clan: " & Reader.ReadInt8()

    End With
    
    frmGuildBrief.Show vbModeless, frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleGuildDetails", Erl)
    

End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowGuildFundationForm_Err
    
    CreandoClan = True
    frmGuildDetails.Show , frmMain
    
    Exit Sub

HandleShowGuildFundationForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowGuildFundationForm", Erl)
    
    
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleParalizeOK_Err
    If EstaSiguiendo Then Exit Sub
    UserParalizado = Not UserParalizado
    
    Exit Sub

HandleParalizeOK_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleParalizeOK", Erl)
    
    
End Sub

Private Sub HandleInmovilizadoOK()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    On Error GoTo HandleInmovilizadoOK_Err
    If EstaSiguiendo Then Exit Sub
    UserInmovilizado = Not UserInmovilizado
    
    Exit Sub

HandleInmovilizadoOK_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleInmovilizadoOK", Erl)
    
    
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowUserRequest", Erl)
    

End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim miOferta As Boolean
    
    miOferta = Reader.ReadBool
    Dim i          As Byte
    Dim nombreItem As String
    Dim cantidad   As Integer
    Dim grhItem    As Long
    Dim ObjIndex   As Integer

    If miOferta Then
        Dim OroAEnviar As Long
        OroAEnviar = Reader.ReadInt32
        frmComerciarUsu.lblOroMiOferta.Caption = PonerPuntos(OroAEnviar)
        frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserGLD - OroAEnviar)

        For i = 1 To 6

            With OtroInventario(i)
                ObjIndex = Reader.ReadInt16
                nombreItem = Reader.ReadString8
                grhItem = Reader.ReadInt32
                cantidad = Reader.ReadInt32

                If cantidad > 0 Then
                    Call frmComerciarUsu.InvUserSell.SetItem(i, ObjIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)

                End If

            End With

        Next i
        
        Call frmComerciarUsu.InvUserSell.ReDraw
    Else
        frmComerciarUsu.lblOro.Caption = PonerPuntos(Reader.ReadInt32)

        ' frmComerciarUsu.List2.Clear
        For i = 1 To 6
            
            With OtroInventario(i)
                ObjIndex = Reader.ReadInt16
                nombreItem = Reader.ReadString8
                grhItem = Reader.ReadInt32
                cantidad = Reader.ReadInt32

                If cantidad > 0 Then
                    Call frmComerciarUsu.InvOtherSell.SetItem(i, ObjIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)

                End If

            End With

        Next i
        
        Call frmComerciarUsu.InvOtherSell.ReDraw
    
    End If
    
    frmComerciarUsu.lblEstadoResp.visible = False
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleChangeUserTradeSlot", Erl)
    

End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    frmSpawnList.ListaCompleta = Reader.ReadBool

    Call frmSpawnList.FillList

    frmSpawnList.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSpawnList", Erl)
    

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim sosList()      As String

    Dim i              As Long

    Dim nombre         As String

    Dim Consulta       As String

    Dim TipoDeConsulta As String
    
    sosList = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        nombre = ReadField(1, sosList(i), Asc("�"))
        Consulta = ReadField(2, sosList(i), Asc("�"))
        TipoDeConsulta = ReadField(3, sosList(i), Asc("�"))
        frmPanelgm.List1.AddItem nombre & "(" & TipoDeConsulta & ")"
        frmPanelgm.List2.AddItem Consulta
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowSOSForm", Erl)
    

End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    frmCambiaMotd.txtMotd.Text = Reader.ReadString8()
    frmCambiaMotd.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowMOTDEditionForm", Erl)
    

End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowGMPanelForm_Err
    
    Dim MiCargo As Integer
    
    
    frmPanelgm.txtHeadNumero = Reader.ReadInt16
    frmPanelgm.txtBodyYo = Reader.ReadInt16
    frmPanelgm.txtCasco = Reader.ReadInt16
    frmPanelgm.txtArma = Reader.ReadInt16
    frmPanelgm.txtEscudo = Reader.ReadInt16
    frmPanelgm.Show vbModeless, frmMain
    
    MiCargo = charlist(UserCharIndex).priv
    
    Select Case MiCargo ' ReyarB ajustar privilejios
    
        Case 1
        frmPanelgm.mnuChar.visible = False
        frmPanelgm.cmdHerramientas.visible = False
        frmPanelgm.Admin(0).visible = False
        
        Case 2 'Consejeros
        frmPanelgm.mnuChar.visible = False
        frmPanelgm.cmdHerramientas.visible = False
        frmPanelgm.Admin(0).visible = False
        frmPanelgm.cmdConsulta.visible = False
        frmPanelgm.cmdMatarNPC.visible = False
        frmPanelgm.cmdEventos.visible = False
        frmPanelgm.cmdBody0(2).visible = False
        frmPanelgm.cmdHead0.visible = False
        frmPanelgm.SendGlobal.visible = False
        frmPanelgm.Mensajeria.visible = False
        frmPanelgm.cmdMapeo.visible = False
        frmPanelgm.cmdMapeo.enabled = False
        frmPanelgm.cmdcrearevento.enabled = False
        frmPanelgm.cmdcrearevento.visible = False
        frmPanelgm.txtMod.Width = 4580
        frmPanelgm.Height = 7580
        frmPanelgm.mnuTraer.visible = False
        frmPanelgm.mnuIra.visible = False
                
        Case 3 ' Semidios
        frmPanelgm.mnuChar.visible = False
        frmPanelgm.mnuChar.visible = False
        frmPanelgm.cmdHerramientas.visible = True
        frmPanelgm.Admin(0).visible = False
        frmPanelgm.cmdcrearevento.enabled = False
        frmPanelgm.cmdcrearevento.visible = False
        frmPanelgm.mnuHerramientas(23).visible = False
        
        Case 4 ' Dios
        frmPanelgm.mnuChar.visible = True
        frmPanelgm.mnuChar.visible = True
        frmPanelgm.cmdHerramientas.visible = True
        frmPanelgm.Admin(0).visible = False
        
        Case 5
        
    
    End Select
    
    Exit Sub

HandleShowGMPanelForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowGMPanelForm", Erl)
    
    
End Sub

Private Sub HandleShowFundarClanForm()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowFundarClanForm_Err
    
    CreandoClan = True
    frmGuildDetails.Show vbModeless, frmMain
    
    Exit Sub

HandleShowFundarClanForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowFundarClanForm", Erl)
    
    
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim userList() As String

    Dim i          As Long
    
    userList = Split(Reader.ReadString8(), SEPARATOR)
    
    If frmPanelgm.visible Then
        frmPanelgm.cboListaUsus.Clear

        For i = 0 To UBound(userList())
            Call frmPanelgm.cboListaUsus.AddItem(userList(i))
        Next i

        If frmPanelgm.cboListaUsus.ListCount > 0 Then frmPanelgm.cboListaUsus.ListIndex = 0

    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserNameList", Erl)
    

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim charindex   As Integer

    Dim status      As Byte

    Dim NombreYClan As String

    Dim group_index As Integer
    
    charindex = Reader.ReadInt16()
    status = Reader.ReadInt8()
    NombreYClan = Reader.ReadString8()
        
    Dim Pos As Integer
    Pos = InStr(NombreYClan, "<")

    If Pos = 0 Then Pos = InStr(NombreYClan, "[")
    If Pos = 0 Then Pos = Len(NombreYClan) + 2
    
    charlist(charindex).nombre = Left$(NombreYClan, Pos - 2)
    charlist(charindex).clan = mid$(NombreYClan, Pos)
    
    group_index = Reader.ReadInt16()
    
    'Update char status adn tag!
    charlist(charindex).status = status
    
    charlist(charindex).group_index = group_index
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateTagAndStatus", Erl)
    

End Sub

Private Sub HandleUserOnline()
    
    On Error GoTo ErrHandler

    Dim rdata As Integer
    
    rdata = Reader.ReadInt16()
    
    usersOnline = rdata
    frmMain.onlines = "Online: " & usersOnline
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUserOnline", Erl)
    

End Sub

Private Sub HandleParticleFXToFloor()
    
    On Error GoTo HandleParticleFXToFloor_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Integer

    Dim y As Integer

    Dim ParticulaIndex As Byte

    Dim Time           As Long

    Dim Borrar         As Boolean
     
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadInt16()
    Time = Reader.ReadInt32()

    If Time = 1 Then
        Time = -1

    End If
    
    If Time = 0 Then
        Borrar = True

    End If

    If Borrar Then
        Graficos_Particulas.Particle_Group_Remove (MapData(rrX(x), rrY(y)).particle_group)
    Else

        If MapData(rrX(x), rrY(y)).particle_group = 0 Then
            MapData(rrX(x), rrY(y)).particle_group = 0
            General_Particle_Create ParticulaIndex, x, y, Time
        Else
            Call General_Char_Particle_Create(ParticulaIndex, MapData(rrX(x), rrY(y)).charindex, Time)

        End If

    End If
    
    Exit Sub

HandleParticleFXToFloor_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleParticleFXToFloor", Erl)
    
    
End Sub

Private Sub HandleLightToFloor()
    
    On Error GoTo HandleLightToFloor_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Integer

    Dim y As Integer

    Dim Color       As Long
    
    Dim color_value As RGBA

    Dim Rango       As Byte
     
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    Color = Reader.ReadInt32()
    Rango = Reader.ReadInt8()
    
    Call Long_2_RGBA(color_value, Color)

    Dim id  As Long

    Dim id2 As Long

    If Color = 0 Then
   
        If MapData(rrX(x), rrY(y)).luz.Rango > 100 Then
            LucesRedondas.Delete_Light_To_Map x, y
            Exit Sub
        Else
            id = LucesCuadradas.Light_Find(x & y)
            LucesCuadradas.Light_Remove id
            MapData(rrX(x), rrY(y)).luz.Color = COLOR_EMPTY
            MapData(rrX(x), rrY(y)).luz.Rango = 0
            Exit Sub

        End If

    End If
    
    MapData(rrX(x), rrY(y)).luz.Color = color_value
    MapData(rrX(x), rrY(y)).luz.Rango = Rango
    
    If Rango < 100 Then
        id = x & y
        LucesCuadradas.Light_Create x, y, color_value, Rango, id
    Else

        LucesRedondas.Create_Light_To_Map x, y, color_value, Rango - 99
    End If
    
    Exit Sub

HandleLightToFloor_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleLightToFloor", Erl)
    
    
End Sub

Private Sub HandleParticleFX()
    
    On Error GoTo HandleParticleFX_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex      As Integer

    Dim ParticulaIndex As Integer

    Dim Time           As Long

    Dim Remove         As Boolean
    Dim grh            As Long
    
    Dim x As Integer, y As Integer
    
    charindex = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadInt16()
    Time = Reader.ReadInt32()
    Remove = Reader.ReadBool()
    grh = Reader.ReadInt32()
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = charindex Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(rrX(x), rrY(y)).charindex = charindex
            End If
        End With
    End If
    If Remove Then
        Call Char_Particle_Group_Remove(charindex, ParticulaIndex)
        charlist(charindex).Particula = 0
    
    Else
        charlist(charindex).Particula = ParticulaIndex
        charlist(charindex).ParticulaTime = Time
        If grh > 0 Then
            Call General_Char_Particle_Create(ParticulaIndex, charindex, Time, grh)
        Else
            Call General_Char_Particle_Create(ParticulaIndex, charindex, Time)
        End If

    End If
    
    Exit Sub

HandleParticleFX_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleParticleFX", Erl)
    
    
End Sub

Private Sub HandleParticleFXWithDestino()
    
    On Error GoTo HandleParticleFXWithDestino_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Emisor         As Integer

    Dim receptor       As Integer

    Dim ParticulaViaje As Integer

    Dim ParticulaFinal As Integer

    Dim Time           As Long

    Dim wav            As Integer

    Dim fX             As Integer
    
    Dim x As Integer, y As Integer
    Emisor = Reader.ReadInt16()
    receptor = Reader.ReadInt16()
    ParticulaViaje = Reader.ReadInt16()
    ParticulaFinal = Reader.ReadInt16()

    Time = Reader.ReadInt32()
    wav = Reader.ReadInt16()
    fX = Reader.ReadInt16()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    If x + y > 0 Then
        With charlist(receptor)
           If .Invisible And receptor <> UserCharIndex Then
                If MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = receptor Then MapData(rrX(.Pos.x), rrY(.Pos.y)).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(rrX(x), rrY(y)).charindex = receptor
            End If
        End With
    End If

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_Begin(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), ParticulaFinal, Time, receptor, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
    Exit Sub

HandleParticleFXWithDestino_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleParticleFXWithDestino", Erl)
    
    
End Sub

Private Sub HandleParticleFXWithDestinoXY()
    
    On Error GoTo HandleParticleFXWithDestinoXY_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Emisor         As Integer

    Dim ParticulaViaje As Integer

    Dim ParticulaFinal As Integer

    Dim Time           As Long

    Dim wav            As Integer

    Dim fX             As Integer

    Dim x As Integer

    Dim y As Integer
     
    Emisor = Reader.ReadInt16()
    ParticulaViaje = Reader.ReadInt16()
    ParticulaFinal = Reader.ReadInt16()

    Time = Reader.ReadInt32()
    wav = Reader.ReadInt16()
    fX = Reader.ReadInt16()
    
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    
    ' Debug.Print "RECIBI FX= " & fX

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_BeginXY(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), x, y, ParticulaFinal, Time, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
    Exit Sub

HandleParticleFXWithDestinoXY_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleParticleFXWithDestinoXY", Erl)
    
    
End Sub

Private Sub HandleAuraToChar()
    
    On Error GoTo HandleAuraToChar_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    
    Dim charindex      As Integer

    Dim ParticulaIndex As String

    Dim Remove         As Boolean

    Dim TIPO           As Byte
     
    charindex = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadString8()

    Remove = Reader.ReadBool()
    TIPO = Reader.ReadInt8()
    
    If TIPO = 1 Then
        charlist(charindex).Arma_Aura = ParticulaIndex
    ElseIf TIPO = 2 Then
        charlist(charindex).Body_Aura = ParticulaIndex
    ElseIf TIPO = 3 Then
        charlist(charindex).Escudo_Aura = ParticulaIndex
    ElseIf TIPO = 4 Then
        charlist(charindex).Head_Aura = ParticulaIndex
    ElseIf TIPO = 5 Then
        charlist(charindex).Otra_Aura = ParticulaIndex
    ElseIf TIPO = 6 Then
        charlist(charindex).DM_Aura = ParticulaIndex
    Else
        charlist(charindex).RM_Aura = ParticulaIndex

    End If
    
    Exit Sub

HandleAuraToChar_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAuraToChar", Erl)
    
    
End Sub

Private Sub HandleSpeedToChar()
    
    On Error GoTo HandleSpeedToChar_Err

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    
    Dim charindex As Integer

    Dim Speeding  As Single
     
    charindex = Reader.ReadInt16()
    Speeding = Reader.ReadReal32()
   
    charlist(charindex).Speeding = Speeding
    
    Exit Sub

HandleSpeedToChar_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleSpeedToChar", Erl)
    
    
End Sub
Private Sub HandleNieveToggle()
    '**
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '**
    'Remove packet ID

    On Error GoTo HandleNieveToggle_Err

    bNieve = Reader.ReadBool

    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub



    If Not bNieve Then
        If MapDat.Nieve Then
            Engine_MeteoParticle_Set (-1)
        End If
    Else
        If MapDat.Nieve Then
            Engine_MeteoParticle_Set (Particula_Nieve)
        End If
    End If



    Exit Sub

HandleNieveToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNieveToggle", Erl)
    
    
End Sub

Private Sub HandleNieblaToggle()
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNieblaToggle_Err
    
    MaxAlphaNiebla = Reader.ReadInt8()
            
    bNiebla = Not bNiebla
    frmMain.TimerNiebla.enabled = True
    
    Exit Sub

HandleNieblaToggle_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNieblaToggle", Erl)
    
    
End Sub


Private Sub HandleBindKeys()
    
    On Error GoTo HandleBindKeys_Err

    '***************************************************
    'Macros
    'Pablo Mercavides
    '***************************************************
    
    ChatCombate = Reader.ReadInt8()
    ChatGlobal = Reader.ReadInt8()

    If ChatCombate = 1 Then
        frmMain.CombateIcon.Picture = LoadInterface("infoapretado.bmp")
    Else
        frmMain.CombateIcon.Picture = LoadInterface("info.bmp")

    End If

    If ChatGlobal = 1 Then
        frmMain.globalIcon.Picture = LoadInterface("globalapretado.bmp")
    Else
        frmMain.CombateIcon.Picture = LoadInterface("global.bmp")

    End If
    
    Exit Sub

HandleBindKeys_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBindKeys", Erl)
    
    
End Sub

Private Sub HandleBarFx()
    
    On Error GoTo HandleBarFx_Err

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    
    Dim charindex As Integer

    Dim BarTime   As Integer

    Dim BarAccion As Byte
    
    charindex = Reader.ReadInt16()
    BarTime = Reader.ReadInt16()
    BarAccion = Reader.ReadInt8()
    
    charlist(charindex).BarTime = 0
    charlist(charindex).BarAccion = BarAccion
    charlist(charindex).MaxBarTime = BarTime
    
    Exit Sub

HandleBarFx_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleBarFx", Erl)
    
    
End Sub
 
Private Sub HandleQuestDetails()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestDetails del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo ErrHandler
    
    Dim tmpStr         As String

    Dim tmpByte        As Byte

    Dim QuestEmpezada  As Boolean

    Dim i              As Integer
    
    Dim cantidadnpc    As Integer

    Dim NpcIndex       As Integer
    
    Dim cantidadobj    As Integer

    Dim ObjIndex       As Integer
    
    Dim AmountHave     As Integer
    
    Dim QuestIndex     As Integer
    
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    
    FrmQuests.ListView2.ListItems.Clear
    FrmQuests.ListView1.ListItems.Clear
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    FrmQuests.PlayerView.BackColor = RGB(11, 11, 11)
    FrmQuests.Picture1.BackColor = RGB(19, 14, 11)
    FrmQuests.PlayerView.Refresh
    FrmQuests.Picture1.Refresh
    FrmQuests.npclbl.Caption = ""
    FrmQuests.objetolbl.Caption = ""
    
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(Reader.ReadInt8, True, False)
        
        If Not QuestEmpezada Then
        
            QuestIndex = Reader.ReadInt16
        
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
           
            'tmpStr = "Mision: " & .ReadString8 & vbCrLf
            
            LevelRequerido = Reader.ReadInt8
            QuestRequerida = Reader.ReadInt16
           
            If QuestRequerida <> 0 Then
                FrmQuestInfo.Text1.Text = ""
               Call AddToConsole(FrmQuestInfo.Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf & "Quest:" & QuestList(QuestRequerida).RequiredQuest, 128, 128, 128)
            Else
                
                FrmQuestInfo.Text1.Text = ""
                Call AddToConsole(FrmQuestInfo.Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf, 128, 128, 128)
            End If
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay NPCs
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
           
                End If

                For i = 1 To tmpByte
                    cantidadnpc = Reader.ReadInt16
                    NpcIndex = Reader.ReadInt16
               
                    ' tmpStr = tmpStr & "*) Matar " & .ReadInt16 & " " & .ReadString8 & "."
                    If QuestEmpezada Then
                        tmpStr = tmpStr & " (Has matado " & Reader.ReadInt16 & ")" & vbCrLf
                    Else
                        tmpStr = tmpStr & vbCrLf
                       
                        Dim subelemento As ListItem

                        Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(NpcIndex).Name)
                       
                        subelemento.SubItems(1) = cantidadnpc
                        subelemento.SubItems(2) = NpcIndex
                        subelemento.SubItems(3) = 0

                    End If

                Next i

            End If
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay OBJs

                For i = 1 To tmpByte
               
                    cantidadobj = Reader.ReadInt16
                    ObjIndex = Reader.ReadInt16
                    
                    AmountHave = Reader.ReadInt16
                   
                    Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(ObjIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = ObjIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadInt32 & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadInt32 & " puntos de experiencia." & vbCrLf
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")

            subelemento.SubItems(1) = BeautifyBigNumber(Reader.ReadInt32)
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0

            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")

            subelemento.SubItems(1) = BeautifyBigNumber(Reader.ReadInt32)
            subelemento.SubItems(2) = 608
            subelemento.SubItems(3) = 1
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then

                For i = 1 To tmpByte
                    'tmpStr = tmpStr & "*) " & .ReadInt16 & " " & .ReadInt16 & vbCrLf
                   
                    Dim cantidadobjs As Integer

                    Dim obindex      As Integer
                   
                    cantidadobjs = Reader.ReadInt16
                    obindex = Reader.ReadInt16
                   
                    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1

           
                Next i

            End If

        Else
        
            QuestIndex = Reader.ReadInt16
        
            FrmQuests.titulo.Caption = QuestList(QuestIndex).nombre
           
            LevelRequerido = Reader.ReadInt8
            QuestRequerida = Reader.ReadInt16
           
            FrmQuests.detalle.Text = QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf

            If QuestRequerida <> 0 Then
                FrmQuests.detalle.Text = FrmQuests.detalle.Text & vbCrLf & "Quest: " & QuestList(QuestRequerida).nombre

            End If

           
            tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay NPCs

                For i = 1 To tmpByte
                    cantidadnpc = Reader.ReadInt16
                    NpcIndex = Reader.ReadInt16
               
                    Dim matados As Integer
               
                    matados = Reader.ReadInt16
                                     
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , NpcData(NpcIndex).Name)
                       
                    Dim cantok As Integer

                    cantok = cantidadnpc - matados
                       
                    If cantok = 0 Then
                        subelemento.SubItems(1) = "OK"
                    Else
                        subelemento.SubItems(1) = matados & "/" & cantidadnpc

                    End If
                        
                    ' subelemento.SubItems(1) = cantidadnpc - matados
                    subelemento.SubItems(2) = NpcIndex
                    subelemento.SubItems(3) = 0
                    'End If
                Next i

            End If
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay OBJs

                For i = 1 To tmpByte
               
                    cantidadobj = Reader.ReadInt16
                    ObjIndex = Reader.ReadInt16
                    
                    AmountHave = Reader.ReadInt16
                   
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , ObjData(ObjIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = ObjIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf

            Dim tmplong As Long
           
            tmplong = Reader.ReadInt32
           
            If tmplong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Oro")
                subelemento.SubItems(1) = BeautifyBigNumber(tmplong)
                subelemento.SubItems(2) = 12
                subelemento.SubItems(3) = 0

            End If
            
            tmplong = Reader.ReadInt32
           
            If tmplong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Experiencia")
                           
                subelemento.SubItems(1) = BeautifyBigNumber(tmplong)
                subelemento.SubItems(2) = 608
                subelemento.SubItems(3) = 1

            End If
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then

                For i = 1 To tmpByte
                    cantidadobjs = Reader.ReadInt16
                    obindex = Reader.ReadInt16
                   
                    Set subelemento = FrmQuests.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1

           
                Next i

            End If
        
        End If

    'Determinamos que formulario se muestra, seg�n si recibimos la informaci�n y la quest est� empezada o no.
    If QuestEmpezada Then
        FrmQuests.txtInfo.Text = tmpStr
        Call FrmQuests.ListView1_Click
        Call FrmQuests.ListView2_Click
        Call FrmQuests.lstQuests.SetFocus
    Else

        FrmQuestInfo.Show vbModeless, frmMain
        FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
        Call FrmQuestInfo.ListView1_Click
        Call FrmQuestInfo.ListView2_Click

    End If
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleQuestDetails", Erl)
    

End Sub
 
Public Sub HandleQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo ErrHandler
    
    Dim i       As Integer
    Dim tmpByte As Byte
    Dim tmpStr  As String
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = Reader.ReadInt8
    
    'Limpiamos el ListBox y el TextBox del formulario
    FrmQuests.lstQuests.Clear
    FrmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Reader.ReadString8
        
        'Agregamos los items
        For i = 1 To tmpByte
            FrmQuests.lstQuests.AddItem ReadField(i, tmpStr, 59)
        Next i

    End If
    
    'Mostramos el formulario
    
    COLOR_AZUL = RGB(0, 0, 0)
    Call Establecer_Borde(FrmQuests.lstQuests, FrmQuests, COLOR_AZUL, 0, 0)
    FrmQuests.Picture = LoadInterface("ventanadetallemision.bmp")
    FrmQuests.Show vbModeless, frmMain
    
    'Pedimos la informacion de la primer quest (si la hay)
    If tmpByte Then Call WriteQuestDetailsRequest(1)

    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleQuestListSend", Erl)
    

End Sub

Public Sub HandleNpcQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo ErrHandler

    Dim tmpStr         As String
    Dim tmpByte        As Byte
    Dim QuestEmpezada  As Boolean
    Dim i              As Integer
    Dim j              As Byte
    Dim cantidadnpc    As Integer
    Dim NpcIndex       As Integer
    Dim cantidadobj    As Integer
    Dim ObjIndex       As Integer
    Dim QuestIndex     As Integer
    Dim estado         As Byte
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    Dim CantidadQuest  As Byte
    Dim Repetible      As Boolean
    Dim subelemento    As ListItem
    
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear

        CantidadQuest = Reader.ReadInt8
            
        For j = 1 To CantidadQuest
        
            QuestIndex = Reader.ReadInt16
            
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
                              
            QuestList(QuestIndex).RequiredLevel = Reader.ReadInt8
            QuestList(QuestIndex).RequiredQuest = Reader.ReadInt16
            
            tmpByte = Reader.ReadInt8
    
            If tmpByte Then 'Hay NPCs
            
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
               
                End If
                    
                ReDim QuestList(QuestIndex).RequiredNPC(1 To tmpByte)
                    
                For i = 1 To tmpByte
                                                
                    QuestList(QuestIndex).RequiredNPC(i).Amount = Reader.ReadInt16
                    QuestList(QuestIndex).RequiredNPC(i).NpcIndex = Reader.ReadInt16

                Next i

            Else
                ReDim QuestList(QuestIndex).RequiredNPC(0)

            End If
               
            tmpByte = Reader.ReadInt8
    
            If tmpByte Then 'Hay OBJs
                ReDim QuestList(QuestIndex).RequiredOBJ(1 To tmpByte)
    
                For i = 1 To tmpByte
                   
                    QuestList(QuestIndex).RequiredOBJ(i).Amount = Reader.ReadInt16
                    QuestList(QuestIndex).RequiredOBJ(i).ObjIndex = Reader.ReadInt16

                Next i

            Else
                ReDim QuestList(QuestIndex).RequiredOBJ(0)
    
            End If
               
            QuestList(QuestIndex).RewardGLD = Reader.ReadInt32
            QuestList(QuestIndex).RewardEXP = Reader.ReadInt32

            tmpByte = Reader.ReadInt8
    
            If tmpByte Then
                
                ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)
    
                For i = 1 To tmpByte
                                              
                    QuestList(QuestIndex).RewardOBJ(i).Amount = Reader.ReadInt16
                    QuestList(QuestIndex).RewardOBJ(i).ObjIndex = Reader.ReadInt16
               
                Next i

            Else
                ReDim QuestList(QuestIndex).RewardOBJ(0)
    
            End If
                
            estado = Reader.ReadInt8
            Repetible = QuestList(QuestIndex).Repetible = 1
            
            Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre & IIf(Repetible, " (R)", ""))
            subelemento.SubItems(2) = QuestIndex
  
            Select Case estado
                
                Case 0
                    subelemento.SubItems(1) = "Disponible"
                    subelemento.ForeColor = vbWhite
                    subelemento.ListSubItems(1).ForeColor = vbWhite

                Case 1
                    subelemento.SubItems(1) = "En Curso"
                    subelemento.ForeColor = RGB(255, 175, 10)
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)

                Case 2
                    If Repetible Then
                        subelemento.SubItems(1) = "Repetible"
                        subelemento.ForeColor = RGB(180, 180, 180)
                        subelemento.ListSubItems(1).ForeColor = RGB(180, 180, 180)
                    Else
                        subelemento.SubItems(1) = "Finalizada"
                        subelemento.ForeColor = RGB(15, 140, 50)
                        subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
                    End If

                Case 3
                    subelemento.SubItems(1) = "No disponible"
                    subelemento.ForeColor = RGB(255, 10, 10)
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
            End Select
            
            FrmQuestInfo.ListViewQuest.Refresh
                
        Next j

    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest est� empezada o no.
    FrmQuestInfo.Show vbModeless, frmMain
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNpcQuestListSend", Erl)
    
    
End Sub

Private Sub HandleShowPregunta()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim msg As String

    PreguntaScreen = Reader.ReadString8()
    Pregunta = True
    
    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleShowPregunta", Erl)
    

End Sub

Private Sub HandleDatosGrupo()
    
    On Error GoTo HandleDatosGrupo_Err
    
    Dim EnGrupo      As Boolean

    Dim CantMiembros As Byte

    Dim i            As Byte
    
    EnGrupo = Reader.ReadBool()
    
    If EnGrupo Then
        CantMiembros = Reader.ReadInt8()

        For i = 1 To CantMiembros
            FrmGrupo.lstGrupo.AddItem (Reader.ReadString8)
        Next i

    End If
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    ' establece el borde al listbox
    Call Establecer_Borde(FrmGrupo.lstGrupo, FrmGrupo, COLOR_AZUL, 0, 0)

    FrmGrupo.Show , frmMain
    
    Exit Sub

HandleDatosGrupo_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleDatosGrupo", Erl)
    
    
End Sub

Private Sub HandleUbicacion()
    
    On Error GoTo HandleUbicacion_Err
    
    Dim miembro As Byte
    Dim x As Integer
    Dim y       As Byte
    Dim map     As Integer
    
    miembro = Reader.ReadInt8()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    map = Reader.ReadInt16()
    
    
    Exit Sub

HandleUbicacion_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUbicacion", Erl)
    
    
End Sub

Private Sub HandleViajarForm()
    
    On Error GoTo HandleViajarForm_Err
            
    Dim Dest     As String
    Dim DestCant As Byte
    Dim i        As Byte
    Dim tempdest As String

    FrmViajes.List1.Clear
    
    DestCant = Reader.ReadInt8()
        
    ReDim Destinos(1 To DestCant) As Tdestino
        
    For i = 1 To DestCant
        
        tempdest = Reader.ReadString8()
        
        Destinos(i).CityDest = ReadField(1, tempdest, Asc("-"))
        Destinos(i).costo = ReadField(2, tempdest, Asc("-"))
        FrmViajes.List1.AddItem ListaCiudades(Destinos(i).CityDest) & " - " & Destinos(i).costo & " monedas"

    Next i
        
    Call Establecer_Borde(FrmViajes.List1, FrmViajes, COLOR_AZUL, 0, 0)
         
    ViajarInterface = Reader.ReadInt8()
        
    FrmViajes.Picture = LoadInterface("viajes" & ViajarInterface & ".bmp")
        
    If ViajarInterface = 1 Then
        FrmViajes.Image1.Top = 4690
        FrmViajes.Image1.Left = 3810
    Else
        FrmViajes.Image1.Top = 4680
        FrmViajes.Image1.Left = 3840

    End If

    FrmViajes.Show , frmMain
    
    Exit Sub

HandleViajarForm_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleViajarForm", Erl)
    
    
End Sub

Private Sub HandleSeguroResu()
    
    'Get data and update form
    SeguroResuX = Reader.ReadBool()
    
    If SeguroResuX Then
        Call AddToConsole("Seguro de resurrecci�n activado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    Else
        Call AddToConsole("Seguro de resurrecci�n desactivado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-off.bmp")

    End If
    
End Sub

Private Sub HandleStopped()

    UserStopped = Reader.ReadBool()

End Sub

Private Sub HandleInvasionInfo()

    InvasionActual = Reader.ReadInt8
    InvasionPorcentajeVida = Reader.ReadInt8
    InvasionPorcentajeTiempo = Reader.ReadInt8
    
    frmMain.Evento.enabled = False
    frmMain.Evento.Interval = 0
    frmMain.Evento.Interval = 10000
    frmMain.Evento.enabled = True

End Sub

Private Sub HandleCommerceRecieveChatMessage()
    
    Dim message As String
    message = Reader.ReadString8
        
    Call AddToConsole(frmComerciarUsu.RecTxt, message, 255, 255, 255, 0, False, True, False)
    
End Sub

Private Sub HandleDoAnimation()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim charindex As Integer

    Dim TempInt   As Integer

    Dim headIndex As Integer

    charindex = Reader.ReadInt16()
    
    With charlist(charindex)
        .AnimatingBody = Reader.ReadInt16()
        .Body = BodyData(.AnimatingBody)
        'Start animation
        .Body.Walk(.Heading).started = FrameTime
        .Body.Walk(.Heading).Loops = 0
    End With
    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleDoAnimation", Erl)
    
    
End Sub

Private Sub HandleOpenCrafting()

    Dim TIPO As Byte
    TIPO = Reader.ReadInt8

    frmCrafteo.Picture = LoadInterface(TipoCrafteo(TIPO).Ventana)
    frmCrafteo.InventoryGrhIndex = TipoCrafteo(TIPO).Inventario
    frmCrafteo.TipoGrhIndex = TipoCrafteo(TIPO).Icono
    
    Dim i As Long
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmCrafteo.InvCraftUser.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))
        End With
    Next i
    
    For i = 1 To MAX_SLOTS_CRAFTEO
        Call frmCrafteo.InvCraftItems.ClearSlot(i)
    Next i

    Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)

    Call frmCrafteo.SetResult(0, 0, 0)

    Comerciando = True

    frmCrafteo.Show , frmMain

End Sub

Private Sub HandleCraftingItem()
    Dim Slot As Byte, ObjIndex As Integer
    Slot = Reader.ReadInt8
    ObjIndex = Reader.ReadInt16
    
    If ObjIndex <> 0 Then
        With ObjData(ObjIndex)
            Call frmCrafteo.InvCraftItems.SetItem(Slot, ObjIndex, 1, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, 0)
        End With
    Else
        Call frmCrafteo.InvCraftItems.ClearSlot(Slot)
    End If
    
End Sub

Private Sub HandleCraftingCatalyst()
    Dim ObjIndex As Integer, Amount As Integer, Porcentaje As Byte
    ObjIndex = Reader.ReadInt16
    Amount = Reader.ReadInt16
    Porcentaje = Reader.ReadInt8
    
    If ObjIndex <> 0 Then
        With ObjData(ObjIndex)
            Call frmCrafteo.InvCraftCatalyst.SetItem(1, ObjIndex, Amount, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, 0)
        End With
    Else
        Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)
    End If

    frmCrafteo.PorcentajeAcierto = Porcentaje
    
End Sub

Private Sub HandleCraftingResult()
    Dim ObjIndex As Integer
    ObjIndex = Reader.ReadInt16

    If ObjIndex > 0 Then
        Dim Porcentaje As Byte, Precio As Long
        Porcentaje = Reader.ReadInt8
        Precio = Reader.ReadInt32
        Call frmCrafteo.SetResult(ObjData(ObjIndex).GrhIndex, Porcentaje, Precio)
    Else
        Call frmCrafteo.SetResult(0, 0, 0)
    End If
End Sub

Private Sub HandleForceUpdate()
    On Error GoTo HandleCerrarleCliente_Err
    
    Call MsgBox("�Nueva versi�n disponible! Se abrir� el lanzador para que puedas actualizar.", vbOKOnly, "Argentum World")
    
    Shell App.path & "\..\..\Launcher\LauncherAOWorld.exe"
    
    EngineRun = False

    Call CloseClient
    
    Exit Sub

HandleCerrarleCliente_Err:
    Call RegistrarError(err.Number, err.Description, "Protocol.HandleCerrarleCliente", Erl)
    
End Sub

Public Sub HandleAnswerReset()
    On Error GoTo ErrHandler

    If MsgBox("�Est� seguro que desea resetear el personaje? Los items que no sean depositados se perder�n.", vbYesNo, "Resetear personaje") = vbYes Then
        Call WriteResetearPersonaje
    End If

    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleAnswerReset", Erl)
End Sub
Public Sub HandleUpdateBankGld()

    On Error GoTo ErrHandler
    
    Dim UserBoveOro As Long
        
    UserBoveOro = Reader.ReadInt32
    
    Call frmGoliath.UpdateBankGld(UserBoveOro)
    Exit Sub
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleUpdateBankGld", Erl)

End Sub

Public Sub HandlePelearConPezEspecial()
    On Error GoTo ErrHandler
    
    PosicionBarra = 1
    DireccionBarra = 1
    Dim i As Integer
    
    For i = 1 To MAX_INTENTOS
        intentosPesca(i) = 0
    Next i
    PescandoEspecial = True
    Call Sound.Sound_Play(55)
    ContadorIntentosPescaEspecial_Fallados = 0
    ContadorIntentosPescaEspecial_Acertados = 0
    startTimePezEspecial = GetTickCount()
    Call Char_Dialog_Set(UserCharIndex, "Oh! Creo que tengo un super pez en mi linea, intentare obtenerlo con la letra P", &H1FFFF, 200, 130)
    Exit Sub
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePelearConPezEspecial", Erl)
End Sub

Public Sub HandlePrivilegios()
    On Error GoTo ErrHandler
    
    EsGM = Reader.ReadBool
    If EsGM Then
        frmMain.panelGM.visible = True
        frmMain.createObj.visible = True
        frmMain.btnInvisible.visible = True
        frmMain.btnSpawn.visible = True
    Else
        frmMain.panelGM.visible = False
        frmMain.createObj.visible = False
        frmMain.btnInvisible.visible = False
        frmMain.btnSpawn.visible = False
    End If
    Exit Sub
ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandlePrivilegios", Erl)
End Sub

Public Sub HandleShopPjsInit()
    frmShopPjsAO20.Show , frmMain
End Sub
Public Sub HandleShopInit()
    
    Dim cant_obj_shop As Long, i As Long
    
    cant_obj_shop = Reader.ReadInt16
    
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
    
    ReDim ObjShop(1 To cant_obj_shop) As ObjDatas
    
    For i = 1 To cant_obj_shop
        ObjShop(i).objNum = Reader.ReadInt32
        ObjShop(i).Valor = Reader.ReadInt32
        ObjShop(i).Name = Reader.ReadString8
         
        Call frmShopAO20.lstItemShopFilter.AddItem(ObjShop(i).Name & " (Valor: " & ObjShop(i).Valor & ")", i - 1)
    Next i
    
    frmShopAO20.Show , frmMain
End Sub

Public Sub HandleUpdateShopClienteCredits()
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
End Sub

Public Sub HandleObjQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 29/08/2021 by HarThaoS
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    On Error GoTo ErrHandler

    Dim tmpStr         As String
    Dim tmpByte        As Byte
    Dim QuestEmpezada  As Boolean
    Dim i              As Integer
    Dim cantidadnpc    As Integer
    Dim NpcIndex       As Integer
    Dim cantidadobj    As Integer
    Dim ObjIndex       As Integer
    Dim QuestIndex     As Integer
    Dim estado         As Byte
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    Dim CantidadQuest  As Byte
    Dim Repetible      As Boolean
    Dim subelemento    As ListItem

    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear


    QuestIndex = Reader.ReadInt16

    FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre

    QuestList(QuestIndex).RequiredLevel = Reader.ReadInt8
    QuestList(QuestIndex).RequiredQuest = Reader.ReadInt16


    tmpByte = Reader.ReadInt8

    If tmpByte Then 'Hay NPCs

        If tmpByte > 5 Then
            FrmQuestInfo.ListView1.FlatScrollBar = False
        Else
            FrmQuestInfo.ListView1.FlatScrollBar = True

        End If

        ReDim QuestList(QuestIndex).RequiredNPC(1 To tmpByte)

        For i = 1 To tmpByte

            QuestList(QuestIndex).RequiredNPC(i).Amount = Reader.ReadInt16
            QuestList(QuestIndex).RequiredNPC(i).NpcIndex = Reader.ReadInt16

        Next i

    Else
        ReDim QuestList(QuestIndex).RequiredNPC(0)

    End If

    tmpByte = Reader.ReadInt8

    If tmpByte Then 'Hay OBJs
        ReDim QuestList(QuestIndex).RequiredOBJ(1 To tmpByte)

        For i = 1 To tmpByte

            QuestList(QuestIndex).RequiredOBJ(i).Amount = Reader.ReadInt16
            QuestList(QuestIndex).RequiredOBJ(i).ObjIndex = Reader.ReadInt16

        Next i

    Else
        ReDim QuestList(QuestIndex).RequiredOBJ(0)

    End If

    QuestList(QuestIndex).RewardGLD = Reader.ReadInt32
    QuestList(QuestIndex).RewardEXP = Reader.ReadInt32

    tmpByte = Reader.ReadInt8

    If tmpByte Then

        ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)

        For i = 1 To tmpByte

            QuestList(QuestIndex).RewardOBJ(i).Amount = Reader.ReadInt16
            QuestList(QuestIndex).RewardOBJ(i).ObjIndex = Reader.ReadInt16

        Next i

    Else
        ReDim QuestList(QuestIndex).RewardOBJ(0)

    End If

    estado = Reader.ReadInt8
    Repetible = QuestList(QuestIndex).Repetible = 1

    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre & IIf(Repetible, " (R)", ""))
    subelemento.SubItems(2) = QuestIndex

    Select Case estado

        Case 0
            subelemento.SubItems(1) = "Disponible"
            subelemento.ForeColor = vbWhite
            subelemento.ListSubItems(1).ForeColor = vbWhite

        Case 1
            subelemento.SubItems(1) = "En Curso"
            subelemento.ForeColor = RGB(255, 175, 10)
            subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)

        Case 2
            If Repetible Then
                subelemento.SubItems(1) = "Repetible"
                subelemento.ForeColor = RGB(180, 180, 180)
                subelemento.ListSubItems(1).ForeColor = RGB(180, 180, 180)
            Else
                subelemento.SubItems(1) = "Finalizada"
                subelemento.ForeColor = RGB(15, 140, 50)
                subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
            End If

        Case 3
            subelemento.SubItems(1) = "No disponible"
            subelemento.ForeColor = RGB(255, 10, 10)
            subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
    End Select

    FrmQuestInfo.ListViewQuest.Refresh

    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest est� empezada o no.
    FrmQuestInfo.Show vbModeless, frmMain
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)

    Exit Sub

ErrHandler:

    Call RegistrarError(err.Number, err.Description, "Protocol.HandleNpcQuestListSend", Erl)


End Sub
Public Sub HandleComboCooldown()

    ComboCooldownTime = Reader.ReadInt16
    StartComboCooldownTime = GetTickCount
End Sub
Public Sub HandleAccountCharacterList()

    CantidadDePersonajesEnCuenta = Reader.ReadInt

    Dim ii As Byte
    Dim privs As Integer
     'name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing
    For ii = 1 To MAX_PERSONAJES_EN_CUENTA
        Pjs(ii).nombre = ""
        Pjs(ii).Head = 0 ' si is_sailing o muerto, cabeza en 0
        Pjs(ii).Clase = 0
        Pjs(ii).Body = 0
        Pjs(ii).Mapa = 0
        Pjs(ii).PosX = 0
        Pjs(ii).PosY = 0
        Pjs(ii).nivel = 0
        Pjs(ii).Criminal = 0
        Pjs(ii).Casco = 0
        Pjs(ii).Escudo = 0
        Pjs(ii).Arma = 0
        Pjs(ii).ClanName = ""
        Pjs(ii).NameMapa = ""
    Next ii
    
    For ii = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
        Pjs(ii).nombre = Reader.ReadString8
        Pjs(ii).Body = Reader.ReadInt16
        Pjs(ii).Head = Reader.ReadInt16
        Pjs(ii).Clase = Reader.ReadInt16
        Pjs(ii).Mapa = Reader.ReadInt16
        Pjs(ii).PosX = Reader.ReadInt16
        Pjs(ii).PosY = Reader.ReadInt16
        Pjs(ii).nivel = Reader.ReadInt16
        Pjs(ii).Criminal = Reader.ReadInt8
        Pjs(ii).Casco = Reader.ReadInt16
        Pjs(ii).Escudo = Reader.ReadInt16
        Pjs(ii).Arma = Reader.ReadInt16
        privs = Reader.ReadInt8
        If privs > 1 Then Pjs(ii).priv = Log(privs) / Log(2)
        Pjs(ii).ClanName = "" ' "<" & "pepito" & ">"

    Next ii
    'Criminal = 0
    'Ciudadano = 1
    'caos = 2
    'armada = 3
    'concilio = 4
    'consejo = 5
    Dim i As Long
    For i = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
        If Pjs(i).priv > 1 Then
            Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(Pjs(i).priv).r, ColoresPJ(Pjs(i).priv).G, ColoresPJ(Pjs(i).priv).B)
        Else
            Select Case Pjs(i).Criminal
                Case 0 'Criminal
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B)
                    Pjs(i).priv = 0
                Case 1 'Ciudadano
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B)
                    Pjs(i).priv = 0
                Case 2 'Caos
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).B)
                    Pjs(i).priv = 0
                Case 3 'Armada
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).B)
                    Pjs(i).priv = 0
                Case 4 'Concilio
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).B)
                    Pjs(i).priv = 0
                Case 5 'Consejo
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).B)
                    Pjs(i).priv = 0
                Case Else
            End Select
        End If
    Next i
    
    
    AlphaRenderCuenta = MAX_ALPHA_RENDER_CUENTA
   
    If CantidadDePersonajesEnCuenta > 0 Then
        PJSeleccionado = 1
        LastPJSeleccionado = 1
        
        If Pjs(1).Mapa <> 0 Then
            Call SwitchMap(Pjs(1).Mapa)
            RenderCuenta_PosX = Pjs(1).PosX
            RenderCuenta_PosY = Pjs(1).PosY
        End If
    End If
     
    ' FrmCuenta.Show
    AlphaNiebla = 30

    frmConnect.visible = True
    QueRender = 2
    
    'UserMap = 323
    
    'Call SwitchMap(UserMap)
    
    SugerenciaAMostrar = RandomNumber(1, NumSug)
        
    ' LogeoAlgunaVez = True
    Call Sound.Sound_Play(192)
    
    Call Sound.Sound_Stop(SND_LLUVIAIN)
    '  Sound.NextMusic = 2
    '  Sound.Fading = 350
      
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
                
    If frmNewAccount.visible Then
        Unload frmNewAccount
    End If
    
    If FrmLogear.visible Then
        Unload FrmLogear

        'Unload frmConnect
    End If
    
    If frmMain.visible Then
        '  frmMain.Visible = False
        
        UserParalizado = False
        UserInmovilizado = False
        UserStopped = False
        
        InvasionActual = 0
        frmMain.Evento.enabled = False
     
        'BUG CLONES

        For i = 1 To LastChar
            Call EraseChar(i)
        Next i
        

    End If
End Sub