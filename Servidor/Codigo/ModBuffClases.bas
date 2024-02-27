Attribute VB_Name = "ModBuffClases"
Option Explicit
Public Enum tipoDanio
    t_golpe = 1
    t_magia
End Enum
Public combos As Dictionary

Public Sub DañoExtraPorCombo(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer, ByRef Daño As Long, ByVal tipo As tipoDanio)
    
    'Me fijo si ya está en un combo
    
    If GetTickCount - UserList(attackerIndex).Counters.buff_start_time < 3000 Then
        'Si Ya está me fijo si la víctima es su target del combo
        If UserList(attackerIndex).flags.victim_index_combo = VictimIndex And UserList(VictimIndex).flags.attacker_index_combo = attackerIndex Then
           
            Dim combo As clsBuffCombo
            
            If combos.Exists(UserList(attackerIndex).clase & "-" & UserList(VictimIndex).clase) Then
                
                Set combo = combos(UserList(attackerIndex).clase & "-" & UserList(VictimIndex).clase)
                
                Dim combea As Boolean
                
                Select Case tipo
                    Case tipoDanio.t_golpe
                        If combo.golpe Then combea = True
                    Case tipoDanio.t_magia
                        If combo.magia Then combea = True
                End Select
                
                If combea Then
                
                    UserList(attackerIndex).Counters.buff_combo_count = UserList(attackerIndex).Counters.buff_combo_count + 1
                    
                    If UserList(attackerIndex).Counters.buff_combo_count > combo.maxCombos Then
                        UserList(attackerIndex).Counters.buff_combo_count = combo.maxCombos
                    End If
                    
                    Daño = Daño + Daño * (combo.getBuffs(UserList(attackerIndex).Counters.buff_combo_count) / 100)
                    Call WriteSendComboCooldown(attackerIndex, 3000)
                End If
                
                
            End If
        Else
            UserList(attackerIndex).Counters.buff_combo_count = 0
        End If
    Else
        UserList(attackerIndex).Counters.buff_combo_count = 0
    End If
    
     UserList(attackerIndex).Counters.buff_start_time = GetTickCount
     UserList(attackerIndex).flags.victim_index_combo = VictimIndex
     UserList(VictimIndex).flags.attacker_index_combo = attackerIndex
End Sub

Private Function porecentajeExtra() As Byte
    
End Function

Public Sub LoadCombos()

    
    Set combos = New Dictionary
    
    Dim i As Byte
    Dim File As clsIniManager

    Set File = New clsIniManager
    Call File.Initialize(DatPath & "Buffs.ini")
    
    Dim cantCombos As Byte
        
    cantCombos = val(File.GetValue("INIT", "CantidadDeCombos"))
    
    Dim combo As clsBuffCombo
    If cantCombos > 0 Then
        
        For i = 1 To cantCombos
            Set combo = New clsBuffCombo
            combo.golpe = val(File.GetValue("Combo" & i, "Golpe"))
            combo.magia = val(File.GetValue("Combo" & i, "Magia"))
            
            Dim arrAtacantes() As String, arrVictimas() As String
            
            arrAtacantes = Split(File.GetValue("Combo" & i, "Atacantes"), "-")
            arrVictimas = Split(File.GetValue("Combo" & i, "Victimas"), "-")
                        
            Dim keys() As String
            Dim index As Byte
            Dim j As Byte
            Dim K As Byte
            
            index = 0
            
            For j = 0 To UBound(arrAtacantes)
                For K = 0 To UBound(arrVictimas)
                    index = index + 1
                    ReDim Preserve keys(1 To index) As String
                    keys(index) = classIdByName(arrAtacantes(j)) & "-" & classIdByName(arrVictimas(K))
                Next K
            Next j
            
            combo.maxCombos = val(File.GetValue("Combo" & i, "MaxCombos"))
            
            If combo.maxCombos > 0 Then
                For j = 1 To combo.maxCombos
                    Call combo.setBuffs(j, val(File.GetValue("Combo" & i, "Buff" & j)))
                Next j
            End If
            
            For j = 1 To UBound(keys)
                Call combos.Add(keys(j), combo)
            Next j
            
            
        Next i
    End If
End Sub

Private Function classIdByName(ByVal className As String) As Byte
    Select Case className
        Case "Mago"
            classIdByName = e_Class.Mage
        Case "Clerigo"
            classIdByName = e_Class.Cleric
        Case "Guerrero"
            classIdByName = e_Class.Warrior
        Case "Asesino"
            classIdByName = e_Class.Assasin
        Case "Bardo"
            classIdByName = e_Class.Bard
        Case "Druida"
            classIdByName = e_Class.Druid
        Case "Paladin"
            classIdByName = e_Class.Paladin
        Case "Cazador"
            classIdByName = e_Class.Hunter
        Case "Pirata"
            classIdByName = e_Class.Pirat
        Case "Ladron"
            classIdByName = e_Class.Thief
        Case "Bandido"
            classIdByName = e_Class.Bandit
    End Select
End Function



 
