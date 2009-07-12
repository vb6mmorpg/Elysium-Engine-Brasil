Attribute VB_Name = "modGameLogic"
' Copyright (c) 2008 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

Option Explicit

Sub AddToGrid(ByVal NewMap, _
   ByVal NewX, _
   ByVal NewY)
    Grid(NewMap).Loc(NewX, NewY).Blocked = True
End Sub

Sub AttackNpc(ByVal Attacker As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim N As Long, i As Long, x As Long, o As Long
    Dim MapNum As Long, NpcNum As Long

    ' Checar por RTE9
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Checar por armas
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        N = 0
    End If

    ' Enviar esta packet, então eles podem ver a pessoa atacando.
    Call SendDataToMap(GetPlayerMap(Attacker), "ATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & END_CHAR)
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(NpcNum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then

        ' Checar por uma arma e falar dano
        Call BattleMsg(Attacker, "Você matou um(a) " & Name, BrightRed, 0)
        Dim add As String

        add = 0

        If GetPlayerWeaponSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AddEXP
        End If

        If GetPlayerArmorSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerArmorSlot(Attacker))).AddEXP
        End If

        If GetPlayerShieldSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerShieldSlot(Attacker))).AddEXP
        End If

        If GetPlayerHelmetSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerHelmetSlot(Attacker))).AddEXP
        End If

        If add > 0 Then
            If add < 100 Then
                If add < 10 Then
                    add = 0 & ".0" & Right$(add, 2)
                Else
                    add = 0 & "." & Right$(add, 2)
                End If

            Else
                add = Mid$(add, 1, 1) & "." & Right$(add, 2)
            End If
        End If

        ' Calcular experiência dada ao atacante
        If add > 0 Then
            Exp = Npc(NpcNum).Exp + (Npc(NpcNum).Exp * Val(add))
        Else
            Exp = Npc(NpcNum).Exp
        End If

        ' Ter certeza que não dar experiência menor que 0.
        If Exp < 0 Then
            Exp = 1
        End If

        ' Checar se está em grupo, se sim, dividir experiência.
        If Player(Attacker).InParty = 0 Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "Você não pode mais ganhar experiência!", BrightBlue, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "Você ganhou " & Exp & " de experiência.", BrightBlue, 0)
            End If

        Else
            o = 0

            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(i) <> Attacker Then
                    If Party(Player(Attacker).PartyID).Member(i) <> 0 Then
                        If GetPlayerMap(Attacker) = GetPlayerMap(Party(Player(Attacker).PartyID).Member(i)) Then
                            o = o + 1
                        End If
                    End If
                End If

            Next

            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "Você não pode mais ganhar experiência!", BrightBlue, 0)
            Else

                If o <> 0 Then
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Int(Exp * 0.75))
                    Call BattleMsg(Attacker, "Você ganhou " & Int(Exp * 0.75) & " de experiência e dividiu " & Int(Exp * 0.25) & " com seu grupo.", BrightBlue, 0)
                Else
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                    Call BattleMsg(Attacker, "Você ganhou " & Exp & " de experiência, mas não pôde dividir com o seu grupo.", BrightBlue, 0)
                End If
            End If

            If o <> 0 Then

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Attacker).PartyID).Member(i) <> Attacker And Party(Player(Attacker).PartyID).Member(i) <> 0 Then
                        If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(i), Experience(MAX_LEVEL))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(i), "Você não pode mais ganhar experiência!", BrightBlue, 0)
                        Else
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(i), GetPlayerExp(Party(Player(Attacker).PartyID).Member(i)) + Int(Exp * (0.25 / o)))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(i), "Você ganhou " & Int(Exp * (0.25 / o)) & " de experiência do seu grupo.", BrightBlue, 0)
                        End If
                    End If

                Next

            End If
        End If

        For i = 1 To MAX_NPC_DROPS

            ' Drop the goods if they get it
            N = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
            End If

        Next

        ' Agora, setar HP para 0
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)

        ' Checar por level up!
        Call CheckPlayerLevelUp(Attacker)

        ' Checar para ver se algum membro do grupo ganhou level com a experiência
        If Player(Attacker).InParty = YES Then

            For x = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(x) <> 0 Then
                    Call CheckPlayerLevelUp(Party(Player(Attacker).PartyID).Member(x))
                End If

            Next

        End If

        Call TakeFromGrid(MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)

        ' Checar se o alvo que morreu era um NPC e setar alvo para 0
          If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If

    Else

        ' NPC não morreu, apenas fazer o dano
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage

        ' Checar por arma e mandar mensagem
        Call BattleMsg(Attacker, "Você atacou um(a) " & Name & " tirando " & Damage & " de dano.", White, 0)

        If N = 0 Then

            'Call PlayerMsg(Attacker, "Você atacou um(a) " & Name & " tirando " & Damage & " de dano.", White)
        Else

            'Call PlayerMsg(Attacker, "Você atacou um(a) " & Name & " com um(a) " & Trim$(Item(n).Name) & " tirando " & Damage & " de dano.", White)
        End If

        ' Checar se devemos mandar uma mensagem
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " : " & Trim$(Npc(NpcNum).AttackSay), SayColor)
            End If
        End If

        ' Setar o alvo do NPC para o jogador.
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_PLAYER

        ' Agora checar pela IA.
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then

            For i = 1 To MAX_MAP_NPCS

                If MapNpc(MapNum, i).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, i).Target = Attacker
                    MapNpc(MapNum, i).TargetType = TARGET_TYPE_PLAYER
                End If

            Next

        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
    ' Resetar o timer de ataque
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub AttackPlayer(ByVal Attacker As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
    Dim Exp As Long
    Dim N As Long
    Dim OldMap, oldx, oldy As Long

    ' Checar por Subscript out range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Checar por arma
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        N = 0
    End If

    ' Enviar esta packet para que se veja a pessoa atacando =P
    Call SendDataToMap(GetPlayerMap(Attacker), "ATTACKPLAYER" & SEP_CHAR & Attacker & SEP_CHAR & Victim & END_CHAR)

    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then

            ' Setar HP para nada
            Call SetPlayerHP(Victim, 0)

            ' Checar por arma e falar dano
            Call BattleMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " tirando " & Damage & " de dano.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " ataca você tirando " & Damage & " de dano.", BrightRed, 1)

            ' Jogador morto
            Call GlobalMsg(GetPlayerName(Victim) & " foi morto por " & GetPlayerName(Attacker), BrightRed)
            Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Morte" & END_CHAR)

            If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Principal.txt", "DropItems " & Victim
                Else

                    If GetPlayerWeaponSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                    End If

                    If GetPlayerArmorSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                    End If

                    If GetPlayerHelmetSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                    End If

                    If GetPlayerShieldSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                    End If
                End If

                ' Calcular Exp para o atacante
                Exp = Int(GetPlayerExp(Victim) / 10)

                ' Ter certeza que não é menor que 0
                If Exp < 0 Then
                    Exp = 0
                End If

                If GetPlayerLevel(Victim) = MAX_LEVEL Then
                    Call BattleMsg(Victim, "Você não pode perder experiência!", BrightRed, 1)
                    Call BattleMsg(Attacker, GetPlayerName(Victim) & " possui o nível máximo!", BrightBlue, 0)
                Else

                    If Exp = 0 Then
                        Call BattleMsg(Victim, "Você não perdeu experiência.", BrightRed, 1)
                        Call BattleMsg(Attacker, "Você não recebeu experiência.", BrightBlue, 0)
                    Else
                        Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                        Call BattleMsg(Victim, "Você perdeu " & Exp & " de experiência.", BrightRed, 1)
                        Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                        Call BattleMsg(Attacker, "Você conseguiu " & Exp & " de experiência por matar " & GetPlayerName(Victim) & ".", BrightBlue, 0)
                    End If
                End If
            End If

            OldMap = GetPlayerMap(Victim)
            oldx = GetPlayerX(Victim)
            oldy = GetPlayerY(Victim)

            ' Teleportar jogador
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Principal.txt", "OnDeath " & Victim
            Else
                Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
            End If

            Call UpdateGrid(OldMap, oldx, oldy, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim))

            ' Restorar vitais
            Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
            Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
            Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
            Call SendHP(Victim)
            Call SendMP(Victim)
            Call SendSP(Victim)

            ' Checar por um level up!
            Call CheckPlayerLevelUp(Attacker)

            ' Checar se o alvo é o player que morreu e setar para 0Check
            If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
                Player(Attacker).Target = 0
                Player(Attacker).TargetType = 0
            End If

            If GetPlayerPK(Victim) = NO Then
                If GetPlayerPK(Attacker) = NO Then
                    Call SetPlayerPK(Attacker, YES)
                    Call SendPlayerData(Attacker)
                    Call GlobalMsg(GetPlayerName(Attacker) & " se declarou um Player Killer!", BrightRed)
                End If

            Else
                Call SetPlayerPK(Victim, NO)
                Call SendPlayerData(Victim)
                Call GlobalMsg(GetPlayerName(Victim) & " pagou o preço por ser um Player Killer!", BrightRed)
            End If

        Else

            ' Jogador não morreu, apenas dano
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)

            ' Checar por arma e falar dano
            Call BattleMsg(Attacker, "Você atacou " & GetPlayerName(Victim) & " tirando " & Damage & " de dano.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " ataca você retirando " & Damage & " de dano.", BrightRed, 1)

            If N = 0 Then

                'Call PlayerMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " tirando " & Damage & " de dano.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " ataca você retirando " & Damage & " de dano.", BrightRed)
            Else

                'Call PlayerMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " com um(a) " & Trim$(Item(n).Name) & " retirando " & Damage & " de dano.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " ataca você com um(a) " & Trim$(Item(n).Name) & " retirando " & Damage & " de dano.", BrightRed)
            End If
        End If

    ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then

        If Damage >= GetPlayerHP(Victim) Then

            ' Setar HP para 0
            Call SetPlayerHP(Victim, 0)

            ' Checar por arma e falar dano
            Call BattleMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " tirando " & Damage & " de dano.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " ataca você retirando " & Damage & " de dano.", BrightRed, 1)

            If N = 0 Then

                'Call PlayerMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " retirando " & Damage & " de dano.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " ataca você retirando " & Damage & " de dano.", BrightRed)
            Else

                'Call PlayerMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " com um(a) " & Trim$(Item(n).Name) & " retirando " & Damage & " de dano.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " ataca você com um(a) " & Trim$(Item(n).Name) & " retirando " & Damage & " de dano.", BrightRed)
            End If

            ' Jogador morto (tô chorando)
            Call GlobalMsg(GetPlayerName(Victim) & " foi morto na arena por " & GetPlayerName(Attacker), BrightRed)
            Call UpdateGrid(GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim), Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data1, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data2, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data3)

            ' Teleportar
            Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data1, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data2, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data3)

            ' Restorar vitais
            Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
            Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
            Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
            Call SendHP(Victim)
            Call SendMP(Victim)
            Call SendSP(Victim)

            ' Checar se o alvo foi o jogador que morreu, e setar para 0.
            If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
                Player(Attacker).Target = 0
                Player(Attacker).TargetType = 0
            End If

        Else

            ' Jogador não morreu, apenas dano. (eba!)
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)

            ' Checar por arma e falar dano.
            Call BattleMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " retirando " & Damage & " de dano.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " ataca você retirando " & Damage & " de dano.", BrightRed, 1)

            If N = 0 Then

                'Call PlayerMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " retirando " & Damage & " de dano.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " ataca você retirando " & Damage & " de dano.", BrightRed)
            Else

                'Call PlayerMsg(Attacker, "Você ataca " & GetPlayerName(Victim) & " com um(a) " & Trim$(Item(n).Name) & " retirando " & Damage & " de dano.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " ataca você com um(a) " & Trim$(Item(n).Name) & " retirando " & Damage & " de dano.", BrightRed)
            End If
        End If
    End If

    ' Resetar timer de ataque
    Player(Attacker).AttackTimer = GetTickCount
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Dor" & END_CHAR)
End Sub

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim AttackSpeed As Long
    Dim x As Long
    Dim y As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
    End If

    CanAttackNpc = False

    ' Checar por subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Checar por subscript out of range (de novo? aff)
    If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Ter certeza que o npc não morreu
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Ter certeza que estão no mesmo mapa
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then

            ' Check if at same coordinates
            x = DirToX(GetPlayerX(Attacker), GetPlayerDir(Attacker))
            y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

            If (MapNpc(MapNum, MapNpcNum).y = y) And (MapNpc(MapNum, MapNpcNum).x = x) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanAttackNpc = True
                Else

                    If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                        Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " : " & Trim$(Npc(NpcNum).AttackSay), Green)
                    End If

                    If Npc(NpcNum).Speech <> 0 Then
                        Call SendDataTo(Attacker, "STARTSPEECH" & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR & 0 & SEP_CHAR & NpcNum & END_CHAR)
                    End If
                End If
            End If
        End If
    End If

End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim AttackSpeed As Long
    Dim Dir As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
    End If

    CanAttackNpcWithArrow = False

    ' Checar por subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Checar por subscript out of range
    If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Ter certeza que o NPC não morreu
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Ter certeza que estão no mesmo mapa
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
            If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                CanAttackNpcWithArrow = True
            Else

                If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                    Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " : " & Trim$(Npc(NpcNum).AttackSay), Green)
                End If

                If Npc(NpcNum).Speech <> 0 Then

                    For Dir = 0 To 3

                        If DirToX(GetPlayerX(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).x And DirToY(GetPlayerY(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).y Then
                            Call SendDataTo(Attacker, "STARTSPEECH" & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR & 0 & SEP_CHAR & NpcNum & END_CHAR)
                        End If

                    Next Dir

                End If
            End If
        End If
    End If

End Function

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim AttackSpeed As Long
    Dim x As Long
    Dim y As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
    End If

    CanAttackPlayer = False

    ' Checar por Subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Function
    End If

    ' Ter certeza que não tem 0 de HP
    If GetPlayerHP(Victim) <= 0 Then
        Exit Function
    End If

    ' Ter certeza que não estamos atacando enquanto ele troca de mapa
    If Player(Victim).GettingMap = YES Then
        Exit Function
    End If

    ' Ter certeza que estão no mesmo mapa
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(Attacker).AttackTimer + AttackSpeed) Then
        x = DirToX(GetPlayerX(Attacker), GetPlayerDir(Attacker))
        y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

        If (GetPlayerY(Victim) = y) And (GetPlayerX(Victim) = x) Then
            If Map(GetPlayerMap(Victim)).Tile(x, y).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then

                ' Ter certeza que eles não tem acesso
                If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                    Call PlayerMsg(Attacker, "Você não pode atacar um jogador sendo um administrador!", BrightBlue)
                Else

                    ' Checar se a vitima não é um administrador
                    If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "Você não pode atacar " & GetPlayerName(Victim) & "!", BrightRed)
                    Else

                        ' Checar se o mapa é atacavel
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then

                            ' Ter certeza que se possui level suficiente
                            If GetPlayerLevel(Attacker) < 10 Then
                                Call PlayerMsg(Attacker, "Você está abaixo do nível 10, portanto, você não pode atacar um jogador!", BrightRed)
                            Else

                                If GetPlayerLevel(Victim) < 10 Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " está abaixo do nível 10, portanto não pode ser atacado!", BrightRed)
                                Else

                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "Você não pode atacar um jogador do seu clã!", BrightRed)
                                        End If

                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If

                        Else
                            Call PlayerMsg(Attacker, "Esta é uma zona segura!", BrightRed)
                        End If
                    End If
                End If

            ElseIf Map(GetPlayerMap(Victim)).Tile(x, y).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                CanAttackPlayer = True
            End If
        End If
    End If

End Function

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    CanAttackPlayerWithArrow = False

    ' Checar por subscript of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Function
    End If

    ' Ter certeza que não se tem menos de 0 HP
    If GetPlayerHP(Victim) <= 0 Then
        Exit Function
    End If

    ' Ter certeza que não estão atacando o jogador se ele está trocando de mapas
    If Player(Victim).GettingMap = YES Then
        Exit Function
    End If

    ' Ter certeza que estão no mesmo mapa.
    If GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then
        If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then

            ' Ter certeza quanto ao acesso
            If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                Call PlayerMsg(Attacker, "Você não pode atacar um jogador sendo um administrador!", BrightBlue)
            Else

                ' Check to make sure the victim isn't an admin
                If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                    Call PlayerMsg(Attacker, "Você não pode atacar " & GetPlayerName(Victim) & "!", BrightRed)
                Else

                    ' Check if map is attackable
                    If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then

                        ' Make sure they are high enough level
                        If GetPlayerLevel(Attacker) < 10 Then
                            Call PlayerMsg(Attacker, "Você está abaixo do nível 10, portanto, você não pode atacar um jogador!", BrightRed)
                        Else

                            If GetPlayerLevel(Victim) < 10 Then
                                Call PlayerMsg(Attacker, GetPlayerName(Victim) & " está abaixo do nível 10, portanto não pode ser atacado!", BrightRed)
                            Else

                                If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                    If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                        CanAttackPlayerWithArrow = True
                                    Else
                                        Call PlayerMsg(Attacker, "Você não pode atacar um jogador do seu clã!", BrightRed)
                                    End If

                                Else
                                    CanAttackPlayerWithArrow = True
                                End If
                            End If
                        End If

                    Else
                        Call PlayerMsg(Attacker, "Esta é uma zona segura!", BrightRed)
                    End If
                End If
            End If

        ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
            CanAttackPlayerWithArrow = True
        End If
    End If

End Function

Function CanNpcAttackPet(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim x As Long
    Dim y As Long

    CanNpcAttackPet = False

    ' Checar por subscript of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then
        Exit Function
    End If

    ' Checar por subscript of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).num <= 0 Then
        Exit Function
    End If

    MapNum = Player(Index).Pet.Map
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Ter certeza que o NPC morreu
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Ter certeza que os npcs não vão atacar mais de uma vez por segundo
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount

    ' Ter certeza que se está no mesmo mapa
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            x = DirToX(MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).Dir)
            y = DirToY(MapNpc(MapNum, MapNpcNum).y, MapNpc(MapNum, MapNpcNum).Dir)

            ' Checar as coordenadas
            If (Player(Index).Pet.y = y) And (Player(Index).Pet.x = x) Then
                CanNpcAttackPet = True
            End If
        End If
    End If

End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim x As Long
    Dim y As Long

    CanNpcAttackPlayer = False

    ' Checar por subscript of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then
        Exit Function
    End If

    ' Checar por subscript of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Ter certeza que o NPC morreu
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Ter certeza que os npcs não vão atacar mais de uma vez por segundo
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Ter certeza que não se está trocando os mapas
    If Player(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount

    ' Ter certeza que está no mesmo mapa
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            x = DirToX(MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).Dir)
            y = DirToY(MapNpc(MapNum, MapNpcNum).y, MapNpc(MapNum, MapNpcNum).Dir)

            ' Checar as coordenadas
            If (GetPlayerY(Index) = y) And (GetPlayerX(Index) = x) Then
                CanNpcAttackPlayer = True
            End If
        End If
    End If

End Function

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim x As Long, y As Long

    CanNpcMove = False

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    x = DirToX(MapNpc(MapNum, MapNpcNum).x, Dir)
    y = DirToY(MapNpc(MapNum, MapNpcNum).y, Dir)

    If Not IsValid(x, y) Then Exit Function
    If Grid(MapNum).Loc(x, y).Blocked = True Then Exit Function
    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE And Map(MapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then Exit Function
    CanNpcMove = True
End Function

Function CanPetAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim x As Long
    Dim y As Long
    Dim Dir As Long

    CanPetAttackNpc = False

    ' Checar por Subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Checar por Subscript out of range
    If MapNpc(Player(Attacker).Pet.Map, MapNpcNum).num <= 0 Then
        Exit Function
    End If

    MapNum = Player(Attacker).Pet.Map
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Ter certeza que o NPC não morreu
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Ter certeza que estão no mesmo mapa
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > Player(Attacker).Pet.AttackTimer + 1000 Then
            If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then

                For Dir = 0 To 3

                    ' Checar as coordenadas
                    x = DirToX(Player(Attacker).Pet.x, Dir)
                    y = DirToY(Player(Attacker).Pet.y, Dir)

                    If (MapNpc(MapNum, MapNpcNum).y = y) And (MapNpc(MapNum, MapNpcNum).x = x) Then
                        CanPetAttackNpc = True
                    End If

                Next

            End If
        End If
    End If

End Function

Function CanPetMove(ByVal PetNum As Long, ByVal Dir) As Boolean
    Dim x As Long, y As Long
    Dim i As Long, Packet As String

    CanPetMove = False

    If PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    x = DirToX(Player(PetNum).Pet.x, Dir)
    y = DirToY(Player(PetNum).Pet.y, Dir)

    If Not IsValid(x, y) Then
        If Dir = DIR_UP Then
            If Map(Player(PetNum).Pet.Map).Up > 0 And Map(Player(PetNum).Pet.Map).Up = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
        End If

        If Dir = DIR_DOWN Then
            If Map(Player(PetNum).Pet.Map).Down > 0 And Map(Player(PetNum).Pet.Map).Down = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
        End If

        If Dir = DIR_LEFT Then
            If Map(Player(PetNum).Pet.Map).Left > 0 And Map(Player(PetNum).Pet.Map).Left = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
        End If

        If Dir = DIR_RIGHT Then
            If Map(Player(PetNum).Pet.Map).Right > 0 And Map(Player(PetNum).Pet.Map).Right = Player(PetNum).Pet.MapToGo Then

                'i = Player(PetNum).Pet.Map
                'Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
                'Packet = "PETDATA" & SEP_CHAR
                'Packet = Packet & PetNum & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Alive & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Map & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.x & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.y & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Dir & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Sprite & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.HP & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Level * 5 & SEP_CHAR
                'Packet = Packet & END_CHAR
                'Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
                'Call SendDataToMap(i, Packet)
                CanPetMove = True
            End If
        End If

        Exit Function
    End If

    If Grid(Player(PetNum).Pet.Map).Loc(x, y).Blocked = True Then Exit Function
    CanPetMove = True
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long, N As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    ShieldSlot = GetPlayerShieldSlot(Index)

    If ShieldSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    Dim i As Long, N As Long

    CanPlayerCriticalHit = False

    If GetPlayerWeaponSlot(Index) > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerstr(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Sub CastSpell(ByVal Index As Long, _
   ByVal SpellSlot As Long)
    Dim SpellNum As Long, i As Long, N As Long, Damage As Long
    Dim Casted As Boolean
    Dim x As Long, y As Long
    Dim Packet As String

    Casted = False
    Call SendPlayerXY(Index)

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    SpellNum = GetPlayerSpell(Index, SpellSlot)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call BattleMsg(Index, "Você não possui esta magia!", BrightRed, 0)
        Exit Sub
    End If

    i = GetSpellReqLevel(SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(Index, "Sem mana para usar a magia!", BrightRed, 0)
        Exit Sub
    End If

    ' Make sure they are the right level
    If i > GetPlayerLevel(Index) Then
        Call BattleMsg(Index, "Você precisa ser nível " & i & " para usar esta magia.", BrightRed, 0)
        Exit Sub
    End If

    ' Checar se o timer está normal
    If GetTickCount < Player(Index).AttackTimer + 1000 Then
        Exit Sub
    End If

    ' Que se foda isso XD
    'If Spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
    '
    '    N = FindOpenInvSlot(Index, Spell(SpellNum).Data1)
    '    If N > 0 Then
    '
    '        Call GiveItem(Index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
    '        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & ".", BrightBlue)
    '        ' Take away the mana points
    '        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
    '        Call SendMP(Index)
    '        Casted = True
    '
    '    Else
    '
    '        Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
    '
    '    End If
    '    Exit Sub
    'End If
    ' Checar se a magia é um summon
    If Spell(SpellNum).Type = SPELL_TYPE_PET Then
        Player(Index).Pet.Alive = YES
        Player(Index).Pet.Sprite = Spell(SpellNum).Data1
        Player(Index).Pet.Dir = DIR_UP
        Player(Index).Pet.Map = GetPlayerMap(Index)
        Player(Index).Pet.MapToGo = 0
        Player(Index).Pet.x = GetPlayerX(Index) + Int(Rnd * 3 - 1)

        If Player(Index).Pet.x < 0 Or Player(Index).Pet.x > MAX_MAPX Then Player(Index).Pet.x = GetPlayerX(Index)
        Player(Index).Pet.XToGo = -1
        Player(Index).Pet.y = GetPlayerY(Index) + Int(Rnd * 3 - 1)

        If Player(Index).Pet.y < 0 Or Player(Index).Pet.y > MAX_MAPY Then Player(Index).Pet.y = GetPlayerY(Index)
        Player(Index).Pet.YToGo = -1
        Player(Index).Pet.Level = Spell(SpellNum).Range
        Player(Index).Pet.HP = Player(Index).Pet.Level * 5
        Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.x & SEP_CHAR
        Packet = Packet & Player(Index).Pet.y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR

        ' Bagunça aqui XD
        Call PlayerMsg(Index, "Você conjura uma fera!", White)
        Call SendDataToMap(GetPlayerMap(Index), Packet)
        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
        Call SendMP(Index)
        Casted = True
        Exit Sub
    End If

    If Spell(SpellNum).AE = 1 Then

        For y = GetPlayerY(Index) - Spell(SpellNum).Range To GetPlayerY(Index) + Spell(SpellNum).Range
            For x = GetPlayerX(Index) - Spell(SpellNum).Range To GetPlayerX(Index) + Spell(SpellNum).Range
                N = -1

                If IsValid(x, y) Then

                    For i = 1 To MAX_PLAYERS

                        If IsPlaying(i) = True Then
                            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                                If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                                    If i = Index Then
                                        If Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
                                            Player(Index).Target = i
                                            Player(Index).TargetType = TARGET_TYPE_PLAYER
                                            N = Player(Index).Target
                                        End If

                                    Else
                                        Player(Index).Target = i
                                        Player(Index).TargetType = TARGET_TYPE_PLAYER
                                        N = Player(Index).Target
                                    End If
                                End If
                            End If
                        End If

                    Next

                    For i = 1 To MAX_MAP_NPCS

                        If MapNpc(GetPlayerMap(Index), i).num > 0 Then
                            If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then
                                Player(Index).Target = i
                                Player(Index).TargetType = TARGET_TYPE_NPC
                                N = Player(Index).Target
                            End If
                        End If

                    Next

                    If N < 0 Then
                        Player(Index).Target = MakeLoc(x, y)
                        Player(Index).TargetType = TARGET_TYPE_LOCATION
                        N = MakeLoc(x, y)
                    End If

                    Casted = False

                    If N > 0 Then
                        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                            If IsPlaying(N) Then
                                If N <> Index Then
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER

                                    If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then

                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_SUBHP
                                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(N) + (Rnd * 5) - 2

                                                If Damage > 0 Then
                                                    Call AttackPlayer(Index, N, Damage)
                                                Else
                                                    Call BattleMsg(Index, "A magia foi muita fraca e não feriu! " & GetPlayerName(N) & "!", BrightRed, 0)
                                                End If

                                            Case SPELL_TYPE_SUBMP
                                                Call SetPlayerMP(N, GetPlayerMP(N) - Spell(SpellNum).Data1)
                                                Call SendMP(N)

                                            Case SPELL_TYPE_SUBSP
                                                Call SetPlayerSP(N, GetPlayerSP(N) - Spell(SpellNum).Data1)
                                                Call SendSP(N)
                                        End Select

                                        Casted = True
                                    Else

                                        If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then

                                            Select Case Spell(SpellNum).Type

                                                Case SPELL_TYPE_ADDHP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerHP(N, GetPlayerHP(N) + Spell(SpellNum).Data1)
                                                    Call SendHP(N)

                                                Case SPELL_TYPE_ADDMP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerMP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)

                                                Case SPELL_TYPE_ADDSP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerSP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)
                                            End Select

                                            Casted = True
                                        Else
                                            Call PlayerMsg(Index, "Não se pôde usar essa magia!", BrightRed)
                                        End If
                                    End If

                                Else
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER

                                    If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then
                                    Else

                                        If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then

                                            Select Case Spell(SpellNum).Type

                                                Case SPELL_TYPE_ADDHP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerHP(N, GetPlayerHP(N) + Spell(SpellNum).Data1)
                                                    Call SendHP(N)

                                                Case SPELL_TYPE_ADDMP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerMP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)

                                                Case SPELL_TYPE_ADDSP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerSP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)
                                            End Select

                                            Casted = True
                                        Else
                                            Call BattleMsg(Index, "Não se pôde usar esta magia!", BrightRed, 0)
                                        End If
                                    End If
                                End If

                            Else
                                Call BattleMsg(Index, "Não se pôde usar esta magia!", BrightRed, 0)
                            End If

                        Else

                            If Player(Index).TargetType = TARGET_TYPE_NPC Then
                                If Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                                    If Spell(SpellNum).Type >= SPELL_TYPE_SUBHP And Spell(SpellNum).Type <= SPELL_TYPE_SUBSP Then

                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em um(a) " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_SUBHP
                                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), N).num).DEF / 2) + (Rnd * 5) - 2

                                                If Damage > 0 Then
                                                    Call AttackNpc(Index, N, Damage)
                                                Else
                                                    Call BattleMsg(Index, "A magia foi muito fraca para machucar " & Trim$(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightRed, 0)
                                                End If

                                            Case SPELL_TYPE_SUBMP
                                                MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP - Spell(SpellNum).Data1

                                            Case SPELL_TYPE_SUBSP
                                                MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP - Spell(SpellNum).Data1
                                        End Select

                                        Casted = True
                                    Else

                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_ADDHP

                                                'MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1
                                            Case SPELL_TYPE_ADDMP

                                                'MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1
                                            Case SPELL_TYPE_ADDSP

                                                'MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                                        End Select

                                        Casted = False
                                    End If

                                Else
                                    Call BattleMsg(Index, "Não se pôde usar esta magia!", BrightRed, 0)
                                End If

                            Else
                                Player(Index).TargetType = TARGET_TYPE_LOCATION
                                Casted = True
                            End If
                        End If
                    End If

                    If Casted = True Then
                        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & END_CHAR)

                        'Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & Spell(SpellNum).Sound & END_CHAR)
                    End If
                End If

            Next
        Next

        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
        Call SendMP(Index)
    Else
        N = Player(Index).Target

        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
            If IsPlaying(N) Then
                If GetPlayerName(N) <> GetPlayerName(Index) Then
                    If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(N)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(N)) ^ 2))) > Spell(SpellNum).Range Then
                        Call BattleMsg(Index, "Você está muito distante do alvo para atingi-lo.", BrightRed, 0)
                        Exit Sub
                    End If
                End If

                Player(Index).TargetType = TARGET_TYPE_PLAYER

                If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then

                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em " & GetPlayerName(n) & ".", BrightBlue)
                    Select Case Spell(SpellNum).Type

                        Case SPELL_TYPE_SUBHP
                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(N) + (Rnd * 5) - 2

                            If Damage > 0 Then
                                Call AttackPlayer(Index, N, Damage)
                            Else
                                Call BattleMsg(Index, "A magia foi muito graca para ferir " & GetPlayerName(N) & "!", BrightRed, 0)
                            End If

                        Case SPELL_TYPE_SUBMP
                            Call SetPlayerMP(N, GetPlayerMP(N) - Spell(SpellNum).Data1)
                            Call SendMP(N)

                        Case SPELL_TYPE_SUBSP
                            Call SetPlayerSP(N, GetPlayerSP(N) - Spell(SpellNum).Data1)
                            Call SendSP(N)
                    End Select

                    ' Retirar os MPs
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else

                    If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then

                        Select Case Spell(SpellNum).Type

                            Case SPELL_TYPE_ADDHP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerHP(N, GetPlayerHP(N) + Spell(SpellNum).Data1)
                                Call SendHP(N)

                            Case SPELL_TYPE_ADDMP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerMP(N) + Spell(SpellNum).Data1)
                                Call SendMP(N)

                            Case SPELL_TYPE_ADDSP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerSP(N) + Spell(SpellNum).Data1)
                                Call SendMP(N)
                        End Select

                        ' Retirar MPs
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                        Call SendMP(Index)
                        Casted = True
                    Else
                        Call BattleMsg(Index, "Não se pôde usar esta magia!", BrightRed, 0)
                    End If
                End If

            Else
                Call PlayerMsg(Index, "Não se pôde usar esta magia!", BrightRed)
            End If

        Else

            If CInt(Sqr((GetPlayerX(Index) - MapNpc(GetPlayerMap(Index), N).x) ^ 2 + ((GetPlayerY(Index) - MapNpc(GetPlayerMap(Index), N).y) ^ 2))) > Spell(SpellNum).Range Then
                Call BattleMsg(Index, "Você está muito distante do alvo para atingi-lo.", BrightRed, 0)
                Exit Sub
            End If

            Player(Index).TargetType = TARGET_TYPE_NPC

            If Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then

                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " usa " & Trim$(Spell(SpellNum).Name) & " em um(a) " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                Select Case Spell(SpellNum).Type

                    Case SPELL_TYPE_ADDHP
                        MapNpc(GetPlayerMap(Index), N).HP = MapNpc(GetPlayerMap(Index), N).HP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBHP
                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), N).num).DEF / 2 + (Rnd * 5) - 2)

                        If Damage > 0 Then
                            Call AttackNpc(Index, N, Damage)
                        Else
                            Call BattleMsg(Index, "A magia foi muito fraca para ferir " & Trim$(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightRed, 0)
                        End If

                    Case SPELL_TYPE_ADDMP
                        MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBMP
                        MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP - Spell(SpellNum).Data1

                    Case SPELL_TYPE_ADDSP
                        MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBSP
                        MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP - Spell(SpellNum).Data1
                End Select

                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                Call SendMP(Index)
                Casted = True
            Else
                Call BattleMsg(Index, "Não se pôde usar esta magia!", BrightRed, 0)
            End If
        End If
    End If

    If Casted = True Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & END_CHAR)

        If Spell(SpellNum).sound > 0 Then Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Magic" & Spell(SpellNum).sound & END_CHAR)
    End If

End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long, ItemNum As Long

    ' Vamos checar se um admin pega um objeto e blablabla
    Slot = GetPlayerWeaponSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                Call SetPlayerWeaponSlot(Index, 0)
            End If

        Else
            Call SetPlayerWeaponSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
                Call SetPlayerArmorSlot(Index, 0)
            End If

        Else
            Call SetPlayerArmorSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
                Call SetPlayerHelmetSlot(Index, 0)
            End If

        Else
            Call SetPlayerHelmetSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(Index, 0)
            End If

        Else
            Call SetPlayerShieldSlot(Index, 0)
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim C As Long

    C = 0

    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        If GetPlayerLevel(Index) < MAX_LEVEL Then
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Principal.txt", "PlayerLevelUp " & Index
            Else

                Do Until GetPlayerExp(Index) < GetPlayerNextLevel(Index)
                    DoEvents

                    If GetPlayerLevel(Index) < MAX_LEVEL Then
                        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
                            d = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
                            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                            i = Int(GetPlayerSPEED(Index) / 10)

                            If i < 1 Then i = 1
                            If i > 3 Then i = 3
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "LevelUp" & END_CHAR)
                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
                            Call SetPlayerExp(Index, d)
                            C = C + 1
                        Else
                            Exit Do
                        End If
                    End If

                Loop

                If C > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " ganhou " & C & " níveis!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " ganhou um nível!", 6)
                End If

                Call BattleMsg(Index, "Você possui " & GetPlayerPOINTS(Index) & " pontos.", 9, 0)
            End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & END_CHAR)
        End If

        If GetPlayerLevel(Index) = MAX_LEVEL Then
            Call SetPlayerExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
End Sub

' Vamos jogar fora os statments antigos?

Public Function DirToX(ByVal x As Long, _
   ByVal Dir As Byte) As Long
    DirToX = x

    If Dir = DIR_UP Or Dir = DIR_DOWN Then Exit Function

    ' LEFT = 2, RIGHT = 3
    ' 2 * 2 = 4, 4 - 5 = -1
    ' 3 * 2 = 6, 6 - 5 = 1
    DirToX = x + ((Dir * 2) - 5)
End Function

Public Function DirToY(ByVal y As Long, _
   ByVal Dir As Byte) As Long
    DirToY = y

    If Dir = DIR_LEFT Or Dir = DIR_RIGHT Then Exit Function

    ' UP = 0, DOWN = 1
    ' 0 * 2 = 0, 0 - 1 = -1
    ' 1 * 2 = 2, 2 - 1 = 1
    DirToY = y + ((Dir * 2) - 1)
End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    FindOpenInvSlot = 0

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' ¬¬
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Tentar achar um slot aberto
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long

    FindOpenMapItemSlot = 0

    ' Checar por subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenPlayerSlot() As Long
    Dim i As Long

    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

    FindOpenSpellSlot = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then

            ' Ter certeza que não queremos um nome pequeno
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Function GetNpcHPRegen(ByVal NpcNum As Long)
    Dim i As Long

    ' Prevenir subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If

    i = Int(Npc(NpcNum).DEF / 3)

    If i < 1 Then i = 1
    GetNpcHPRegen = i
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long)

    ' Prevenir subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If

    GetNpcMaxHP = Npc(NpcNum).MaxHp
End Function

Function GetNpcMaxMP(ByVal NpcNum As Long)

    ' Prevenir subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If

    GetNpcMaxMP = Npc(NpcNum).Magi * 2
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long)

    ' Prevenir subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If

    GetNpcMaxSP = Npc(NpcNum).Speed * 2
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim WeaponSlot As Long

    GetPlayerDamage = (Rnd * 5) - 2

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    GetPlayerDamage = Int(GetPlayerstr(Index) / 2)

    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2

        If GetPlayerInvItemDur(Index, WeaponSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)

            If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                Call BattleMsg(Index, "Sua " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " se quebrou.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, WeaponSlot) <= 10 Then
                    Call BattleMsg(Index, "Sua " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " está quase quebrando! Dur: " & GetPlayerInvItemDur(Index, WeaponSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1), Yellow, 0)
                End If
            End If

        Else

            If GetPlayerInvItemDur(Index, WeaponSlot) < 0 Then
                Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) + 1)

                If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                    Call BattleMsg(Index, "Sua " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " se quebrou.", Yellow, 0)
                    Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
                Else

                    If GetPlayerInvItemDur(Index, WeaponSlot) >= -10 Then
                        Call BattleMsg(Index, "Sua " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " está quase quebrando! Dur: " & GetPlayerInvItemDur(Index, WeaponSlot) * -1 & "/" & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1) * -1, Yellow, 0)
                    End If
                End If
            End If
        End If
    End If

    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If

End Function

Function GetPlayerHPRegen(ByVal Index As Long)
    Dim i As Long

    If GetVar(App.Path & "\Dados.ini", "CONFIG", "HPRegen") = 1 Then

        ' Prevenir subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerHPRegen = 0
            Exit Function
        End If

        i = Int(GetPlayerDEF(Index) / 2)

        If i < 2 Then i = 2
        GetPlayerHPRegen = i
    End If

End Function

Function GetPlayerMPRegen(ByVal Index As Long)
    Dim i As Long

    If GetVar(App.Path & "\Dados.ini", "CONFIG", "MPRegen") = 1 Then

        ' Prevenir subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerMPRegen = 0
            Exit Function
        End If

        i = Int(GetPlayerMAGI(Index) / 2)

        If i < 2 Then i = 2
        GetPlayerMPRegen = i
    End If

End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long

    GetPlayerProtection = 0

    ' Prevenir subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    ShieldSlot = GetPlayerShieldSlot(Index)
    GetPlayerProtection = Int(GetPlayerDEF(Index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2

        If GetPlayerInvItemDur(Index, ArmorSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)

            If GetPlayerInvItemDur(Index, ArmorSlot) = 0 Then
                Call BattleMsg(Index, "Sua " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " quebrou.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, ArmorSlot) <= 10 Then
                    Call BattleMsg(Index, "Sua " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " está quase quebrando! Dur: " & GetPlayerInvItemDur(Index, ArmorSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If

    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2

        If GetPlayerInvItemDur(Index, HelmSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
                Call BattleMsg(Index, "Seu " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " quebrou.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, HelmSlot) <= 10 Then
                    Call BattleMsg(Index, "Seu " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " está quase quebrando! Dur: " & GetPlayerInvItemDur(Index, HelmSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If

    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data2

        If GetPlayerInvItemDur(Index, ShieldSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) <= 0 Then
                Call BattleMsg(Index, "Seu " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " quebrou.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ShieldSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, ShieldSlot) <= 10 Then
                    Call BattleMsg(Index, "Seu " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " está quase quebrando! Dur: " & GetPlayerInvItemDur(Index, ShieldSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If

End Function

Function GetPlayerSPRegen(ByVal Index As Long)
    Dim i As Long

    If GetVar(App.Path & "\Dados.ini", "CONFIG", "SPRegen") = 1 Then

        ' Prevenir subscript...
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerSPRegen = 0
            Exit Function
        End If

        i = Int(GetPlayerSPEED(Index) / 2)

        If i < 2 Then i = 2
        GetPlayerSPRegen = i
    End If

End Function

Function GetSpellReqLevel(ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
End Function

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long, N As Long

    N = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            N = N + 1
        End If

    Next

    GetTotalMapPlayers = N
End Function

Sub GiveItem(ByVal Index As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long)
    Dim i As Long

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Checar se o inventário está cheio
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If

        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Seu inventário está cheio.", BrightRed)
    End If

End Sub

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    HasItem = 0

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Checar pra ver se o jogador possui o item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    HasSpell = False

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Public Function IsValid(ByVal x As Long, _
   ByVal y As Long) As Boolean
    IsValid = True

    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then IsValid = False
End Function

Sub JoinGame(ByVal Index As Long)
    Dim MOTD As String

    ' Setar a flag pra saber quem tá no jogo
    Player(Index).InGame = True

    ' Mandar um Ok pro cliente pra receber info
    Call SendDataTo(Index, "LOGINOK" & SEP_CHAR & Index & END_CHAR)
    Call SendDataTo(Index, "sound" & SEP_CHAR & "Entrando" & END_CHAR)

    ' Mandar mais besteiras... Não precisa explicar
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendSpeech(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendOnlineList
    Call SendFriendListTo(Index)
    Call SendFriendListToNeeded(GetPlayerName(Index))

    ' Teleportar o jogador para a localização salva
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), False)
    Call SendPlayerData(Index)

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Principal.txt", "JoinGame " & Index
    Else

        If Not ExistVar("motd.ini", "MOTD", "Msg") Then Call MsgBox("OMG OMG!")
        MOTD = GetVar("motd.ini", "MOTD", "Msg")

        ' Mandar uma mensagem global que ele entrou...
        If GetPlayerAccess(Index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(Index) & " entrou no " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " entrou no " & GAME_NAME & "!", 15)
        End If

        Call SendDataToAllBut(Index, "sound" & SEP_CHAR & "JogadorEntrou" & END_CHAR)

        ' Mandar um bem vindo
        Call PlayerMsg(Index, "Bem vindo ao " & GAME_NAME & "!", 15)

        ' Mandar motd
        If Trim$(MOTD) <> vbNullString Then
            Call PlayerMsg(Index, "MOTD: " & MOTD, 11)
        End If
    End If

    ' Mandar quem está online
    Call SendWhosOnline(Index)
    Call ShowPLR(Index)

    ' Mandar a flag, assim vão poder fazer algo
    Call SendDataTo(Index, "INGAME" & END_CHAR)
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim N As Long
    Dim i As Long

    If Player(Index).InGame = True Then
        Player(Index).InGame = False
        Call SendDataTo(Index, "sound" & SEP_CHAR & "SaindoDoServidor" & END_CHAR)
        Call SendDataToAllBut(Index, "sound" & SEP_CHAR & "JogadorSaiu" & END_CHAR)

        ' Checar se o player é o único no mapa, caso sim para npc
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If

        ' Checar se o player tá em uma party e cancelar para que o ele não continue pegando experiência a partir dessa party.
        If Player(Index).InParty = YES Then
            N = 0

            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Index).PartyID).Member(i) = Index Then N = i
            Next

            For i = N To MAX_PARTY_MEMBERS - 1
                Party(Player(Index).PartyID).Member(i) = Party(Player(Index).PartyID).Member(i + 1)
            Next

            Party(Player(Index).PartyID).Member(MAX_PARTY_MEMBERS) = 0
            N = 0

            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Index).PartyID).Member(i) <> 0 And Party(Player(Index).PartyID).Member(i) <> Index Then
                    N = N + 1
                    Call PlayerMsg(Party(Player(Index).PartyID).Member(i), GetPlayerName(Index) & " saiu do grupo.", Pink)
                End If

            Next

            If N < 2 Then
                If Party(Player(Index).PartyID).Member(1) <> 0 Then
                    Call PlayerMsg(Party(Player(Index).PartyID).Member(1), "O grupo debandou.", Pink)
                    Player(Party(Player(Index).PartyID).Member(1)).InParty = 0
                    Player(Party(Player(Index).PartyID).Member(1)).PartyID = 0
                    Party(Player(Index).PartyID).Member(1) = 0
                End If
            End If

            Player(Index).PartyID = 0
            Player(Index).InParty = 0
        End If

        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Principal.txt", "LeftGame " & Index
        Else

            ' Checar por mapa de boot
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
                Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
                Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
            End If

            ' Mandar uma mensagem dizendo que ele/ela saiu
            If GetPlayerAccess(Index) <= 1 Then
                Call GlobalMsg(GetPlayerName(Index) & " saiu do " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(Index) & " saiu do " & GAME_NAME & "!", 15)
            End If
        End If

        Call TakeFromGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Call TakeFromGrid(GetPlayerMap(Index), Player(Index).Pet.x, Player(Index).Pet.y)

        Call SavePlayer(Index)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " saiu do " & GAME_NAME & ".", True)
        Call SendLeftGame(Index)
        Call RemovePLR

        For N = 1 To MAX_PLAYERS
            Call ShowPLR(N)
        Next

    End If

    Call SendFriendListToNeeded(GetPlayerName(Index))
    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

' Sistema de loc (localização)

Public Function MakeLoc(ByVal x As Long, _
   ByVal y As Long) As Long
    MakeLoc = (y * MAX_MAPX) + x
End Function

Public Function MakeX(ByVal Loc As Long) As Long
    MakeX = Loc - (MakeY(Loc) * MAX_MAPX)
End Function

Public Function MakeY(ByVal Loc As Long) As Long
    MakeY = Int(Loc / MAX_MAPX)
End Function

Sub NpcAttackPet(ByVal MapNpcNum As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
    Dim Name As String
    Dim MapNum As Long
    Dim Packet As String

    ' Checar por subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Checar por subscript out of range
    If MapNpc(Player(Victim).Pet.Map, MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Mandar esta packet, assim vão ver o npc atacando.
    Call SendDataToMap(Player(Victim).Pet.Map, "NPCATTACKPET" & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & END_CHAR)
    MapNum = Player(Victim).Pet.Map
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= Player(Victim).Pet.HP Then
        Call BattleMsg(Victim, "Your pet died!", Red, 1)
        Player(Victim).Pet.Alive = NO
        Call TakeFromGrid(Player(Victim).Pet.Map, Player(Victim).Pet.x, Player(Victim).Pet.y)
        MapNpc(MapNum, MapNpcNum).Target = 0
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Victim & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.x & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.y & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataTo(Victim, Packet)
        Call SendDataToMapBut(Victim, Player(Victim).Pet.Map, Packet)
    Else

        ' Pet não morreu =), apenas fazer dano
        Player(Victim).Pet.HP = Player(Victim).Pet.HP - Damage
        Packet = "PETHP" & SEP_CHAR & Player(Victim).Pet.Level * 5 & SEP_CHAR & Player(Victim).Pet.HP & END_CHAR
        Call SendDataTo(Victim, Packet)
    End If

    'Call SendDataTo(Victim, "BLITNPCDMGPET" & SEP_CHAR & Damage & END_CHAR)
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long
    Dim OldMap, oldx, oldy As Long

    ' Checar por subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Checar por subscript out of range
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Mandar esta packet, assim podem ver a pessoa atacando
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & END_CHAR)
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= GetPlayerHP(Victim) Then

        ' Falar dano
        Call BattleMsg(Victim, "Você foi atacado e perdeu " & Damage & " de dano.", BrightRed, 1)

        'Call PlayerMsg(Victim, "Um(a) " & Name & " ataca você tirando " & Damage & " de dano.", BrightRed)
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " foi morto por um(a) " & Name, BrightRed)
        Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Morte" & END_CHAR)

        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Principal.txt", "DropItems " & Victim
            Else

                If GetPlayerWeaponSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                End If

                If GetPlayerArmorSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                End If

                If GetPlayerHelmetSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                End If

                If GetPlayerShieldSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                End If
            End If

            ' Calcular EXP retirada do jogador
            Exp = Int(GetPlayerExp(Victim) / 3)

            ' Ter certeza de não retirar menos que 0
            If Exp < 0 Then
                Exp = 0
            End If

            If Exp = 0 Then
                Call BattleMsg(Victim, "Você não perdeu experiência.", BrightRed, 0)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, "Você perdeu " & Exp & " de experiência.", BrightRed, 0)
            End If
        End If

        OldMap = GetPlayerMap(Victim)
        oldx = GetPlayerX(Victim)
        oldy = GetPlayerY(Victim)

        ' Warp player away
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Principal.txt", "OnDeath " & Victim
        Else
            Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        End If

        Call UpdateGrid(OldMap, oldx, oldy, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim))

        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)

        ' Set NPC target to 0
        MapNpc(MapNum, MapNpcNum).Target = 0

        ' A vitima é um PK?
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If

    Else

        ' Jogador não está morto, apenas fazer dano
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)

        ' Falar dano
        Call BattleMsg(Victim, "Você foi atacado e perdeu " & Damage & " de dano.", BrightRed, 1)

        'Call PlayerMsg(Victim, "Um(a) " & Name & " ataca você tirando " & Damage & " de dano.", BrightRed)
    End If

    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Dor" & END_CHAR)
End Sub

Sub NpcDir(ByVal MapNum As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Dir As Long)
    Dim Packet As String

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub NpcMove(ByVal MapNum As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
    Dim Packet As String
    Dim x As Long
    Dim y As Long

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    x = DirToX(MapNpc(MapNum, MapNpcNum).x, Dir)
    y = DirToY(MapNpc(MapNum, MapNpcNum).y, Dir)
    Call UpdateGrid(MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y, MapNum, x, y)
    MapNpc(MapNum, MapNpcNum).y = y
    MapNpc(MapNum, MapNpcNum).x = x
    Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PetAttackNpc(ByVal Attacker As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Damage As Long)
    Dim Name As String
    Dim N As Long, i As Long
    Dim MapNum As Long, NpcNum As Long
    Dim Dir As Long, x As Long, y As Long
    Dim Packet As String

    ' Checar por subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Mandar packet pra ver pet atacando
    Call SendDataToMap(Player(Attacker).Pet.Map, "PETATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & END_CHAR)
    MapNum = Player(Attacker).Pet.Map
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(NpcNum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount

    For Dir = 0 To 3

        If MapNpc(MapNum, NpcNum).x = DirToX(Player(Attacker).Pet.x, Dir) And MapNpc(MapNum, NpcNum).y = DirToY(Player(Attacker).Pet.y, Dir) Then
            Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & Attacker & END_CHAR
            Call SendDataToMap(Player(Attacker).Pet.Map, Packet)
        End If

    Next

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then

        For i = 1 To MAX_NPC_DROPS

            ' Drops de coisas que o pet matou
            N = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
            End If

        Next

        Call BattleMsg(Attacker, "Seu pet matou um(a) " & Name & ".", Red, 1)

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
        Call TakeFromGrid(MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)

        ' Setar alvo para 0
        If Player(Attacker).Pet.TargetType = TARGET_TYPE_NPC And Player(Attacker).Pet.Target = MapNpcNum Then
            Player(Attacker).Pet.Target = 0
            Player(Attacker).Pet.TargetType = 0
            Player(Attacker).Pet.MapToGo = 0
        End If

    Else

        ' NPC não está morto, apenas fazer dano
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage

        ' Setar o alvo do npc para o pet
        MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_PET
        MapNpc(MapNum, MapNpcNum).Target = Attacker

        ' Agora, checar pela IA dos guardas.
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then

            For i = 1 To MAX_MAP_NPCS

                If MapNpc(MapNum, i).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, i).TargetType = TARGET_TYPE_PET
                    MapNpc(MapNum, i).Target = Attacker
                End If

            Next

        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
    ' Reset attack timer
    Player(Attacker).Pet.AttackTimer = GetTickCount
End Sub

Sub PetMove(ByVal PetNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
    Dim Packet As String
    Dim x As Long
    Dim y As Long
    Dim i As Long

    If GetPlayerMap(PetNum) <= 0 Or GetPlayerMap(PetNum) > MAX_MAPS Or PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    Player(PetNum).Pet.Dir = Dir
    x = DirToX(Player(PetNum).Pet.x, Dir)
    y = DirToY(Player(PetNum).Pet.y, Dir)

    If IsValid(x, y) Then
        If Grid(Player(PetNum).Pet.Map).Loc(x, y).Blocked = True Then
            Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & PetNum & END_CHAR
            Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
            Exit Sub
        End If

        Call UpdateGrid(Player(PetNum).Pet.Map, Player(PetNum).Pet.x, Player(PetNum).Pet.y, Player(PetNum).Pet.Map, x, y)
        Player(PetNum).Pet.y = y
        Player(PetNum).Pet.x = x
        Packet = "PETMOVE" & SEP_CHAR & PetNum & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
    Else
        i = Player(PetNum).Pet.Map

        If Dir = DIR_UP Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Up
            Player(PetNum).Pet.y = MAX_MAPY
        End If

        If Dir = DIR_DOWN Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Down
            Player(PetNum).Pet.y = 0
        End If

        If Dir = DIR_LEFT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Left
            Player(PetNum).Pet.x = MAX_MAPX
        End If

        If Dir = DIR_RIGHT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
            Player(PetNum).Pet.x = 0
        End If

        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & PetNum & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Map & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.x & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.y & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.HP & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
        Call SendDataToMap(i, Packet)
    End If

End Sub

Sub PlayerMapDropItem(ByVal Index As Long, _
   ByVal InvNum As Long, _
   ByVal Amount As Long)
    Dim i As Long

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(Index))

        If i <> 0 Then
            MapItem(GetPlayerMap(Index), i).Dur = 0

            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type

                Case ITEM_TYPE_ARMOR

                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_WEAPON

                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_HELMET

                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_SHIELD

                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select

            MapItem(GetPlayerMap(Index), i).num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)

            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then

                ' Checar se há mais e então dropar
                If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " deixou " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(Index), i).Value = Amount
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " deixou " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                End If

            Else

                ' Não é um objeto, então é fácil
                MapItem(GetPlayerMap(Index), i).Value = 0

                If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " deixou um(a) " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - Ind.", Yellow)
                    Else

                        If Item(GetPlayerInvItemNum(Index, InvNum)).Data1 > 0 Then
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " deixou um(a) " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".", Yellow)
                        Else
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " deixou um(a) " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 * -1 & ".", Yellow)
                        End If
                    End If

                Else
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " deixou um(a) " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                End If

                Call SetPlayerInvItemNum(Index, InvNum, 0)
                Call SetPlayerInvItemValue(Index, InvNum, 0)
                Call SetPlayerInvItemDur(Index, InvNum, 0)
            End If

            ' Send inventory update
            Call SendInventoryUpdate(Index, InvNum)

            ' Spawnar o item antes de setar o número
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).num, Amount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsg(Index, "Existem muitos itens no chão.", BrightRed)
        End If
    End If

End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim N As Long
    Dim MapNum As Long
    Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS

        ' Ver se tem um item por aqui...
        If (MapItem(MapNum, i).num > 0) And (MapItem(MapNum, i).num <= MAX_ITEMS) Then

            ' Checar se o item está no mesmo lugar que o jogador
            If (MapItem(MapNum, i).x = GetPlayerX(Index)) And (MapItem(MapNum, i).y = GetPlayerY(Index)) Then

                ' Achar um slot aberto
                N = FindOpenInvSlot(Index, MapItem(MapNum, i).num)

                ' Slot livre?
                If N <> 0 Then

                    ' Setar item no inventário do jogador
                    Call SetPlayerInvItemNum(Index, N, MapItem(MapNum, i).num)

                    If Item(GetPlayerInvItemNum(Index, N)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(Index, N, GetPlayerInvItemValue(Index, N) + MapItem(MapNum, i).Value)
                        Msg = "Você pegou um(a) " & MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, N)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(Index, N, 0)
                        Msg = "Você pegou um(a) " & Trim$(Item(GetPlayerInvItemNum(Index, N)).Name) & "."
                    End If

                    Call SetPlayerInvItemDur(Index, N, MapItem(MapNum, i).Dur)

                    ' Erase item from the map
                    MapItem(MapNum, i).num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).x = 0
                    MapItem(MapNum, i).y = 0
                    Call SendInventoryUpdate(Index, N)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call PlayerMsg(Index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(Index, "Seu inventário está cheio.", BrightRed)
                    Exit Sub
                End If
            End If
        End If

    Next

End Sub

Sub PlayerMove(ByVal Index As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
    Dim Packet As String
    Dim MapNum As Long
    Dim x As Long
    Dim y As Long
    Dim oldx As Long
    Dim oldy As Long
    Dim OldMap As Long
    Dim Moved As Byte

    ' Tentaram nos hackear!!!! =/
    'If Moved = NO Then
    'Call HackingAttempt(index, "Modificação de Posição")
    'Exit Sub
    'End If
    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Moved = NO
    x = DirToX(GetPlayerX(Index), Dir)
    y = DirToY(GetPlayerY(Index), Dir)
    Call TakeFromGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Mover o pet do jogador se precisar
    If Player(Index).Pet.Alive = YES Then
        If Player(Index).Pet.Map = GetPlayerMap(Index) And Player(Index).Pet.x = x And Player(Index).Pet.y = y Then
            If Grid(GetPlayerMap(Index)).Loc(DirToX(x, Dir), DirToY(y, Dir)).Blocked = False Then
                Call UpdateGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y, Player(Index).Pet.Map, DirToX(x, Dir), DirToY(y, Dir))
                Player(Index).Pet.y = DirToY(y, Dir)
                Player(Index).Pet.x = DirToX(x, Dir)
                Packet = "PETMOVE" & SEP_CHAR & Index & SEP_CHAR & DirToX(x, Dir) & SEP_CHAR & DirToY(y, Dir) & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
                Call SendDataToMap(Player(Index).Pet.Map, Packet)
            End If
        End If
    End If

    ' Checar por boundries (WTF?)
    If IsValid(x, y) Then

        ' Ter certeza se a tile pode ser andada
        If Grid(GetPlayerMap(Index)).Loc(x, y).Blocked = False Then

            ' Ter certeza se a tile requer uma chave e se está aberta
            If (Map(GetPlayerMap(Index)).Tile(x, y).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(x, y).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES) Then
                Call SetPlayerX(Index, x)
                Call SetPlayerY(Index, y)
                Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
                Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                Moved = YES
            End If
        End If

    Else

        ' Checar para ver se podemos move-la para outro mapa
        If Map(GetPlayerMap(Index)).Up > 0 And Dir = DIR_UP Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
            Moved = YES
        End If

        If Map(GetPlayerMap(Index)).Down > 0 And Dir = DIR_DOWN Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
            Moved = YES
        End If

        If Map(GetPlayerMap(Index)).Left > 0 And Dir = DIR_LEFT Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
            Moved = YES
        End If

        If Map(GetPlayerMap(Index)).Right > 0 And Dir = DIR_RIGHT Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
            Moved = YES
        End If
    End If

    If Moved = NO Then Call SendPlayerXY(Index)
    If GetPlayerX(Index) < 0 Or GetPlayerY(Index) < 0 Or GetPlayerX(Index) > MAX_MAPX Or GetPlayerY(Index) > MAX_MAPY Or GetPlayerMap(Index) <= 0 Then
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If

    ' Código das tiles que recuperam
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
        If GetPlayerHP(Index) < GetPlayerMaxHP(Index) Then
            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            Call SendHP(Index)
            Call PlayerMsg(Index, "Você sente uma rejuvenação no seu corpo!", BrightGreen)
        End If
    End If

    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(Index, 0)
        Call PlayerMsg(Index, "Você sente calafrios, pois a morte se aproxima. Nada pôde ser feito, agora você está morto.", BrightRed)

        ' Teleportar jogador
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Principal.txt", "OnDeath " & Index
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If

        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Moved = YES
    End If

    If IsValid(x, y) Then
        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
            If TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Chave" & END_CHAR)
            End If
        End If
    End If

    ' Checar quanto às warp tiles
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
        Call PlayerWarp(Index, MapNum, x, y)
        Moved = YES
    End If

    Call AddToGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Checar pela Chave
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2

        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)

            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(Index), "Uma porta foi destrancada!", White)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
            End If

            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Chave" & END_CHAR)
        End If
    End If

    ' Check for shop
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SHOP Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 > 0 Then
            Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
        Else
            Call PlayerMsg(Index, "Não há lojas aqui.", BrightRed)
        End If
    End If

    ' Checar se o jogador pisou nas tiles de mudança de sprite
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "Você já usa essa sprite!", BrightRed)
            Exit Sub
        Else

            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 0 & END_CHAR)
            Else

                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "Essa sprite irá custar " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(Index, "Essa sprite irá custar um(a) " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                End If

                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 1 & END_CHAR)
            End If
        End If
    End If

    ' Checar se o jogador pisou nas tiles de mudança de sprite
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 > 0 Then
            If GetPlayerClass(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                Call PlayerMsg(Index, "Você não está na classe requerida!", BrightRed)
                Exit Sub
            End If
        End If

        If GetPlayerClass(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "Você já é dessa classe!", BrightRed)
        Else

            If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
                If GetPlayerSprite(Index) = Class(GetPlayerClass(Index)).MaleSprite Then
                    Call SetPlayerSprite(Index, Class(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).MaleSprite)
                End If

            Else

                If GetPlayerSprite(Index) = Class(GetPlayerClass(Index)).FemaleSprite Then
                    Call SetPlayerSprite(Index, Class(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).FemaleSprite)
                End If
            End If

            Call SetPlayerstr(Index, (Player(Index).Char(Player(Index).CharNum).STR - Class(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF - Class(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi - Class(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - Class(GetPlayerClass(Index)).Speed))
            Call SetPlayerClass(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
            Call SetPlayerstr(Index, (Player(Index).Char(Player(Index).CharNum).STR + Class(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF + Class(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi + Class(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + Class(GetPlayerClass(Index)).Speed))
            Call PlayerMsg(Index, "Sua nova classe é " & Trim$(Class(GetPlayerClass(Index)).Name) & "!", BrightGreen)
            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
        End If
    End If

    ' Checar se o jogador pisou em uma tile de notice x_X
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), Black)
        End If

        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), Grey)
        End If

        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & END_CHAR)
    End If

    ' Mesma coisa do de cima, sendo que de som
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & END_CHAR)
    End If

    If SCRIPTING = 1 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Principal.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        End If
    End If

End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional sound As Boolean = True)
    Dim OldMap As Long

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Checar se há um npc no mapa que o jogador está saindo e entãoa falar a mensagem de despedida do servidor.
    'If Trim$(Shop(ShopNum).LeaveSay) <> "" Then
    'Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & " : " & Trim$(Shop(ShopNum).LeaveSay) & "", SayColor)
    'End If
    ' Salvar o mapa
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    Call UpdateGrid(OldMap, GetPlayerX(Index), GetPlayerY(Index), MapNum, x, y)
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)

    If Player(Index).Pet.Alive = YES Then
        Player(Index).Pet.MapToGo = -1
        Player(Index).Pet.XToGo = -1
        Player(Index).Pet.YToGo = -1
        Player(Index).Pet.Map = MapNum
        Player(Index).Pet.x = x
        Player(Index).Pet.y = y
    End If

    ' Não precisa explicar
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Setar para saber os npcs nos mapas
    PlayersOnMap(MapNum) = YES
    Player(Index).GettingMap = YES

    If sound Then Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Portal" & END_CHAR)
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
End Sub

Public Sub RemovePLR()
    frmServer.lvUsers.ListItems.Clear
End Sub

Sub SetUpGrid()
    Dim i As Long
    Dim x As Long
    Dim y As Long

    Call ClearGrid

    For i = 1 To MAX_MAPS
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY

                If Map(i).Tile(x, y).Type = TILE_TYPE_BLOCKED Then Grid(i).Loc(x, y).Blocked = True
            Next
        Next

        For x = 1 To MAX_MAP_NPCS

            If MapNpc(i, x).num > 0 Then
                Grid(i).Loc(MapNpc(i, x).x, MapNpc(i, x).y).Blocked = True
            End If

        Next
    Next

End Sub

Public Sub ShowPLR(ByVal Index As Long)
    Dim ls As ListItem

    On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) = True Then
        frmServer.lvUsers.ListItems.Remove Index
    End If

    Set ls = frmServer.lvUsers.ListItems.add(Index, , Index)

    If IsPlaying(Index) = False Then
        ls.SubItems(1) = vbNullString
        ls.SubItems(2) = vbNullString
        ls.SubItems(3) = vbNullString
        ls.SubItems(4) = vbNullString
        ls.SubItems(5) = vbNullString
    Else
        ls.SubItems(1) = GetPlayerLogin(Index)
        ls.SubItems(2) = GetPlayerName(Index)
        ls.SubItems(3) = GetPlayerLevel(Index)
        ls.SubItems(4) = GetPlayerSprite(Index)
        ls.SubItems(5) = GetPlayerAccess(Index)
    End If

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnItem(ByVal ItemNum As Long, _
   ByVal ItemVal As Long, _
   ByVal MapNum As Long, _
   ByVal x As Long, _
   ByVal y As Long)
    Dim i As Long

    ' Checar por subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Procurar um slot de mapa livre
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long, _
   ByVal ItemDur As Long, _
   ByVal MapNum As Long, _
   ByVal x As Long, _
   ByVal y As Long)
    Dim Packet As String
    Dim i As Long

    ' Checar por subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).num = ItemNum
        MapItem(MapNum, i).Value = ItemVal

        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If

        Else
            MapItem(MapNum, i).Dur = 0
        End If

        MapItem(MapNum, i).x = x
        MapItem(MapNum, i).y = y
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim x As Long
    Dim y As Long

    ' Checar por subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawnar o que nós temos
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX

            ' Checar se a tile é um item caso alguem tenha dropado um item nela.
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Mesmo de cima, mas setar valor pra 0
                If Item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If

        Next
    Next

End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim NpcNum As Long
    Dim i As Long, x As Long, y As Long
    Dim Spawned As Boolean

    ' Checar por subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    Spawned = False
    NpcNum = Map(MapNum).Npc(MapNpcNum)

    If NpcNum > 0 Then
        If GameTime = TIME_NIGHT Then
            If Npc(NpcNum).SpawnTime = 1 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If

        Else

            If Npc(NpcNum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        End If

        MapNpc(MapNum, MapNpcNum).num = NpcNum
        MapNpc(MapNum, MapNpcNum).Target = 0
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

        If Map(MapNum).NpcSpawn(MapNpcNum).Used <> 1 Then

            ' Tentar randomicamente recolocar as sprites.
            For i = 1 To 100
                x = Int(Rnd * MAX_MAPX)
                y = Int(Rnd * MAX_MAPY)

                ' Olhar se a tile pode ser andada
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum, MapNpcNum).x = x
                    MapNpc(MapNum, MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If

            Next

            ' Não spawnar, apenas olhar as tiles
            If Not Spawned Then

                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX

                        If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                            MapNpc(MapNum, MapNpcNum).x = x
                            MapNpc(MapNum, MapNpcNum).y = y
                            Spawned = True
                            Exit For
                        End If

                    Next
                Next

            End If

        Else
            MapNpc(MapNum, MapNpcNum).x = Map(MapNum).NpcSpawn(MapNpcNum).x
            MapNpc(MapNum, MapNpcNum).y = Map(MapNum).NpcSpawn(MapNpcNum).y
            Spawned = True
        End If

        ' Se nós sucedemos, mandar mensagem para todos.
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Call AddToGrid(MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
End Sub

Sub TakeFromGrid(ByVal OldMap, _
   ByVal oldx, _
   ByVal oldy)
    Grid(OldMap).Loc(oldx, oldy).Blocked = False

    If Map(OldMap).Tile(oldx, oldy).Type = TILE_TYPE_BLOCKED Then Grid(OldMap).Loc(oldx, oldy).Blocked = True
End Sub

Sub TakeItem(ByVal Index As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long)
    Dim i As Long, N As Long
    Dim TakeItem As Boolean

    TakeItem = False

    ' Checar por subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For i = 1 To MAX_INV

        ' Checar para ver se o jogador possui o item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Tamos tentando pegar o que eles tem? Setar para zero!
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If

            Else

                ' Checar para ver se há algum tipo de Armadura/Arma.
                Select Case Item(GetPlayerInvItemNum(Index, i)).Type

                    Case ITEM_TYPE_WEAPON

                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If i = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Ver se o item que estamos pegando já está equipado
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If

                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_ARMOR

                        If GetPlayerArmorSlot(Index) > 0 Then
                            If i = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Ver se o item que estamos pegando já está equipado
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If

                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_HELMET

                        If GetPlayerHelmetSlot(Index) > 0 Then
                            If i = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Ver se o item que estamos pegando já está equipado
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If

                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_SHIELD

                        If GetPlayerShieldSlot(Index) > 0 Then
                            If i = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Ver se o item que estamos pegando já está equipado
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If

                        Else
                            TakeItem = True
                        End If

                End Select

                N = Item(GetPlayerInvItemNum(Index, i)).Type

                ' Checar se já não é uma arma equipavel, caso sim, não pegar ela.
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                End If
            End If

            If TakeItem = True Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemDur(Index, i, 0)

                ' Mandar o update de inventário
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        End If

    Next

End Sub

Function TotalOnlinePlayers() As Long
    Dim i As Long

    TotalOnlinePlayers = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Sub UpdateGrid(ByVal OldMap, _
   ByVal oldx, _
   ByVal oldy, _
   ByVal NewMap, _
   ByVal NewX, _
   ByVal NewY)
    Grid(OldMap).Loc(oldx, oldy).Blocked = False
    Grid(NewMap).Loc(NewX, NewY).Blocked = True

    If Map(OldMap).Tile(oldx, oldy).Type = TILE_TYPE_BLOCKED Then Grid(OldMap).Loc(oldx, oldy).Blocked = True
End Sub

Sub ResetMapGrid(ByVal MapNum As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long

    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            Grid(MapNum).Loc(x, y).Blocked = False

            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_BLOCKED Then Grid(MapNum).Loc(x, y).Blocked = True
        Next
    Next

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum, i).num > 0 Then
            Grid(MapNum).Loc(MapNpc(MapNum, i).x, MapNpc(MapNum, i).y).Blocked = True
        End If

    Next

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Grid(MapNum).Loc(GetPlayerX(i), GetPlayerY(i)).Blocked = True
            End If
        End If

    Next

End Sub
