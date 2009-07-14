Attribute VB_Name = "modServerTCP"
' Copyright (c) 2009 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

Option Explicit

'Dim ZePacket() As String ' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO
'Dim NumParse As Long ' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO
'Dim ParseIndex As Long ' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO
Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then

            ' Woohoo, nós podemos conectá-los
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String
    Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If

    Next

End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & END_CHAR
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
End Sub

Sub CloseSocket(ByVal Index As Long)

    ' Garante que o jogador está/estava jogando, se sim, salva ele.
    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd(frmServer.txtText(0), "Conexão com " & GetPlayerIP(Index) & " foi terminada.", True)
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(Index, "Você perdeu a conexão com o " & GAME_NAME & " por (" & Reason & ").")
    End If

End Sub

'                --- INSTRUCTIONS ON HOW TO TURN SAFE MODE OFF ---

' FALTOU PACIÊNCIA, QUEM SABE NA PRÓXIMA

'  INTRO:
'  Safe Mode is meant to prevent your server from getting knocked down.
'  It fixes all parse subscript out of range errors.
'  It is recommended to be kept OFF, but you should turn it on when testing.
'  If you believe a person may be trying to hack, turn it on!
'  A person may knock down your server by sending invalid packet data.
'  This prevents that and the server going down because of stupid coding errors.
'  A person has other methods to knock a server down, but this is the easiest way.
'  INSTRUCTIONS:
'  Search this module for all occurences of "SAFE MODE"
'  Follow the instructions!

Sub HandleData(ByVal Index As Long, ByVal Data As String)
    Dim Parse() As String ' MODO DE SEGURANÇA -- "Descomente" para DESLIGÁ-LO, comente para LIGÁ-LO
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim CharNum As Long
    Dim Msg As String
    Dim MsgTo As Long
    Dim Dir As Long
    Dim InvNum As Long
    Dim Amount As Long
    Dim Damage As Long
    Dim PointType As Byte
    Dim PointQuant As Integer
    Dim Movement As Long
    Dim i As Long, N As Long, x As Long, y As Long, f As Long
    Dim MapNum As Long
    Dim s As String
    Dim ShopNum As Long, ItemNum As Long
    Dim DurNeeded As Long, GoldNeeded As Long
    Dim z As Long
    Dim Packet As String
    Dim o As Long

    'ParseIndex = index ' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO
    ' Handle Data
    Parse = Split(Data, SEP_CHAR) ' MODO DE SEGURANÇA -- "Descomente" para DESLIGÁ-LO, comente para LIGÁ-LO

    'ZePacket = Split(Data, SEP_CHAR) ' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO
    'NumParse = UBound(ZePacket) ' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO
    ' Parse's Without Being Online
    If Not IsPlaying(Index) Then

        Select Case LCase$(Parse(0))

            Case "getinfo"
                Call SendInfo(Index)
                Exit Sub

            Case "gatglasses"
                Call SendNewCharClasses(Index)
                Exit Sub

            Case "newfaccountied"

                If Not IsLoggedIn(Index) Then
                    Name = Parse(1)
                    Password = Parse(2)

                    For i = 1 To Len(Name)
                        N = Asc(Mid$(Name, i, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(Index, "Nome Inválido! Apenas use letras, números e espaços.", 1)
                            Exit Sub
                        End If

                    Next

                    If Not AccountExist(Name) Then
                        Call PlainMsg(Index, "Sua conta foi criada com sucesso.", 1)
                        Call AddAccount(Index, Name, Password)
                        Call TextAdd(frmServer.txtText(0), "Conta " & Name & " foi criada.", True)
                        Call AddLog("Conta " & Name & " foi criada.", PLAYER_LOG)
                        Call CloseSocket(Index)
                    Else
                        Call PlainMsg(Index, "Desculpe, essa conta já foi pega!", 1)
                    End If
                End If

                Exit Sub

            Case "delimaccounted"

                If Not IsLoggedIn(Index) Then
                    Name = Parse(1)
                    Password = Parse(2)

                    If Not AccountExist(Name) Then
                        Call PlainMsg(Index, "Essa conta não existe.", 2)
                        Exit Sub
                    End If

                    If Not PasswordOK(Name, Password) Then
                        Call PlainMsg(Index, "Senha incorreta.", 2)
                        Exit Sub
                    End If

                    Call LoadPlayer(Index, Name)

                    For i = 1 To MAX_CHARS

                        If Trim$(Player(Index).Char(i).Name) <> vbNullString Then
                            Call DeleteName(Player(Index).Char(i).Name)
                        End If

                    Next

                    Call ClearPlayer(Index)
                    Call Kill(App.Path & "\Contas\" & Trim$(Name) & ".ini")
                    Call AddLog("Conta " & Trim$(Name) & " foi deletada.", PLAYER_LOG)
                    Call PlainMsg(Index, "Sua conta foi deletada.", 2)
                End If

                Exit Sub

            Case "logination"

                If Not IsLoggedIn(Index) Then
                    Name = Parse(1)
                    Password = Parse(2)

                    For i = 1 To Len(Name)
                        N = Asc(Mid$(Name, i, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(Index, "Contas duplas não são aceitadas!", 3)
                            Exit Sub
                        End If

                    Next

                    If Parse$(3) & "." & Parse$(4) & "." & Parse$(5) & "." & Parse$(6) <> CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION & "." & SEC_CODE Then
                        Call SendDataTo(Index, "sound" & SEP_CHAR & "NovaVersão" & END_CHAR)
                        Call PlainMsg(Index, "Versão desatualizada! Visite " & Trim$(GetVar(App.Path & "\Dados.ini", "CONFIG", "WebSite")), 3)
                        Exit Sub
                    End If

                    If Not AccountExist(Name) Then
                        Call PlainMsg(Index, "Essa conta não existe.", 3)
                        Exit Sub
                    End If

                    If Not PasswordOK(Name, Password) Then
                        Call PlainMsg(Index, "Senha incorreta.", 3)
                        Exit Sub
                    End If

                    If IsMultiAccounts(Name) Then
                        Call PlainMsg(Index, "Logins múltiplos não são aceitos.", 3)
                        Exit Sub
                    End If

                    If frmServer.Closed.Value = Checked Then
                        Call PlainMsg(Index, "O servidor está em manutenção!", 3)
                        Exit Sub
                    End If

                    Dim Packs As String

                    Packs = "MAXINFO" & SEP_CHAR
                    Packs = Packs & GAME_NAME & SEP_CHAR
                    Packs = Packs & MAX_PLAYERS & SEP_CHAR
                    Packs = Packs & MAX_ITEMS & SEP_CHAR
                    Packs = Packs & MAX_NPCS & SEP_CHAR
                    Packs = Packs & MAX_SHOPS & SEP_CHAR
                    Packs = Packs & MAX_SPELLS & SEP_CHAR
                    Packs = Packs & MAX_MAPS & SEP_CHAR
                    Packs = Packs & MAX_MAP_ITEMS & SEP_CHAR
                    Packs = Packs & MAX_MAPX & SEP_CHAR
                    Packs = Packs & MAX_MAPY & SEP_CHAR
                    Packs = Packs & MAX_EMOTICONS & SEP_CHAR
                    Packs = Packs & MAX_SPEECH & SEP_CHAR
                    Packs = Packs & END_CHAR
                    Call SendDataTo(Index, Packs)
                    Call LoadPlayer(Index, Name)
                    Call SendChars(Index)
                    Call AddLog(GetPlayerLogin(Index) & " se logou com o IP " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " se logou com o IP " & GetPlayerIP(Index) & ".", True)
                End If

                Exit Sub

            Case "addachara"
                Name = Parse(1)
                Sex = Val(Parse(2))
                Class = Val(Parse(3))
                CharNum = Val(Parse(4))

                For i = 1 To Len(Name)
                    N = Asc(Mid$(Name, i, 1))

                    If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                    Else
                        Call PlainMsg(Index, "Nome Inválido! Use apenas letras, números e espaços.", 4)
                        Exit Sub
                    End If

                Next

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "CharNum Inválido")
                    Exit Sub
                End If

                If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                    Call HackingAttempt(Index, "Sexo Inválido")
                    Exit Sub
                End If

                If Class < 1 Or Class > Max_Classes Then
                    Call HackingAttempt(Index, "Classe Inválida")
                    Exit Sub
                End If

                If CharExist(Index, CharNum) Then
                    Call PlainMsg(Index, "O personagem já existe!", 4)
                    Exit Sub
                End If

                If FindChar(Name) Then
                    Call PlainMsg(Index, "Desculpe, mas este nome já está em uso!", 4)
                    Exit Sub
                End If

                Call AddChar(Index, Name, Sex, Class, CharNum)
                Call SavePlayer(Index)
                Call AddLog("O personagem " & Name & " foi adicionado na conta de " & GetPlayerLogin(Index) & ".", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "O personagem foi criado!", 5)
                Exit Sub

            Case "delimbocharu"
                CharNum = Val(Parse(1))

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "CharNum Inválido")
                    Exit Sub
                End If

                Call DelChar(Index, CharNum)
                Call AddLog("Personagem deletado na conta de " & GetPlayerLogin(Index) & ".", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Personagem foi deletado!", 5)
                Exit Sub

            Case "usagakarim"
                CharNum = Val(Parse(1))

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "CharNum Inválido")
                    Exit Sub
                End If

                If CharExist(Index, CharNum) Then
                    Player(Index).CharNum = CharNum

                    If frmServer.GMOnly.Value = Checked Then
                        If GetPlayerAccess(Index) <= 0 Then
                            Call PlainMsg(Index, "No momento, o servidor está online apenas para GMs!", 5)

                            'Call HackingAttempt(Index, "No momento, o servidor está online apenas para GMs!")
                            Exit Sub
                        End If
                    End If

                    Call JoinGame(Index)
                    CharNum = Player(Index).CharNum
                    Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " começou a jogar " & GAME_NAME & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " começou a jogar " & GAME_NAME & ".", True)
                    Call UpdateCaption

                    If Not FindChar(GetPlayerName(Index)) Then
                        f = FreeFile
                        Open App.Path & "\Contas\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(Index)
                        Close #f
                    End If

                Else
                    Call PlainMsg(Index, "O personagem não existe!", 5)
                End If

                Exit Sub
        End Select

    End If

    ' Online e Jogando
    If IsPlaying(Index) = False Then Exit Sub
    If IsConnected(Index) = False Then Exit Sub

    Select Case LCase$(Parse(0))

            ' :::::::::::::::::::
            ' :: Guilds Packet ::
            ' :::::::::::::::::::
            ' Acesso
        Case "guildchangeaccess"

            ' Requerimentos
            If FindPlayer(Parse(1)) = 0 Then
                Call PlayerMsg(Index, "O jogador está offline.", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
                Call PlayerMsg(Index, "O jogador não está em sua guild.", Red)
                Exit Sub
            End If

            'Setar o novo acesso do jador
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Parse(2))
            Call SendPlayerData(FindPlayer(Parse(1)))
            Exit Sub

            ' Disown
        Case "guilddisown"

            ' Todos os requerimentos
            If FindPlayer(Parse(1)) = 0 Then
                Call PlayerMsg(Index, "O jogador está offline.", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
                Call PlayerMsg(Index, "O jogador não está em sua guild.", Red)
                Exit Sub
            End If

            If GetPlayerGuildAccess(FindPlayer(Parse(1))) > GetPlayerGuildAccess(Index) Then
                Call PlayerMsg(Index, "O jogador tem um guild level maior que o seu.", Red)
                Exit Sub
            End If

            'Tire o jogador da Guild
            Call SetPlayerGuild(FindPlayer(Parse(1)), vbNullString)
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
            Call SendPlayerData(FindPlayer(Parse(1)))
            Exit Sub

            ' Sair da Guild
        Case "guildleave"

            ' Veja se ele pode sair!
            If GetPlayerGuild(Index) = vbNullString Then
                Call PlayerMsg(Index, "Você não está em uma guild.", Red)
                Exit Sub
            End If

            Call SetPlayerGuild(Index, vbNullString)
            Call SetPlayerGuildAccess(Index, 0)
            Call SendPlayerData(Index)
            Exit Sub

            ' Faça uma nova guild
        Case "makeguild"

            ' O dono está online?
            If FindPlayer(Parse(1)) = 0 Then
                Call PlayerMsg(Index, "O jogador está offline.", White)
                Exit Sub
            End If

            ' Ele já está em uma guild?
            If GetPlayerGuild(FindPlayer(Parse(1))) <> vbNullString Then
                Call PlayerMsg(Index, "O jogador já está em uma guild.", Red)
                Exit Sub
            End If

            ' Se estiver tudo bem, faça a Guild
            Call SetPlayerGuild(FindPlayer(Parse(1)), (Parse(2)))
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 3)
            Call SendPlayerData(FindPlayer(Parse(1)))
            Exit Sub

            ' Colocando membro
        Case "guildmember"

            ' Requerimentos
            If FindPlayer(Parse(1)) = 0 Then
                Call PlayerMsg(Index, "O jogador está offline.", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
                Call PlayerMsg(Index, "Aquele jogador não está na sua guild.", Red)
                Exit Sub
            End If

            If GetPlayerGuildAccess(FindPlayer(Parse(1))) > 1 Then
                Call PlayerMsg(Index, "Aquele jogador já foi admitido em sua guild.", Red)
                Exit Sub
            End If

            'Tudo foi bem, acesso setado para 1
            Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(Index))
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 1)
            Call SendPlayerData(FindPlayer(Parse(1)))
            Exit Sub

            ' Fazer um Trainie
        Case "guildtrainee"

            ' Requerimentos
            If FindPlayer(Parse(1)) = 0 Then
                Call PlayerMsg(Index, "O jogador está offline.", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse(1))) <> vbNullString Then
                Call PlayerMsg(Index, "O jogador já está em sua guild.", Red)
                Exit Sub
            End If

            ' Se tudo okay, setar acesso dele da Guild para 0
            Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(Index))
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
            Call SendPlayerData(FindPlayer(Parse(1)))
            Exit Sub

            ' ::::::::::::::::::::
            ' :: Social packets ::
            ' ::::::::::::::::::::
        Case "saymsg"
            Msg = Parse(1)

            ' Proteção
            For i = 1 To Len(Msg)

                If Not ((Asc(Mid$(Msg, i, 1)) > 31 And Asc(Mid$(Msg, i, 1)) < 127) Or (Asc(Mid$(Msg, i, 1)) = 128) Or (Asc(Mid$(Msg, i, 1)) > 159 And Asc(Mid$(Msg, i, 1)) < 256)) And (Mid$(Msg, i, 1) <> SEP_CHAR And Mid$(Msg, i, 1) <> END_CHAR) Then
                    Call PlayerMsg(Index, "Uma ou mais caracteres digitadas são inválidas.", Red)
                    Exit Sub
                End If

            Next

            If frmServer.chkM.Value = Unchecked Then
                If GetPlayerAccess(Index) <= 0 Then
                    Call PlayerMsg(Index, "Mensagens do mapa foram desativadas pelo servidor!", BrightRed)
                    Exit Sub
                End If
            End If

            Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & ": " & Msg, PLAYER_LOG)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Msg, SayColor)
            Call MapMsg2(GetPlayerMap(Index), Msg, Index)
            TextAdd frmServer.txtText(3), GetPlayerName(Index) & " no mapa " & GetPlayerMap(Index) & ": " & Msg, True
            Exit Sub

        Case "emotemsg"
            Msg = Parse(1)

            ' Proteção
            For i = 1 To Len(Msg)

                If Not ((Asc(Mid$(Msg, i, 1)) > 31 And Asc(Mid$(Msg, i, 1)) < 127) Or (Asc(Mid$(Msg, i, 1)) = 128) Or (Asc(Mid$(Msg, i, 1)) > 159 And Asc(Mid$(Msg, i, 1)) < 256)) And (Mid$(Msg, i, 1) <> SEP_CHAR And Mid$(Msg, i, 1) <> END_CHAR) Then
                    Call PlayerMsg(Index, "Uma ou mais caracteres digitadas são inválidas.", Red)
                    Exit Sub
                End If

            Next

            If frmServer.chkE.Value = Unchecked Then
                If GetPlayerAccess(Index) <= 0 Then
                    Call PlayerMsg(Index, "Mensagens emotivas foram desativadas pelo servidor!", BrightRed)
                    Exit Sub
                End If
            End If

            Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
            TextAdd frmServer.txtText(6), GetPlayerName(Index) & " " & Msg, True
            Exit Sub

        Case "broadcastmsg"
            Msg = Parse(1)

            ' Proteção
            For i = 1 To Len(Msg)

                If Not ((Asc(Mid$(Msg, i, 1)) > 31 And Asc(Mid$(Msg, i, 1)) < 127) Or (Asc(Mid$(Msg, i, 1)) = 128) Or (Asc(Mid$(Msg, i, 1)) > 159 And Asc(Mid$(Msg, i, 1)) < 256)) And (Mid$(Msg, i, 1) <> SEP_CHAR And Mid$(Msg, i, 1) <> END_CHAR) Then
                    Call PlayerMsg(Index, "Uma ou mais caracteres digitadas são inválidas.", Red)
                    Exit Sub
                End If

            Next

            If frmServer.chkBC.Value = Unchecked Then
                If GetPlayerAccess(Index) <= 0 Then
                    Call PlayerMsg(Index, "Mensagens globais foram desativadas!", BrightRed)
                    Exit Sub
                End If
            End If

            If Player(Index).Mute = True Then Exit Sub
            s = GetPlayerName(Index) & ": " & Msg
            Call AddLog(s, PLAYER_LOG)
            Call GlobalMsg(s, BroadcastColor)
            Call TextAdd(frmServer.txtText(0), s, True)
            TextAdd frmServer.txtText(1), GetPlayerName(Index) & ": " & Msg, True
            Exit Sub

        Case "globalmsg"
            Msg = Parse(1)

            ' Proteção
            For i = 1 To Len(Msg)

                If Not ((Asc(Mid$(Msg, i, 1)) > 31 And Asc(Mid$(Msg, i, 1)) < 127) Or (Asc(Mid$(Msg, i, 1)) = 128) Or (Asc(Mid$(Msg, i, 1)) > 159 And Asc(Mid$(Msg, i, 1)) < 256)) And (Mid$(Msg, i, 1) <> SEP_CHAR And Mid$(Msg, i, 1) <> END_CHAR) Then
                    Call PlayerMsg(Index, "Uma ou mais caracteres digitadas são inválidas.", Red)
                    Exit Sub
                End If

            Next

            If frmServer.chkG.Value = Unchecked Then
                If GetPlayerAccess(Index) <= 0 Then
                    Call PlayerMsg(Index, "Mensagens globais foram desativadas pelo servidor!", BrightRed)
                    Exit Sub
                End If
            End If

            If Player(Index).Mute = True Then Exit Sub
            If GetPlayerAccess(Index) > 0 Then
                s = "(global) " & GetPlayerName(Index) & ": " & Msg
                Call AddLog(s, ADMIN_LOG)
                Call GlobalMsg(s, GlobalColor)
                Call TextAdd(frmServer.txtText(0), s, True)
            End If

            TextAdd frmServer.txtText(2), GetPlayerName(Index) & ": " & Msg, True
            Exit Sub

        Case "adminmsg"
            Msg = Parse(1)

            ' Proteção
            For i = 1 To Len(Msg)

                If Not ((Asc(Mid$(Msg, i, 1)) > 31 And Asc(Mid$(Msg, i, 1)) < 127) Or (Asc(Mid$(Msg, i, 1)) = 128) Or (Asc(Mid$(Msg, i, 1)) > 159 And Asc(Mid$(Msg, i, 1)) < 256)) And (Mid$(Msg, i, 1) <> SEP_CHAR And Mid$(Msg, i, 1) <> END_CHAR) Then
                    Call PlayerMsg(Index, "Uma ou mais caracteres digitadas são inválidas.", Red)
                    Exit Sub
                End If

            Next

            If frmServer.chkA.Value = Unchecked Then
                Call PlayerMsg(Index, "Mensagens de admins foram desativadas pelo servidor!", BrightRed)
                Exit Sub
            End If

            If GetPlayerAccess(Index) > 0 Then
                Call AddLog("(Admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
                Call AdminMsg("(Admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
            End If

            TextAdd frmServer.txtText(5), GetPlayerName(Index) & ": " & Msg, True
            Exit Sub

        Case "playermsg"
            MsgTo = FindPlayer(Parse$(1))
            Msg = Parse$(2)

            ' Proteção
            For i = 1 To Len(Msg)

                If Not ((Asc(Mid$(Msg, i, 1)) > 31 And Asc(Mid$(Msg, i, 1)) < 127) Or (Asc(Mid$(Msg, i, 1)) = 128) Or (Asc(Mid$(Msg, i, 1)) > 159 And Asc(Mid$(Msg, i, 1)) < 256)) And (Mid$(Msg, i, 1) <> SEP_CHAR And Mid$(Msg, i, 1) <> END_CHAR) Then
                    Call PlayerMsg(Index, "Uma ou mais caracteres digitadas são inválidas.", Red)
                    Exit Sub
                End If

            Next

            'If frmServer.chkP.Value = Unchecked Then
            '    If GetPlayerAccess(Index) <= 0 Then
            '        Call PlayerMsg(Index, "Mensagens Privadas foram desativadas pelo servidor!", BrightRed)
            '        Exit Sub
            '    End If
            'End If

            ' Veja se ele está tentando falar com si mesmo!
            If MsgTo <> Index Then
                If MsgTo > 0 Then
                    Call AddLog(GetPlayerName(Index) & " fala para " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                    Call PlayerMsg(MsgTo, GetPlayerName(Index) & ":" & Msg, TellColor)
                    Call PlayerMsg(Index, "Você diz para " & GetPlayerName(MsgTo) & ": " & Msg, TellColor)
                Else
                    Call PlayerMsg(Index, Parse$(1) & " não está online.", White)
                    Exit Sub
                End If

            Else
                Call AddLog("Mapa #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " começa a falar algo para ele mesmo...", PLAYER_LOG)
                Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " começa a falar algo para ele mesmo...", Green)
            End If

            TextAdd frmServer.txtText(4), "De " & GetPlayerName(MsgTo) & " para " & GetPlayerName(Index) & ": " & Msg, True
            Exit Sub

            ' :::::::::::::::::::::::::::::::::
            ' :: Packet de Movimento do Char ::
            ' :::::::::::::::::::::::::::::::::
        Case "playermove"

            If Player(Index).GettingMap = YES Then
                Exit Sub
            End If

            Dir = Val(Parse(1))
            Movement = Val(Parse(2))

            ' Proteção
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(Index, "Direção Inválida")
                Exit Sub
            End If

            ' Proteção
            If Movement < 1 Or Movement > 2 Then
                Call HackingAttempt(Index, "Movimento Inválido")
                Exit Sub
            End If

            ' Se o jogador usou uma magia, ele não pode andar
            If Player(Index).CastedSpell = YES Then

                ' Não vamos deixar ele se mover
                If GetTickCount > Player(Index).AttackTimer + 1000 Then
                    Player(Index).CastedSpell = NO
                Else
                    Call SendPlayerXY(Index)
                    Exit Sub
                End If
            End If

            Call PlayerMove(Index, Dir, Movement)
            Exit Sub

            ' :::::::::::::::::::::::::::::::::
            ' :: Packet de Movimento do Char ::
            ' :::::::::::::::::::::::::::::::::
        Case "playerdir"

            If Player(Index).GettingMap = YES Then
                Exit Sub
            End If

            Dir = Val(Parse(1))

            ' Proteção
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(Index, "Direção Inválida")
                Exit Sub
            End If

            Call SetPlayerDir(Index, Dir)
            Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::
            ' :: Use item packet ::
            ' :::::::::::::::::::::
        Case "useitem"
            InvNum = Val(Parse(1))
            CharNum = Player(Index).CharNum

            ' Proteção
            If InvNum < 1 Or InvNum > MAX_ITEMS Then
                Call HackingAttempt(Index, "InvNum Inválido")
                Exit Sub
            End If

            ' Proteção
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "CharNum Inválido")
                Exit Sub
            End If

            If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
                N = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
                Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long, n6 As Long

                n1 = Item(GetPlayerInvItemNum(Index, InvNum)).StrReq
                n2 = Item(GetPlayerInvItemNum(Index, InvNum)).DefReq
                n3 = Item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq
                n6 = Item(GetPlayerInvItemNum(Index, InvNum)).MagicReq
                n4 = Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq
                n5 = Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq

                ' Que tipo de item é?
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type

                    Case ITEM_TYPE_ARMOR

                        If InvNum <> GetPlayerArmorSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "Você precisa ser classe " & GetClassName(n4) & " para usar esse item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Seu acesso precisa ser maior que " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Sua força é muito baixa para equipar esse item! Força requerida: (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Sua defesa é muito baixa para equipar esse item! Defesa requerida: (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Sua velocidade é muito baixa para equipar esse item! Velocidade requerida: (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Sua magia é muito baixa para equipar esse item! Magia requerida: (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerArmorSlot(Index, InvNum)
                        Else
                            Call SetPlayerArmorSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_WEAPON

                        If InvNum <> GetPlayerWeaponSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "VocÊ precisa ser classe " & GetClassName(n4) & " para usar esse item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Seu acesso precisa ser maior que " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Sua força é muito baixa para equipar esse item! Força requerida: (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Sua defesa é muito baixa para equipar esse item! Defesa requerida: (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Sua velocidade é muito baixa para equipar esse item! Velocidade requerida: (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Sua magia é muito baixa para equipar esse item! Magia requerida: (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerWeaponSlot(Index, InvNum)
                        Else
                            Call SetPlayerWeaponSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_HELMET

                        If InvNum <> GetPlayerHelmetSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "Você precisa ser classe " & GetClassName(n4) & " para usar esse item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Seu acesso precisa ser maior que " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Sua força é muito baixa para equipar esse item! Força requerida: (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Sua defesa é muito baixa para equipar esse item! Defesa requerida: (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Sua velocidade é muito baixa para equipar esse item! Velocidade requerida: (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Sua magia é muito baixa para equipar esse item! Magia requerida: (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerHelmetSlot(Index, InvNum)
                        Else
                            Call SetPlayerHelmetSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_SHIELD

                        If InvNum <> GetPlayerShieldSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "Você precisa ser classe " & GetClassName(n4) & " para usar esse item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Seu acesso precisa ser maior que " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Sua força é muito baixa para equipar esse item! Força requerida:(" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Sua defesa é muito baixa para equipar esse item! Defesa requerida: (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Sua velocidade é muito baixa para equipar esse item! Velocidade requerida: (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Sua magia é muito baixa para equipar esse item! Magia requerida: (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerShieldSlot(Index, InvNum)
                        Else
                            Call SetPlayerShieldSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_POTIONADDHP
                        Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                        Call SendHP(Index)

                    Case ITEM_TYPE_POTIONADDMP
                        Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                        Call SendMP(Index)

                    Case ITEM_TYPE_POTIONADDSP
                        Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                        Call SendSP(Index)

                    Case ITEM_TYPE_POTIONSUBHP
                        Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                        Call SendHP(Index)

                    Case ITEM_TYPE_POTIONSUBMP
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                        Call SendMP(Index)

                    Case ITEM_TYPE_POTIONSUBSP
                        Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                        Call SendSP(Index)

                    Case ITEM_TYPE_KEY
                        x = DirToX(GetPlayerX(Index), GetPlayerDir(Index))
                        y = DirToY(GetPlayerY(Index), GetPlayerDir(Index))

                        If Not IsValid(x, y) Then Exit Sub

                        ' Veja se a chave existe
                        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                            ' Verificar se a chave é para o mapa
                            If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                                TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)

                                If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                                    Call MapMsg(GetPlayerMap(Index), "A porta foi destrancada!", White)
                                Else
                                    Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
                                End If

                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Chave" & END_CHAR)

                                ' Verificar se devemos tomar a chave
                                If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                    Call PlayerMsg(Index, "A chave desaparece.", Yellow)
                                End If
                            End If
                        End If

                        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Chave" & END_CHAR)
                        End If

                    Case ITEM_TYPE_SPELL

                        ' Pegar o número da magia
                        N = Item(GetPlayerInvItemNum(Index, InvNum)).Data1

                        If N > 0 Then

                            ' Classe correta?
                            If Spell(N).ClassReq = GetPlayerClass(Index) Or Spell(N).ClassReq = 0 Then
                                If Spell(N).LevelReq = 0 And Player(Index).Char(Player(Index).CharNum).Access < 1 Then
                                    Call PlayerMsg(Index, "Essa magia pode ser usada apenas por administradores!", BrightRed)
                                    Exit Sub
                                End If

                                ' Level Certo?
                                i = GetSpellReqLevel(N)

                                If n6 > i Then i = n6
                                If i <= GetPlayerLevel(Index) Then
                                    i = FindOpenSpellSlot(Index)

                                    ' Slot de magia aberto?
                                    If i > 0 Then

                                        ' Já tem a magia?
                                        If Not HasSpell(Index, N) Then
                                            Call SetPlayerSpell(Index, i, N)
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                            Call PlayerMsg(Index, "Você estuda a magia...", Yellow)
                                            Call PlayerMsg(Index, "Você aprendeu uma nova magia!", White)
                                        Else
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                            Call PlayerMsg(Index, "Você já aprendeu essa magia! O pergaminho se esvai .", BrightRed)
                                        End If

                                    Else
                                        Call PlayerMsg(Index, "Você já aprendeu tudo que podia!", BrightRed)
                                    End If

                                Else
                                    Call PlayerMsg(Index, "Você precisa ser level " & i & " para aprender essa magia.", White)
                                End If

                            Else
                                Call PlayerMsg(Index, "Essa magia pode ser aprendida apenas por um " & GetClassName(Spell(N).ClassReq) & ".", White)
                            End If

                        Else
                            Call PlayerMsg(Index, "Esse pergaminho não está associado com uma magia! Contate um GM!", White)
                        End If

                    Case ITEM_TYPE_PET
                        Player(Index).Pet.Alive = YES
                        Player(Index).Pet.Sprite = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                        Player(Index).Pet.Dir = DIR_UP
                        Player(Index).Pet.Map = GetPlayerMap(Index)
                        Player(Index).Pet.MapToGo = 0
                        Player(Index).Pet.x = GetPlayerX(Index) + Int(Rnd * 3 - 1)

                        If Player(Index).Pet.x < 0 Or Player(Index).Pet.x > MAX_MAPX Then Player(Index).Pet.x = GetPlayerX(Index)
                        Player(Index).Pet.XToGo = -1
                        Player(Index).Pet.y = GetPlayerY(Index) + Int(Rnd * 3 - 1)

                        If Player(Index).Pet.y < 0 Or Player(Index).Pet.y > MAX_MAPY Then Player(Index).Pet.y = GetPlayerY(Index)
                        Player(Index).Pet.YToGo = -1
                        Player(Index).Pet.Level = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
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
                        Call SendDataToMap(GetPlayerMap(Index), Packet)

                        ' PRESSA! Desculpe-me pelo código feio!
                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                        Call PlayerMsg(Index, "Você conseguiu um pet!", White)
                End Select

                Call SendStats(Index)
                Call SendHP(Index)
                Call SendMP(Index)
                Call SendSP(Index)
            End If

            Exit Sub

            ' ::::::::::::::::::::::::::
            ' ::   Packet de Ataque   ::
            ' ::::::::::::::::::::::::::
        Case "attack"

            If GetPlayerWeaponSlot(Index) > 0 Then
                If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 > 0 Then
                    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
                    Exit Sub
                End If
            End If

            ' Atacando um jogador...
            For i = 1 To MAX_PLAYERS

                ' Não podemos nos atacar
                If i <> Index Then

                    ' Podemos atacar um jogador?
                    If CanAttackPlayer(Index, i) Then
                        If Not CanPlayerBlockHit(i) Then

                            ' Pegando informação do dano...
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(i) + (Rnd * 5) - 2
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque" & Int(Rnd * 2) + 1 & END_CHAR)
                            Else
                                N = GetPlayerDamage(Index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(i) + (Rnd * 5) - 2
                                Call BattleMsg(Index, "Você sente uma enorme quantidade de energia em seu corpo!", BrightCyan, 0)
                                Call BattleMsg(i, GetPlayerName(Index) & " move-se com uma enorme destreza!", BrightCyan, 1)

                                'Call PlayerMsg(index, "Você sente uma enorme quantidade de energia em seu corpo!", BrightCyan)
                                'Call PlayerMsg(I, GetPlayerName(index) & " move-se com uma enorme destreza!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque3" & END_CHAR)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(Index, i, Damage)
                            Else
                                Call PlayerMsg(Index, "Seu ataque não fez nada.", BrightRed)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Errou" & END_CHAR)
                            End If

                        Else
                            Call BattleMsg(Index, GetPlayerName(i) & " bloqueou seu ataque!", BrightCyan, 0)
                            Call BattleMsg(i, "Você bloqueou o ataque de " & GetPlayerName(Index) & "!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " bloqueou seu ataque!", BrightCyan)
                            'Call PlayerMsg(I, "Your " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " bloqueou o ataque de " & GetPlayerName(index) & "!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Errou" & END_CHAR)
                        End If

                        Exit Sub
                    End If
                End If

            Next

            ' Atacando NPC
            For i = 1 To MAX_MAP_NPCS

                ' Podemos atacar?
                If CanAttackNpc(Index, i) Then

                    ' Pegando o dano...
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), i).num).DEF / 2) + (Rnd * 5) - 2
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque" & Int(Rnd * 2) + 1 & END_CHAR)
                    Else
                        N = GetPlayerDamage(Index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(Index, "Você sente uma enorme quantidade de energia em seu corpo!", BrightCyan, 0)

                        'Call PlayerMsg(index, "Você sente uma grande quantidade de energia em seu corpo!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque3" & END_CHAR)
                    End If

                    If Damage > 0 Then
                        Call AttackNpc(Index, i, Damage)
                        Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & END_CHAR)
                    Else
                        Call BattleMsg(Index, "Seu ataque não fez nada.", BrightRed, 0)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Errou" & END_CHAR)
                    End If

                    Exit Sub
                End If

            Next

            Exit Sub

            ' :::::::::::::::::::::::::
            ' :: Status Point Packet ::
            ' :::::::::::::::::::::::::
        Case "usestatpoint"
        
            ' Proteção contra erros e packets editadas
            If Val(Parse(1)) > 3 Or Val(Parse(1)) < 0 Then Exit Sub
            If Val(Parse(2)) > 1000 Or Val(Parse(2)) = 0 Then Exit Sub
            
            PointType = Val(Parse(1))
            PointQuant = Val(Parse(2))

            ' Verificar se temos pontos necessários
            If GetPlayerPOINTS(Index) >= PointQuant Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Principal.txt", "UsingStatPoints " & Index & "," & PointType & "," & PointQuant
                Else

                    Select Case PointType

                        Case 0
                            Call SetPlayerstr(Index, GetPlayerstr(Index) + PointQuant)
                            Call BattleMsg(Index, "Você adicionou " & PointQuant & " ponto(s) em Força!", 15, 0)
                            Call BattleMsg(Index, "Você ainda possui " & GetPlayerPOINTS(Index) - PointQuant & " pontos para serem gastos.", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "ForSubiu" & END_CHAR)

                        Case 1
                            Call SetPlayerDEF(Index, GetPlayerDEF(Index) + PointQuant)
                            Call BattleMsg(Index, "Você adicionou " & PointQuant & " ponto(s) em Defesa!", 15, 0)
                            Call BattleMsg(Index, "Você ainda possui " & GetPlayerPOINTS(Index) - PointQuant & " pontos para serem gastos.", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "DefSubiu" & END_CHAR)

                        Case 2
                            Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + PointQuant)
                            Call BattleMsg(Index, "Você adicionou " & PointQuant & " ponto(s) em Inteligência!", 15, 0)
                            Call BattleMsg(Index, "Você ainda possui " & GetPlayerPOINTS(Index) - PointQuant & " pontos para serem gastos.", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "IntSubiu" & END_CHAR)

                        Case 3
                            Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + PointQuant)
                            Call BattleMsg(Index, "Você adicionou " & PointQuant & " ponto(s) em Agilidade!", 15, 0)
                            Call BattleMsg(Index, "Você ainda possui " & GetPlayerPOINTS(Index) - PointQuant & " pontos para serem gastos.", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "AgiSubiu" & END_CHAR)
                    End Select

                    Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - PointQuant)
                End If

            Else
                Call BattleMsg(Index, "Você tem apenas " & GetPlayerPOINTS(Index) & " ponto(s) para gastar.", BrightRed, 0)
            End If

            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendStats(Index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Informação de um Jogador ::
            ' ::::::::::::::::::::::::::::::
        Case "playerinforequest"
            Name = Parse(1)
            i = FindPlayer(Name)

            If i > 0 Then
                Call PlayerMsg(Index, "Conta: " & Trim$(Player(i).Login) & ", Nome: " & GetPlayerName(i), BrightGreen)

                If GetPlayerAccess(Index) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "-=- Status de " & GetPlayerName(i) & " -=-", BrightGreen)
                    Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                    Call PlayerMsg(Index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                    Call PlayerMsg(Index, "FOR: " & GetPlayerstr(i) & "  DEF: " & GetPlayerDEF(i) & "  MAG: " & GetPlayerMAGI(i) & "  AGI: " & GetPlayerSPEED(i), BrightGreen)
                    N = Int(GetPlayerstr(i) / 2) + Int(GetPlayerLevel(i) / 2)
                    i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)

                    If N > 100 Then N = 100
                    If i > 100 Then i = 100
                    Call PlayerMsg(Index, "Chance de Dano Crítico: " & N & "%, Chance de Bloqueio: " & i & "%", BrightGreen)
                End If

            Else
                Call PlayerMsg(Index, "O jogador não está online.", White)
            End If

            Exit Sub

            ' ::::::::::::::::::::
            ' ::  Sprite Packet ::
            ' ::::::::::::::::::::
        Case "setsprite"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' A sprite
            N = Val(Parse(1))
            Call SetPlayerSprite(Index, N)
            Call SendPlayerData(Index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Set player sprite packet ::
            ' ::::::::::::::::::::::::::::::
        Case "setplayersprite"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' A sprite
            i = FindPlayer(Parse(1))
            N = Val(Parse(2))
            Call SetPlayerSprite(i, N)
            Call SendPlayerData(i)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::
            ' :: Packet de Pedido de Status ::
            ' ::::::::::::::::::::::::::::::::
        Case "getstats"
            Call PlayerMsg(Index, "-=- Status de " & GetPlayerName(Index) & " -=-", White)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
            Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
            Call PlayerMsg(Index, "FOR: " & GetPlayerstr(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAG: " & GetPlayerMAGI(Index) & "  VEL: " & GetPlayerSPEED(Index), White)
            N = Int(GetPlayerstr(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)

            If N > 100 Then N = 100
            If i > 100 Then i = 100
            Call PlayerMsg(Index, "Chance de Dano Crítico: " & N & "%, Chance de Bloqueio: " & i & "%", White)
            Exit Sub

            ' :::::::::::::::::::::::::::
            ' :: Pedido para novo mapa ::
            ' :::::::::::::::::::::::::::
        Case "requestnewmap"
            Dir = Val(Parse(1))

            ' Proteção
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(Index, "Direção Inválida")
                Exit Sub
            End If

            Call PlayerMove(Index, Dir, 1)
            Exit Sub

            ' ::::::::::::::::::
            ' :: Info do Mapa ::
            ' ::::::::::::::::::
        Case "mapdata"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = 1
            MapNum = GetPlayerMap(Index)
            Call ClearMap(MapNum)
            Map(MapNum).Name = Parse(N + 1)
            Map(MapNum).Revision = Val(Parse(N + 2)) + 1
            Map(MapNum).Moral = Val(Parse(N + 3))
            Map(MapNum).Up = Val(Parse(N + 4))
            Map(MapNum).Down = Val(Parse(N + 5))
            Map(MapNum).Left = Val(Parse(N + 6))
            Map(MapNum).Right = Val(Parse(N + 7))
            Map(MapNum).Music = Parse(N + 8)
            Map(MapNum).BootMap = Val(Parse(N + 9))
            Map(MapNum).BootX = Val(Parse(N + 10))
            Map(MapNum).BootY = Val(Parse(N + 11))
            Map(MapNum).Indoors = Val(Parse(N + 12))
            N = N + 13
            i = GetPlayerMap(Index)

            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Ground = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).GroundSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Mask = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).MaskSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Fringe = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).FringeSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Type = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Data1 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Data2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Data3 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).String1 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).String2 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).String3 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Mask2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Mask2Set = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).M2Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).M2AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).FAnim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).FAnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Fringe2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Fringe2Set = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).Light = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).F2Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(x, y).F2AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    N = N + 1
                Next
            Next

            For x = 1 To MAX_MAP_NPCS
                Map(MapNum).Npc(x) = Val(Parse(N))
                Map(MapNum).NpcSpawn(x).Used = Val(Parse(N + 1))
                Map(MapNum).NpcSpawn(x).x = Val(Parse(N + 2))
                Map(MapNum).NpcSpawn(x).y = Val(Parse(N + 3))
                N = N + 4
                Call ClearMapNpc(x, MapNum)
            Next

            ' Limpar tudo
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
                Call ClearMapItem(i, GetPlayerMap(Index))
            Next

            ' Salvar o mapa
            Call SaveMap(MapNum)

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(Index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(Index))

            ' Resetar grid
            Call ResetMapGrid(GetPlayerMap(Index))

            ' Atualizar mapa para todos
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                    'Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)
                    Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i), False)
                End If

            Next

            Exit Sub

            ' ::::::::::::::::::::::::::::
            ' :: packet do mapa sim/não ::
            ' ::::::::::::::::::::::::::::
        Case "needmap"

            ' Get yes/no value
            s = LCase$(Parse(1))

            If s = "yes" Then
                Call SendMap(Index, GetPlayerMap(Index))
                Call SendMapItemsTo(Index, GetPlayerMap(Index))
                Call SendMapNpcsTo(Index, GetPlayerMap(Index))
                Call SendJoinMap(Index)
                Player(Index).GettingMap = NO
                Call SendDataTo(Index, "MAPDONE" & END_CHAR)
            Else
                Call SendMapItemsTo(Index, GetPlayerMap(Index))
                Call SendMapNpcsTo(Index, GetPlayerMap(Index))
                Call SendJoinMap(Index)
                Player(Index).GettingMap = NO
                Call SendDataTo(Index, "MAPDONE" & END_CHAR)
            End If

            Exit Sub

            ' ::::::::::::::::::::::::::::::::::::::::::::
            ' :: Packet do jogador que tenta pegar algo ::
            ' ::::::::::::::::::::::::::::::::::::::::::::
        Case "mapgetitem"
            Call PlayerMapGetItem(Index)
            Exit Sub

            ' :::::::::::::::::::::::::::::::::::::::::::::
            ' :: Packet do jogador que tenta dropar algo ::
            ' :::::::::::::::::::::::::::::::::::::::::::::
        Case "mapdropitem"
            InvNum = Val(Parse(1))
            Amount = Val(Parse(2))

            ' Proteção
            If InvNum < 1 Or InvNum > MAX_INV Then
                Call HackingAttempt(Index, "InvNum Inválido")
                Exit Sub
            End If

            ' Proteção
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then

                ' Verificar se é dinheiro, se for, não deixar que dropem um valor igual a 0 (ou menor)
                If Amount <= 0 Then
                    Call PlayerMsg(Index, "Você deve dropar algo maior que 0!", BrightRed)
                    Exit Sub
                End If

                If Amount > GetPlayerInvItemValue(Index, InvNum) Then
                    Call PlayerMsg(Index, "Você não tem tudo isso para dropar!", BrightRed)
                    Exit Sub
                End If
            End If

            ' Proteção
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type <> ITEM_TYPE_CURRENCY Then
                If Amount > GetPlayerInvItemValue(Index, InvNum) Then
                    Call HackingAttempt(Index, "Modificação de quantidade de item")
                    Exit Sub
                End If
            End If

            Call PlayerMapDropItem(Index, InvNum, Amount)
            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Packet de Respawn ::
            ' :::::::::::::::::::::::
        Case "maprespawn"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' Limpar tudo
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
                Call ClearMapItem(i, GetPlayerMap(Index))
            Next

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(Index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(Index))

            ' Resetar grid
            Call ResetMapGrid(GetPlayerMap(Index))
            Call PlayerMsg(Index, "Mapa respawnado.", Blue)
            Call AddLog(GetPlayerName(Index) & " apareceu no mapa #" & GetPlayerMap(Index), ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::
            ' :: Kick player packet ::
            ' ::::::::::::::::::::::::
        Case "kickplayer"

            ' Proteção
            If GetPlayerAccess(Index) <= 0 Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' player index
            N = FindPlayer(Parse(1))

            If N <> Index Then
                If N > 0 Then
                    If GetPlayerAccess(N) <= GetPlayerAccess(Index) Then
                        Call AddLog(GetPlayerName(Index) & " expulsou " & GetPlayerName(N) & ".", ADMIN_LOG)
                        Call AlertMsg(N, "Você foi expulso por " & GetPlayerName(Index) & "!")
                    Else
                        Call PlayerMsg(Index, "Ele é um administrador de nível maior que o seu!", White)
                    End If

                Else
                    Call PlayerMsg(Index, "O jogador está offline.", White)
                End If

            Else
                Call PlayerMsg(Index, "Você não pode se expulsar!", White)
            End If

            Exit Sub

            ' :::::::::::::::::::::
            ' :: Ban list packet ::
            ' :::::::::::::::::::::
        Case "banlist"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = 1
            f = FreeFile
            Open App.Path & "\banlist.txt" For Input As #f

            Do While Not EOF(f)
                Input #f, s
                Input #f, Name
                Call PlayerMsg(Index, N & ": IP banido " & s & " por " & Name, White)
                N = N + 1
            Loop

            Close #f
            Exit Sub

            ' ::::::::::::::::::::::::
            ' :: Ban destroy packet ::
            ' ::::::::::::::::::::::::
        Case "bandestroy"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call Kill(App.Path & "\banlist.txt")
            Call PlayerMsg(Index, "Lista de Ban destruída.", White)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Ban player packet ::
            ' :::::::::::::::::::::::
        Case "banplayer"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' player index
            N = FindPlayer(Parse(1))

            If N <> Index Then
                If N > 0 Then
                    If GetPlayerAccess(N) <= GetPlayerAccess(Index) Then
                        Call BanIndex(N, Index)
                    Else
                        Call PlayerMsg(Index, "Ele é um administrador de nível maior que o seu!", White)
                    End If

                Else
                    Call PlayerMsg(Index, "O jogador está offline.", White)
                End If

            Else
                Call PlayerMsg(Index, "Você não pode se banir!", White)
            End If

            Exit Sub

            ' :::::::::::::::::::::::::::::::
            ' :: Packet de Edição de mapas ::
            ' :::::::::::::::::::::::::::::::
        Case "requesteditmap"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "EDITMAP" & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::::::::::::::::::::::
            ' :: Packet de Pedido de Edição de Itens ::
            ' :::::::::::::::::::::::::::::::::::::::::
        Case "requestedititem"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "ITEMEDITOR" & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::
            ' :: Editar itens ::
            ' ::::::::::::::::::
        Case "edititem"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' O número do item
            N = Val(Parse(1))

            ' Proteção
            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Item Inválido")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editando Item #" & N & ".", ADMIN_LOG)
            Call SendEditItemTo(Index, N)
            Exit Sub

            ' :::::::::::::::::
            ' :: Salvar item ::
            ' :::::::::::::::::
        Case "saveitem"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If

            ' Atualizar o item
            Item(N).Name = Parse(2)
            Item(N).Pic = Val(Parse(3))
            Item(N).Type = Val(Parse(4))
            Item(N).Data1 = Val(Parse(5))
            Item(N).Data2 = Val(Parse(6))
            Item(N).Data3 = Val(Parse(7))
            Item(N).StrReq = Val(Parse(8))
            Item(N).DefReq = Val(Parse(9))
            Item(N).SpeedReq = Val(Parse(10))
            Item(N).MagicReq = Val(Parse(11))
            Item(N).ClassReq = Val(Parse(12))
            Item(N).AccessReq = Val(Parse(13))
            Item(N).AddHP = Val(Parse(14))
            Item(N).AddMP = Val(Parse(15))
            Item(N).AddSP = Val(Parse(16))
            Item(N).AddStr = Val(Parse(17))
            Item(N).AddDef = Val(Parse(18))
            Item(N).AddMagi = Val(Parse(19))
            Item(N).AddSpeed = Val(Parse(20))
            Item(N).AddEXP = Val(Parse(21))
            Item(N).Desc = Parse(22)
            Item(N).AttackSpeed = Val(Parse(23))

            ' Save it
            Call SendUpdateItemToAll(N)
            Call SaveItem(N)
            Call AddLog(GetPlayerName(Index) & " salvou Item #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::::::::
            ' ::Packet de Pedido de Edição de NPC ::
            ' ::::::::::::::::::::::::::::::::::::::
        Case "requesteditnpc"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "NPCEDITOR" & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Packet de Edição de npc  ::
            ' ::::::::::::::::::::::::::::::
        Case "editnpc"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' O número do NPC
            N = Val(Parse(1))

            ' Proteção
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid NPC Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editando NPC #" & N & ".", ADMIN_LOG)
            Call SendEditNpcTo(Index, N)
            Exit Sub

            ' ::::::::::::::::::::::::::::
            ' :: Packet para Salvar npc ::
            ' ::::::::::::::::::::::::::::
        Case "savenpc"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            ' Proteção
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(Index, "NPC Inválido")
                Exit Sub
            End If

            ' Update the npc
            Npc(N).Name = Parse(2)
            Npc(N).AttackSay = Parse(3)
            Npc(N).Sprite = Val(Parse(4))
            Npc(N).SpawnSecs = Val(Parse(5))
            Npc(N).Behavior = Val(Parse(6))
            Npc(N).Range = Val(Parse(7))
            Npc(N).STR = Val(Parse(8))
            Npc(N).DEF = Val(Parse(9))
            Npc(N).Speed = Val(Parse(10))
            Npc(N).Magi = Val(Parse(11))
            Npc(N).Big = Val(Parse(12))
            Npc(N).MaxHp = Val(Parse(13))
            Npc(N).Exp = Val(Parse(14))
            Npc(N).SpawnTime = Val(Parse(15))
            Npc(N).Speech = Val(Parse(16))
            z = 17

            For i = 1 To MAX_NPC_DROPS
                Npc(N).ItemNPC(i).Chance = Val(Parse(z))
                Npc(N).ItemNPC(i).ItemNum = Val(Parse(z + 1))
                Npc(N).ItemNPC(i).ItemValue = Val(Parse(z + 2))
                z = z + 3
            Next

            ' Salvar
            Call SendUpdateNpcToAll(N)
            Call SaveNpc(N)
            Call AddLog(GetPlayerName(Index) & " salvou NPC #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::::::::::
            ' :: Packet de Pedido de Edição de Shop ::
            ' ::::::::::::::::::::::::::::::::::::::::
        Case "requesteditshop"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "SHOPEDITOR" & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Packet de Edição de Shop ::
            ' ::::::::::::::::::::::::::::::
        Case "editshop"

            ' Proteção
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' O número do Shop
            N = Val(Parse(1))

            ' Proteção
            If N < 0 Or N > MAX_SHOPS Then
                Call HackingAttempt(Index, "Loja Inválida")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editando Loja #" & N & ".", ADMIN_LOG)
            Call SendEditShopTo(Index, N)
            Exit Sub

        Case "addfriend"
            Name = Trim$(Parse(1))

            If Not FindChar(Name) Then
                Call PlayerMsg(Index, "Esse jogador não existe!", Blue)
                Exit Sub
            End If

            If Name = GetPlayerName(Index) Then
                Call PlayerMsg(Index, "Você não pode se adicionar!", Blue)
                Exit Sub
            End If

            For i = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(i) = Name Then
                    Call PlayerMsg(Index, "Você já tem esse jogador como amigo!", Blue)
                    Exit Sub
                End If

            Next

            For i = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(i) = vbNullString Then
                    Player(Index).Char(Player(Index).CharNum).Friends(i) = Name
                    Call PlayerMsg(Index, "Amigo adicionado.", Blue)
                    Call SendFriendListTo(Index)
                    Exit Sub
                End If

            Next

            Call PlayerMsg(Index, "Desculpe, mas você já tem muitos amigos.", Blue)
            Exit Sub

        Case "removefriend"
            Name = Trim$(Parse(1))

            For i = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(i) = Name Then
                    Player(Index).Char(Player(Index).CharNum).Friends(i) = vbNullString
                    Call PlayerMsg(Index, "Amigo removido.", Blue)
                    Call SendFriendListTo(Index)
                    Exit Sub
                End If

            Next

            Call PlayerMsg(Index, "Essa pessoa não está na sua lista de amigos!", Blue)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Save shop packet ::
            ' ::::::::::::::::::::::
        Case "saveshop"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ShopNum = Val(Parse(1))

            ' Prevent hacking
            If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
                Call HackingAttempt(Index, "Loja Inválida")
                Exit Sub
            End If

            ' Update the shop
            Shop(ShopNum).Name = Parse(2)
            Shop(ShopNum).JoinSay = Parse(3)
            Shop(ShopNum).LeaveSay = Parse(4)
            Shop(ShopNum).FixesItems = Val(Parse(5))
            N = 6

            For z = 1 To 6
                For i = 1 To MAX_TRADES
                    Shop(ShopNum).TradeItem(z).Value(i).GiveItem = Val(Parse(N))
                    Shop(ShopNum).TradeItem(z).Value(i).GiveValue = Val(Parse(N + 1))
                    Shop(ShopNum).TradeItem(z).Value(i).GetItem = Val(Parse(N + 2))
                    Shop(ShopNum).TradeItem(z).Value(i).GetValue = Val(Parse(N + 3))
                    N = N + 4
                Next
            Next

            ' Save it
            Call SendUpdateShopToAll(ShopNum)
            Call SaveShop(ShopNum)
            Call AddLog(GetPlayerName(Index) & " salvou Loja #" & ShopNum & ".", ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit main packet ::
            ' ::::::::::::::::::::::::::::::
        Case "requesteditmain"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            f = FreeFile
            Open App.Path & "\Scripts\Principal.txt" For Input As #f
            Call SendDataTo(Index, "MAINEDITOR" & SEP_CHAR & Input$(LOF(f), f) & END_CHAR)
            Close #f
            Exit Sub

            ' :::::::::::::::::::::::::::::::
            ' :: Request edit spell packet ::
            ' :::::::::::::::::::::::::::::::
        Case "requesteditspell"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "SPELLEDITOR" & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Edit spell packet ::
            ' :::::::::::::::::::::::
        Case "editspell"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The spell #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(Index, "Magia Inválida")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editando Magia #" & N & ".", ADMIN_LOG)
            Call SendEditSpellTo(Index, N)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Save spell packet ::
            ' :::::::::::::::::::::::
        Case "savespell"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' Spell #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If

            ' Update the spell
            Spell(N).Name = Parse(2)
            Spell(N).ClassReq = Val(Parse(3))
            Spell(N).LevelReq = Val(Parse(4))
            Spell(N).Type = Val(Parse(5))
            Spell(N).Data1 = Val(Parse(6))
            Spell(N).Data2 = Val(Parse(7))
            Spell(N).Data3 = Val(Parse(8))
            Spell(N).MPCost = Val(Parse(9))
            Spell(N).sound = Val(Parse(10))
            Spell(N).Range = Val(Parse(11))
            Spell(N).SpellAnim = Val(Parse(12))
            Spell(N).SpellTime = Val(Parse(13))
            Spell(N).SpellDone = Val(Parse(14))
            Spell(N).AE = Val(Parse(15))

            ' Save it
            Call SendUpdateSpellToAll(N)
            Call SaveSpell(N)
            Call AddLog(GetPlayerName(Index) & " salvou Magia #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Set access packet ::
            ' :::::::::::::::::::::::
        Case "setaccess"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Call HackingAttempt(Index, "Tentando usar poderes que você não possui?")
                Exit Sub
            End If

            ' The index
            N = FindPlayer(Parse(1))

            ' The access
            i = Val(Parse(2))

            ' Check for invalid access level
            If i >= 0 Or i <= 3 Then
                If GetPlayerName(Index) <> GetPlayerName(N) Then
                    If GetPlayerAccess(Index) > GetPlayerAccess(N) Then

                        ' Check if player is on
                        If N > 0 Then
                            If GetPlayerAccess(N) <= 0 Then
                                Call GlobalMsg(GetPlayerName(N) & " foi abençoado com acesso administrativo.", BrightBlue)
                            End If

                            Call SetPlayerAccess(N, i)
                            Call SendPlayerData(N)
                            Call AddLog(GetPlayerName(Index) & " modificou o acesso de " & GetPlayerName(N) & ".", ADMIN_LOG)
                        Else
                            Call PlayerMsg(Index, GetPlayerName(N) & " não está online.", White)
                        End If

                    Else
                        Call PlayerMsg(Index, "Seu acesso é menor que o de " & GetPlayerName(N) & ".", Red)
                    End If

                Else
                    Call PlayerMsg(Index, "Você não mode mudar seu acesso.", Red)
                End If

            Else
                Call PlayerMsg(Index, "Nível de acesso inválido.", Red)
            End If

            Exit Sub

        Case "whosonline"
            Call SendWhosOnline(Index)
            Exit Sub

        Case "onlinelist"
            Call SendOnlineList
            Exit Sub

        Case "setmotd"

            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SpecialPutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
            Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
            Call AddLog(GetPlayerName(Index) & " mudou o MOTD para: " & Parse(1), ADMIN_LOG)
            Exit Sub

        Case "traderequest"

            ' Trade num
            N = Val(Parse(1))
            z = Val(Parse(2))

            ' Prevent hacking
            If (N < 1) Or (N > 6) Then
                Call HackingAttempt(Index, "Trade Request Modification")
                Exit Sub
            End If

            ' Prevent hacking
            If (z <= 0) Or (z > (MAX_TRADES * 6)) Then
                Call HackingAttempt(Index, "Trade Request Modification")
                Exit Sub
            End If

            ' Index for shop
            ' I = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
            i = Val(Parse(3))

            ' Check if inv full
            If i <= 0 Then Exit Sub
            x = FindOpenInvSlot(Index, Shop(i).TradeItem(N).Value(z).GetItem)

            If x = 0 Then
                Call PlayerMsg(Index, "Nada feito, seu inventário está cheio.", BrightRed)
                Exit Sub
            End If

            ' Check if they have the item
            If HasItem(Index, Shop(i).TradeItem(N).Value(z).GiveItem) >= Shop(i).TradeItem(N).Value(z).GiveValue Then
                Call TakeItem(Index, Shop(i).TradeItem(N).Value(z).GiveItem, Shop(i).TradeItem(N).Value(z).GiveValue)
                Call GiveItem(Index, Shop(i).TradeItem(N).Value(z).GetItem, Shop(i).TradeItem(N).Value(z).GetValue)
                Call PlayerMsg(Index, "A troca foi efetuada com sucesso!", Yellow)
            Else
                Call PlayerMsg(Index, "Algo deu errado na negociação.", BrightRed)
            End If

            Exit Sub

        Case "fixitem"

            ' Inv num
            N = Val(Parse(1))

            ' Make sure its a equipable item
            If Item(GetPlayerInvItemNum(Index, N)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_SHIELD Then
                Call PlayerMsg(Index, "Você pode apenas reparar armas, armaduras, elmos e escudos.", BrightRed)
                Exit Sub
            End If

            ' Check if they have a full inventory
            If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, N)) <= 0 Then
                Call PlayerMsg(Index, "Seu inventário está cheio!", BrightRed)
                Exit Sub
            End If

            ' Check if you can actually repair the item
            If Item(ItemNum).Data1 < 0 Then
                Call PlayerMsg(Index, "Esse item não é reparável!", BrightRed)
                Exit Sub
            End If

            ' Now check the rate of pay
            ItemNum = GetPlayerInvItemNum(Index, N)
            i = Int(Item(GetPlayerInvItemNum(Index, N)).Data2 / 5)

            If i <= 0 Then i = 1
            DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, N)
            GoldNeeded = Int(DurNeeded * i / 2)

            If GoldNeeded <= 0 Then GoldNeeded = 1

            ' Check if they even need it repaired
            If DurNeeded <= 0 Then
                Call PlayerMsg(Index, "Esse item está em perfeitas condições!", White)
                Exit Sub
            End If

            ' Check if they have enough for at least one point
            If HasItem(Index, 1) >= i Then

                ' Check if they have enough for a total restoration
                If HasItem(Index, 1) >= GoldNeeded Then
                    Call TakeItem(Index, 1, GoldNeeded)
                    Call SetPlayerInvItemDur(Index, N, Item(ItemNum).Data1 * -1)
                    Call PlayerMsg(Index, "Item foi completamente restaurado por " & GoldNeeded & " de ouro!", BrightBlue)
                Else

                    ' They dont so restore as much as we can
                    DurNeeded = (HasItem(Index, 1) / i)
                    GoldNeeded = Int(DurNeeded * i / 2)

                    If GoldNeeded <= 0 Then GoldNeeded = 1
                    Call TakeItem(Index, 1, GoldNeeded)
                    Call SetPlayerInvItemDur(Index, N, GetPlayerInvItemDur(Index, N) + DurNeeded)
                    Call PlayerMsg(Index, "Item foi reparado parcialmente por " & GoldNeeded & " de ouro!", BrightBlue)
                End If

            Else
                Call PlayerMsg(Index, "Ouro insuficiente para reparar o item!", BrightRed)
            End If

            Exit Sub

        Case "search"
            x = Val(Parse(1))
            y = Val(Parse(2))

            ' Prevent subscript out of range
            If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
                Exit Sub
            End If

            ' Check for a player
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y And i <> Index Then

                    ' Consider the player
                    If GetPlayerLevel(i) >= GetPlayerLevel(Index) + 5 Then
                        Call PlayerMsg(Index, "Você não teria chance alguma.", BrightRed)
                    Else

                        If GetPlayerLevel(i) > GetPlayerLevel(Index) Then
                            Call PlayerMsg(Index, "Suas chances seriam baixas.", Yellow)
                        Else

                            If GetPlayerLevel(i) = GetPlayerLevel(Index) Then
                                Call PlayerMsg(Index, "Essa seria uma luta justa.", White)
                            Else

                                If GetPlayerLevel(Index) >= GetPlayerLevel(i) + 5 Then
                                    Call PlayerMsg(Index, "Você aniquilaria ele.", BrightBlue)
                                Else

                                    If GetPlayerLevel(Index) > GetPlayerLevel(i) Then
                                        Call PlayerMsg(Index, "Você teria alguma vantagem sobre ele.", Yellow)
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_PLAYER
                    Call PlayerMsg(Index, "Seu alvo agora é " & GetPlayerName(i) & ".", Yellow)
                    Exit Sub
                End If

            Next

            ' Check for an npc
            For i = 1 To MAX_MAP_NPCS

                If MapNpc(GetPlayerMap(Index), i).num > 0 Then
                    If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then

                        ' Change target
                        Player(Index).Target = i
                        Player(Index).TargetType = TARGET_TYPE_NPC
                        Call PlayerMsg(Index, "Seu alvo agora é " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If

            Next

            ' Check for an item
            For i = 1 To MAX_MAP_ITEMS

                If MapItem(GetPlayerMap(Index), i).num > 0 Then
                    If MapItem(GetPlayerMap(Index), i).x = x And MapItem(GetPlayerMap(Index), i).y = y Then
                        Call PlayerMsg(Index, "Você vê um " & Trim$(Item(MapItem(GetPlayerMap(Index), i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If

            Next

            Exit Sub

        Case "playerchat"
            N = FindPlayer(Parse(1))

            If N < 1 Then
                Call PlayerMsg(Index, GetPlayerName(N) & " não está online.", White)
                Exit Sub
            End If

            If N = Index Then
                Exit Sub
            End If

            If Player(Index).InChat = 1 Then
                Call PlayerMsg(Index, "Você já está conversando com outro jogador!", Pink)
                Exit Sub
            End If

            If Player(N).InChat = 1 Then
                Call PlayerMsg(Index, GetPlayerName(N) & " já está conversando com outro jogador!", Pink)
                Exit Sub
            End If

            Call PlayerMsg(Index, "O pedido de conversa foi enviado para " & GetPlayerName(N) & ".", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " quer conversar com você.", Pink)
            Player(N).ChatPlayer = Index
            Player(Index).ChatPlayer = N
            
                Call ChatRequestWindow(N, Index)
            Exit Sub

        Case "achat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para conversar com você.", Pink)
                Exit Sub
            End If

            If Player(N).ChatPlayer <> Index Then
                Call PlayerMsg(Index, "A conversa falhou.", Pink)
                Exit Sub
            End If

            Call SendDataTo(Index, "PPCHATTING" & SEP_CHAR & N & END_CHAR)
            Call SendDataTo(N, "PPCHATTING" & SEP_CHAR & Index & END_CHAR)
            Exit Sub

        Case "dchat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para conversar com você.", Pink)
                Exit Sub
            End If

            Call PlayerMsg(Index, "O pedido de conversa foi rejeitado.", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " rejeitou seu pedido de conversa.", Pink)
            Player(Index).ChatPlayer = 0
            Player(Index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case "qchat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para conversar com você.", Pink)
                Exit Sub
            End If

            Call SendDataTo(Index, "qchat" & END_CHAR)
            Call SendDataTo(N, "qchat" & END_CHAR)
            Player(Index).ChatPlayer = 0
            Player(Index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case "sendchat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para conversar com você.", Pink)
                Exit Sub
            End If

            Call SendDataTo(N, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & Index & END_CHAR)
            Exit Sub

        Case "pptrade"
            N = FindPlayer(Parse(1))

            ' Check if player is online
            If N < 1 Then
                Call PlayerMsg(Index, Parse(1) & " não está online.", White)
                Exit Sub
            End If

            ' Prevent trading with self
            If N = Index Then
                Exit Sub
            End If

            ' Check if the player is in another trade
            If Player(Index).InTrade = 1 Then
                Call PlayerMsg(Index, "Você já está negociando com uma pessoa!", Pink)
                Exit Sub
            End If

            For i = 0 To 3

                If DirToX(GetPlayerX(Index), i) = GetPlayerX(N) And DirToY(GetPlayerY(Index), i) = GetPlayerY(N) Then

                    ' Check to see if player is already in a trade
                    If Player(N).InTrade = 1 Then
                        Call PlayerMsg(Index, GetPlayerName(N) & " já está negociando com alguém!", Pink)
                        Exit Sub
                    End If

                    Call PlayerMsg(Index, "Pedido de negociação enviado para " & GetPlayerName(N) & ".", Pink)
                    Call PlayerMsg(N, GetPlayerName(Index) & " enviou um pedido de negociação.", Pink)
                    Player(N).TradePlayer = Index
                    Player(Index).TradePlayer = N
                    
                            Call TradeRequestWindow(N, Index)
                    Exit Sub
                End If

            Next

            Call PlayerMsg(Index, "Você precisa estar perto do jogador para negociar!", Pink)
            Call PlayerMsg(N, "O jogador precisa estar perto de você para negociar!", Pink)
            Exit Sub

        Case "atrade"
            N = Player(Index).TradePlayer

            ' Check if anyone requested a trade
            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para negociar com você.", Pink)
                Exit Sub
            End If

            ' Check if its the right player
            If Player(N).TradePlayer <> Index Then
                Call PlayerMsg(Index, "Negociação falhou.", Pink)
                Exit Sub
            End If

            ' Check where both players are
            For i = 0 To 3

                If DirToX(GetPlayerX(Index), i) = GetPlayerX(N) And DirToY(GetPlayerY(Index), i) = GetPlayerY(N) Then
                    Call PlayerMsg(Index, "Você está negociando com " & GetPlayerName(N) & "!", Pink)
                    Call PlayerMsg(N, GetPlayerName(Index) & " aceitou seu pedido de negociação!", Pink)
                    Call SendDataTo(Index, "PPTRADING" & END_CHAR)
                    Call SendDataTo(N, "PPTRADING" & END_CHAR)

                    For o = 1 To MAX_PLAYER_TRADES
                        Player(Index).Trading(o).InvNum = 0
                        Player(Index).Trading(o).InvName = vbNullString
                        Player(N).Trading(o).InvNum = 0
                        Player(N).Trading(o).InvName = vbNullString
                    Next

                    Player(Index).InTrade = 1
                    Player(Index).TradeItemMax = 0
                    Player(Index).TradeItemMax2 = 0
                    Player(N).InTrade = 1
                    Player(N).TradeItemMax = 0
                    Player(N).TradeItemMax2 = 0
                    Exit Sub
                End If

            Next

            Call PlayerMsg(Index, "Você precisa estar perto do jogador para negociar!", Pink)
            Call PlayerMsg(N, "O jogador precisa estar perto de você para negociar!", Pink)
            Exit Sub

        Case "qtrade"
            N = Player(Index).TradePlayer

            ' Check if anyone trade with player
            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para negociar com você.", Pink)
                Exit Sub
            End If

            Call PlayerMsg(Index, "A negociação terminou. Nada feito!", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " parou de negociar com você!", Pink)
            Player(Index).TradeOk = 0
            Player(N).TradeOk = 0
            Player(Index).TradePlayer = 0
            Player(Index).InTrade = 0
            Player(N).TradePlayer = 0
            Player(N).InTrade = 0
            Call SendDataTo(Index, "qtrade" & END_CHAR)
            Call SendDataTo(N, "qtrade" & END_CHAR)
            Exit Sub

        Case "dtrade"
            N = Player(Index).TradePlayer

            ' Check if anyone trade with player
            If N < 1 Then
                Call PlayerMsg(Index, "Ninguém pediu para negociar com você.", Pink)
                Exit Sub
            End If

            Call PlayerMsg(Index, "O pedido de negociação foi rejeitado.", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " rejeitou seu pedido de negociação.", Pink)
            Player(Index).TradePlayer = 0
            Player(Index).InTrade = 0
            Player(N).TradePlayer = 0
            Player(N).InTrade = 0
            Exit Sub

        Case "updatetradeinv"
            N = Val(Parse(1))
            Player(Index).Trading(N).InvNum = Val(Parse(2))
            Player(Index).Trading(N).InvName = Trim$(Parse(3))

            If Player(Index).Trading(N).InvNum = 0 Then
                Player(Index).TradeItemMax = Player(Index).TradeItemMax - 1
                Player(Index).TradeOk = 0
                Player(N).TradeOk = 0
                Call SendDataTo(Index, "trading" & SEP_CHAR & 0 & END_CHAR)
                Call SendDataTo(N, "trading" & SEP_CHAR & 0 & END_CHAR)
            Else
                Player(Index).TradeItemMax = Player(Index).TradeItemMax + 1
            End If

            Call SendDataTo(Player(Index).TradePlayer, "updatetradeitem" & SEP_CHAR & N & SEP_CHAR & Player(Index).Trading(N).InvNum & SEP_CHAR & Player(Index).Trading(N).InvName & END_CHAR)
            Exit Sub

        Case "swapitems"
            N = Player(Index).TradePlayer

            If Player(Index).TradeOk = 0 Then
                Player(Index).TradeOk = 1
                Call SendDataTo(N, "trading" & SEP_CHAR & 1 & END_CHAR)
            ElseIf Player(Index).TradeOk = 1 Then
                Player(Index).TradeOk = 0
                Call SendDataTo(N, "trading" & SEP_CHAR & 0 & END_CHAR)
            End If

            If Player(Index).TradeOk = 1 And Player(N).TradeOk = 1 Then
                Player(Index).TradeItemMax2 = 0
                Player(N).TradeItemMax2 = 0

                For i = 1 To MAX_INV

                    If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(N, i) < 1 Then
                        Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                    End If

                Next

                For i = 1 To MAX_INV

                    If Player(N).TradeItemMax = Player(N).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(Index, i) < 1 Then
                        Player(N).TradeItemMax2 = Player(N).TradeItemMax2 + 1
                    End If

                Next

                If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(N).TradeItemMax2 = Player(N).TradeItemMax Then

                    For i = 1 To MAX_PLAYER_TRADES
                        For x = 1 To MAX_INV

                            If GetPlayerInvItemNum(N, x) < 1 Then
                                If Player(Index).Trading(i).InvNum > 0 Then
                                    Call GiveItem(N, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), 1)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), 1)
                                    Exit For
                                End If
                            End If

                        Next
                    Next

                    For i = 1 To MAX_PLAYER_TRADES
                        For x = 1 To MAX_INV

                            If GetPlayerInvItemNum(Index, x) < 1 Then
                                If Player(N).Trading(i).InvNum > 0 Then
                                    Call GiveItem(Index, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), 1)
                                    Call TakeItem(N, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), 1)
                                    Exit For
                                End If
                            End If

                        Next
                    Next

                    Call PlayerMsg(N, "Trade Successfull!", BrightGreen)
                    Call PlayerMsg(Index, "Trade Successfull!", BrightGreen)
                    Call SendInventory(N)
                    Call SendInventory(Index)
                Else

                    If Player(Index).TradeItemMax2 < Player(Index).TradeItemMax Then
                        Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
                        Call PlayerMsg(N, GetPlayerName(Index) & "'s inventory is full!", BrightRed)
                    End If

                    If Player(N).TradeItemMax2 < Player(N).TradeItemMax Then
                        Call PlayerMsg(N, "Your inventory is full!", BrightRed)
                        Call PlayerMsg(Index, GetPlayerName(N) & "'s inventory is full!", BrightRed)
                    End If
                End If

                Player(Index).TradePlayer = 0
                Player(Index).InTrade = 0
                Player(Index).TradeOk = 0
                Player(N).TradePlayer = 0
                Player(N).InTrade = 0
                Player(N).TradeOk = 0
                Call SendDataTo(Index, "qtrade" & END_CHAR)
                Call SendDataTo(N, "qtrade" & END_CHAR)
            End If

            Exit Sub

        Case "party"
            N = FindPlayer(Parse(1))

            If N = Index Then Exit Sub
            If N > 0 Then
                If GetPlayerAccess(Index) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "Você não pode entrar em um grupo, você é um Admin!", BrightBlue)
                    Exit Sub
                End If

                If GetPlayerAccess(N) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "Admins não podem entrar em grupos!", BrightBlue)
                    Exit Sub
                End If

                If Player(N).InParty = 0 Then
                    If Player(Index).PartyID > 0 Then
                        If Party(Player(Index).PartyID).Member(MAX_PARTY_MEMBERS) <> 0 Then
                            Call PlayerMsg(Index, GetPlayerName(N) & " foi convidado para o seu grupo.", Pink)
                            Call PlayerMsg(N, GetPlayerName(Index) & " convidou você para participar de seu grupo.", Pink)
                            Player(N).Invited = Player(Index).PartyID
                
                            Call PartyRequestWindow(N, Index)
                        Else
                            Call PlayerMsg(Index, "Seu grupo está cheio.", Pink)
                        End If

                    Else
                        o = 0
                        i = MAX_PARTIES

                        Do While i > 0

                            If Party(i).Member(1) = 0 Then o = i
                            i = i - 1
                        Loop

                        If o = 0 Then
                            Call PlayerMsg(Index, "Grupo sobrecarregado.", Pink)
                            Exit Sub
                        End If

                        Party(o).Member(1) = Index
                        Player(Index).InParty = YES
                        Player(Index).PartyID = o
                        Player(Index).Invited = 0
                        Call PlayerMsg(Index, "Grupo criado.", Pink)
                        Call PlayerMsg(Index, GetPlayerName(N) & " foi convidado para o seu grupo.", Pink)
                        Call PlayerMsg(N, GetPlayerName(Index) & " convidou você para entrar em seu grupo.", Pink)
                        Player(N).Invited = Player(Index).PartyID
                        
                            Call PartyRequestWindow(N, Index)
                        Call SendDataToMap(MapNum, Packet)
                    End If

                Else
                    Call PlayerMsg(Index, GetPlayerName(N) & " já está em um grupo.", Pink)
                End If

            Else
                Call PlayerMsg(Index, Parse(1) & " não está online.", White)
            End If

            Exit Sub

        Case "joinparty"

            If Player(Index).Invited > 0 Then
                o = 0

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).Invited).Member(i) = 0 Then
                        If o = 0 Then o = i
                    End If

                Next

                If o <> 0 Then
                    Player(Index).PartyID = Player(Index).Invited
                    Player(Index).InParty = YES
                    Player(Index).Invited = 0
                    Party(Player(Index).PartyID).Member(o) = Index

                    For i = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(Index).PartyID).Member(i) <> 0 And Party(Player(Index).PartyID).Member(i) <> Index Then
                            Call PlayerMsg(Party(Player(Index).PartyID).Member(i), GetPlayerName(Index) & " entrou no seu grupo!", Pink)
                        End If

                    Next

                    Call PlayerMsg(Index, "Você entrou no grupo!", Pink)
                Else
                    Call PlayerMsg(Index, "O grupo está cheio!", Pink)
                End If

            Else
                Call PlayerMsg(Index, "Você não foi convidado para nenhum grupo!", Pink)
            End If

            Exit Sub

        Case "leaveparty"

            If Player(Index).PartyID > 0 Then
                Call PlayerMsg(Index, "Você saiu do grupo.", Pink)
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
                    Player(Party(Player(Index).PartyID).Member(1)).InParty = NO
                    Player(Party(Player(Index).PartyID).Member(1)).PartyID = 0
                    Party(Player(Index).PartyID).Member(1) = 0
                End If
                End If

                Player(Index).InParty = 0
                Player(Index).PartyID = 0
            Else

                If Player(Index).Invited <> 0 Then

                    For i = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(Index).Invited).Member(i) <> 0 And Party(Player(Index).Invited).Member(i) <> Index Then Call PlayerMsg(Party(Player(Index).Invited).Member(i), GetPlayerName(Index) & " rejeitou o convite.", Pink)
                    Next

                    Player(Index).Invited = 0
                    Call PlayerMsg(Index, "Você rejeitou o convite.", Pink)
                Else
                    Call PlayerMsg(Index, "Você não foi convidado para nenhum grupo!", Pink)
                End If
            End If

            Exit Sub

        Case "partychat"

            If Player(Index).PartyID > 0 Then

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(i) <> 0 Then Call PlayerMsg(Party(Player(Index).PartyID).Member(i), GetPlayerName(Index) & ": " & Parse(1), PartyColor)
                Next

            Else
                Call PlayerMsg(Index, "Você não está em um grupo!", Pink)
            End If

            Exit Sub

        Case "guildchat"

            If GetPlayerGuild(Index) <> vbNullString Then

                For i = 1 To MAX_PLAYERS

                    If GetPlayerGuild(Index) = GetPlayerGuild(i) Then Call PlayerMsg(i, GetPlayerName(Index) & ": " & Parse(1), GuildColor)
                Next

            Else
                Call PlayerMsg(Index, "Você não está em uma guild!", Pink)
            End If

            Exit Sub

        Case "newmain"

            If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
                Dim temp As String

                f = FreeFile
                Open App.Path & "\Scripts\Principal.txt" For Input As #f
                temp = Input$(LOF(f), f)
                Close #f
                f = FreeFile
                Open App.Path & "\Scripts\Backup.txt" For Output As #f
                Print #f, temp
                Close #f
                f = FreeFile
                Open App.Path & "\Scripts\Principal.txt" For Output As #f
                Print #f, Parse(1)
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\Scripts\Principal.txt", "Scripts\Principal.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call TextAdd(frmServer.txtText(0), "Scripts atualizados.", True)
                    Call PlayerMsg(Index, "Scripts atualizados.", White)
                End If

                Call AddLog(GetPlayerName(Index) & " atualizou os scripts.", ADMIN_LOG)
            End If

            Exit Sub

        Case "requestbackupmain"

            If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
                Dim nothertemp As String

                f = FreeFile
                Open App.Path & "\Scripts\Backup.txt" For Input As #f
                nothertemp = Input$(LOF(f), f)
                Close #f
                f = FreeFile
                Open App.Path & "\Scripts\Principal.txt" For Output As #f
                Print #f, nothertemp
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\Scripts\Principal.txt", "Scripts\Principal.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call TextAdd(frmServer.txtText(0), "Scripts atualizados.", True)
                    Call PlayerMsg(Index, "Scripts atualizados.", White)
                End If

                Call AddLog(GetPlayerName(Index) & " usou o script de backup.", ADMIN_LOG)
            End If

            Exit Sub

        Case "spells"
            Call SendPlayerSpells(Index)
            Exit Sub

        Case "cast"
            N = Val(Parse(1))
            Call CastSpell(Index, N)
            Exit Sub

        Case "requestlocation"

            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call PlayerMsg(Index, "Mapa: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
            Exit Sub

        Case "refresh"
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), False)
            Call PlayerMsg(Index, "Mapa atualizado.", White)
            Exit Sub

        Case "killpet"
        If Player(Index).Pet.Alive = YES Then
            Player(Index).Pet.Alive = NO
            Player(Index).Pet.Sprite = 0
            Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
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
            Call SendDataToMap(GetPlayerMap(Index), Packet)
        Else
            Call PlayerMsg(Index, "Você não tem um mascote vivo.", Red)
        End If
            Exit Sub

        Case "petmoveselect"
            x = Val(Parse(1))
            y = Val(Parse(2))
            Player(Index).Pet.MapToGo = GetPlayerMap(Index)
            Player(Index).Pet.Target = 0
            Player(Index).Pet.XToGo = x
            Player(Index).Pet.YToGo = y
            Player(Index).Pet.AttackTimer = GetTickCount

            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) Then
                    If GetPlayerMap(i) = Player(Index).Pet.Map Then
                        If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                            Player(Index).Pet.TargetType = TARGET_TYPE_PLAYER
                            Player(Index).Pet.Target = i
                            Call PlayerMsg(Index, "O alvo do seu mascote agora é " & Trim$(GetPlayerName(i)) & ".", Yellow)
                            Exit Sub
                        End If
                    End If
                End If

            Next

            For i = 1 To MAX_MAP_NPCS

                If MapNpc(Player(Index).Pet.Map, i).num > 0 Then
                    If MapNpc(Player(Index).Pet.Map, i).x = x And MapNpc(Player(Index).Pet.Map, i).y = y Then
                        Player(Index).Pet.TargetType = TARGET_TYPE_NPC
                        Player(Index).Pet.Target = i
                        Call PlayerMsg(Index, "O alvo do seu mascote agora é " & Trim$(Npc(MapNpc(Player(Index).Pet.Map, i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If

            Next

            Exit Sub

        Case "buysprite"

            ' Check if player stepped on sprite changing tile
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
                Call PlayerMsg(Index, "You need to be on a sprite tile to buy it!", BrightRed)
                Exit Sub
            End If

            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
                Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                Exit Sub
            End If

            For i = 1 To MAX_INV

                If GetPlayerInvItemNum(Index, i) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                    If Item(GetPlayerInvItemNum(Index, i)).Type = ITEM_TYPE_CURRENCY Then
                        If GetPlayerInvItemValue(Index, i) >= Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 Then
                            Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3)

                            If GetPlayerInvItemValue(Index, i) <= 0 Then
                                Call SetPlayerInvItemNum(Index, i, 0)
                            End If

                            Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                            Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                            Call SendInventory(Index)
                        End If

                    Else

                        If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                            Call SetPlayerInvItemNum(Index, i, 0)
                            Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                            Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                            Call SendInventory(Index)
                        End If
                    End If

                    If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                        Exit Sub
                    End If
                End If

            Next

            Call PlayerMsg(Index, "You dont have enough to buy this sprite!", BrightRed)
            Exit Sub

        Case "checkcommands"
            s = Parse(1)

            If SCRIPTING = 1 Then
                PutVar App.Path & "\Scripts\Comandos.ini", "TEMP", "Text" & Index, Trim$(s)
                MyScript.ExecuteStatement "Scripts\Principal.txt", "Commands " & Index
            Else
                Call PlayerMsg(Index, "Thats not a valid command!", 12)
            End If

            Exit Sub

        Case "prompt"

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Principal.txt", "PlayerPrompt " & Index & "," & Val(Parse(1)) & "," & Val(Parse(2))
            End If

            Exit Sub

        Case "requesteditarrow"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "ARROWEDITOR" & END_CHAR)
            Exit Sub

        Case "editarrow"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_ARROWS Then
                Call HackingAttempt(Index, "Índice de flecha inválido")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editando flecha #" & N & ".", ADMIN_LOG)
            Call SendEditArrowTo(Index, N)
            Exit Sub

        Case "savearrow"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Índice de flecha inválido")
                Exit Sub
            End If

            Arrows(N).Name = Parse(2)
            Arrows(N).Pic = Val(Parse(3))
            Arrows(N).Range = Val(Parse(4))
            Call SendUpdateArrowToAll(N)
            Call SaveArrow(N)
            Call AddLog(GetPlayerName(Index) & " salvou Flecha #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case "checkarrows"
            N = Arrows(Val(Parse(1))).Pic
            Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & N & END_CHAR)
            Exit Sub

        Case "speechscript"

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Principal.txt", "ScriptedTile " & Index & "," & Parse(1)
            End If

            Exit Sub

        Case "requesteditspeech"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "SPEECHEDITOR" & END_CHAR)
            Exit Sub

        Case "editspeech"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(Index, "Invalid Speech Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editing speech #" & N & ".", ADMIN_LOG)
            Call SendEditSpeechTo(Index, N)
            Exit Sub

        Case "savespeech"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(Index, "Invalid Speech Index")
                Exit Sub
            End If

            Speech(N).Name = Parse(2)
            Dim p As Long

            p = 3

            For i = 0 To MAX_SPEECH_OPTIONS
                Speech(N).num(i).Exit = Val(Parse(p))
                Speech(N).num(i).text = Parse(p + 1)
                Speech(N).num(i).SaidBy = Val(Parse(p + 2))
                Speech(N).num(i).Respond = Val(Parse(p + 3))
                Speech(N).num(i).Script = Val(Parse(p + 4))
                p = p + 5

                For o = 1 To 3
                    Speech(N).num(i).Responces(o).Exit = Val(Parse(p))
                    Speech(N).num(i).Responces(o).GoTo = Val(Parse(p + 1))
                    Speech(N).num(i).Responces(o).text = Parse(p + 2)
                    p = p + 3
                Next
            Next

            Call SaveSpeech(N)
            Call SendSpeechToAll(N)
            Call AddLog(GetPlayerName(Index) & " salvou Fala #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case "needspeech"
            Call SendSpeechTo(Index, Val(Parse(1)))
            Exit Sub

        Case "requesteditemoticon"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, "EMOTICONEDITOR" & END_CHAR)
            Exit Sub

        Case "editemoticon"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(Index, "Invalid Emoticon Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editing emoticon #" & N & ".", ADMIN_LOG)
            Call SendEditEmoticonTo(Index, N)
            Exit Sub

        Case "saveemoticon"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(Index, "Invalid Emoticon Index")
                Exit Sub
            End If

            Emoticons(N).Type = Val(Parse(2))
            Emoticons(N).Command = Parse(3)
            Emoticons(N).Pic = Val(Parse(4))
            Emoticons(N).sound = Parse(5)
            Call SendUpdateEmoticonToAll(N)
            Call SaveEmoticon(N)
            Call AddLog(GetPlayerName(Index) & " salvou Emoticon #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case "checkemoticons"
            Call SendDataToMap(GetPlayerMap(Index), "checkemoticons" & SEP_CHAR & Index & SEP_CHAR & Emoticons(Val(Parse(1))).Type & SEP_CHAR & Emoticons(Val(Parse(1))).Pic & SEP_CHAR & Emoticons(Val(Parse(1))).sound & END_CHAR)
            Exit Sub

        Case "mapreport"
            Packs = "mapreport" & SEP_CHAR

            For i = 1 To MAX_MAPS
                Packs = Packs & Map(i).Name & SEP_CHAR
            Next

            Packs = Packs & END_CHAR
            Call SendDataTo(Index, Packs)
            Exit Sub

        Case "gmtime"
            GameTime = Val(Parse(1))
            Call SendTimeToAll
            Exit Sub

        Case "weather"
            GameWeather = Val(Parse(1))
            Call SendWeatherToAll
            Exit Sub

        Case "warpto"
            Call PlayerWarp(Index, Val(Parse(1)), GetPlayerX(Index), GetPlayerY(Index))
            Exit Sub

        Case "warptome"
            N = FindPlayer(Parse(1))

            If N > 0 Then
                Call PlayerWarp(N, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Else
                Call PlayerMsg(Index, "Player not online!", BrightRed)
            End If

            Exit Sub

        Case "warpplayer"

            If Val(Parse(1)) > MAX_MAPS Or Val(Parse(1)) < 1 Then
                If FindPlayer(Trim$(Parse(1))) <> 0 Then
                    Call PlayerWarp(Index, GetPlayerMap(FindPlayer(Trim$(Parse(1)))), GetPlayerX(FindPlayer(Trim$(Parse(1)))), GetPlayerY(FindPlayer(Trim$(Parse(1)))))

                    If Player(Index).Pet.Alive = YES Then
                        Player(Index).Pet.Map = GetPlayerMap(Index)
                        Player(Index).Pet.x = GetPlayerX(Index)
                        Player(Index).Pet.y = GetPlayerY(Index)
                        Player(Index).Pet.MapToGo = -1
                        Player(Index).Pet.XToGo = -1
                        Player(Index).Pet.YToGo = -1
                    End If

                Else
                    Call PlayerMsg(Index, "'" & Parse(1) & "' não é um mapa ou um jogador online válido!", BrightRed)
                    Exit Sub
                End If

            Else
                Call PlayerWarp(Index, Val(Parse(1)), GetPlayerX(Index), GetPlayerY(Index))

                If Player(Index).Pet.Alive = YES Then
                    Player(Index).Pet.Map = GetPlayerMap(Index)
                    Player(Index).Pet.x = GetPlayerX(Index)
                    Player(Index).Pet.y = GetPlayerY(Index)
                    Player(Index).Pet.MapToGo = -1
                    Player(Index).Pet.XToGo = -1
                    Player(Index).Pet.YToGo = -1
                End If
            End If

            Exit Sub

        Case "arrowhit"
            N = Val(Parse(1))
            z = Val(Parse(2))
            x = Val(Parse(3))
            y = Val(Parse(4))

            If N = TARGET_TYPE_PLAYER Then

                ' Make sure we dont try to attack ourselves
                If z <> Index Then

                    ' Can we attack the player?
                    If CanAttackPlayerWithArrow(Index, z) Then
                        If Not CanPlayerBlockHit(z) Then

                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque" & Int(Rnd * 2) + 1 & END_CHAR)
                            Else
                                N = GetPlayerDamage(Index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call BattleMsg(Index, "Você sente uma enorme quantidade de energia em seu arco!", BrightCyan, 0)
                                Call BattleMsg(z, GetPlayerName(Index) & " atira com incrível precisão!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
                                'Call PlayerMsg(z, GetPlayerName(index) & " shoots with amazing accuracy!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque3" & END_CHAR)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(Index, z, Damage)
                            Else
                                Call BattleMsg(Index, "Sua ataque não fez nada.", BrightRed, 0)
                                Call BattleMsg(z, "Ataque de " & GetPlayerName(z) & " não fez nada.", BrightRed, 1)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Errou" & END_CHAR)
                            End If

                        Else
                            Call BattleMsg(Index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                            Call BattleMsg(z, "You blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(z, "Your " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Errou" & END_CHAR)
                        End If

                        Exit Sub
                    End If
                End If

            ElseIf N = TARGET_TYPE_NPC Then

                ' Can we attack the npc?
                If CanAttackNpcWithArrow(Index, z) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque" & Int(Rnd * 2) + 1 & END_CHAR)
                    Else
                        N = GetPlayerDamage(Index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(Index, "Você sente uma enorme quantidade de energia em seu corpo!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Ataque3" & END_CHAR)
                    End If

                    If Damage > 0 Then
                        Call AttackNpc(Index, z, Damage)
                        Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & END_CHAR)
                    Else
                        Call BattleMsg(Index, "Seu ataque não fez nada.", BrightRed, 0)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Errou" & END_CHAR)
                    End If

                    Exit Sub
                End If
            End If

            Exit Sub
    End Select

    Call HackingAttempt(Index, "Invalid packet. (" & Parse(0) & ")")
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim Buffer As String
    Dim Packet As String
    Dim top As String * 3
    Dim Start As Long

    If Index > 0 Then
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength

        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(Index, top)
            Call CloseSocket(Index)
        End If

        Player(Index).Buffer = Player(Index).Buffer & Buffer
        Start = InStr(Player(Index).Buffer, END_CHAR)

        Do While Start > 0
            Packet = Mid$(Player(Index).Buffer, 1, Start - 1)
            Player(Index).Buffer = Mid$(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
            Player(Index).DataPackets = Player(Index).DataPackets + 1
            Start = InStr(Player(Index).Buffer, END_CHAR)

            If Len(Packet) > 0 Then
                Call HandleData(Index, Packet)
            End If

        Loop

        ' Not useful
        ' Check if elapsed time has passed
        Player(Index).DataBytes = Player(Index).DataBytes + DataLength

        If GetTickCount >= Player(Index).DataTimer + 1000 Then
            Player(Index).DataTimer = GetTickCount
            Player(Index).DataBytes = 0
            Player(Index).DataPackets = 0
            Exit Sub
        End If

        ' Check for data flooding
        If Player(Index).DataBytes > 1000 And GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Data Flooding")
            Exit Sub
        End If

        ' Check for packet flooding
        If Player(Index).DataPackets > 25 And GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Packet Flooding")
            Exit Sub
        End If
    End If

End Sub

Function IsBanned(ByVal IP As String) As Boolean
    Dim FileName As String, fIP As String, fName As String
    Dim f As Long

    IsBanned = False
    FileName = App.Path & "\banlist.txt"

    ' Check if file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    f = FreeFile
    Open FileName For Input As #f

    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

    If IsConnected(Index) And Trim$(Player(Index).Login) <> vbNullString Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    IsMultiAccounts = False

    For i = 1 To MAX_PLAYERS

        If IsConnected(i) And LCase$(Trim$(Player(i).Login)) = LCase$(Trim$(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If

    Next

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If Index <= 0 Or Index > MAX_PLAYERS Then
        IsPlaying = False
        Exit Function
    End If

    If IsConnected(Index) And Player(Index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If

End Function

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
    Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal num As Long)
    Dim Packet As String

    Packet = "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub PartyRequestWindow(ByVal Index As Long, ByVal Inviter As Long)
    Dim Packet As String

    Packet = "PARTYWINDOW" & SEP_CHAR & Inviter & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub TradeRequestWindow(ByVal Index As Long, ByVal Inviter As Long)
    Dim Packet As String

    Packet = "TRADEWINDOW" & SEP_CHAR & Inviter & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub ChatRequestWindow(ByVal Index As Long, ByVal Inviter As Long)
    Dim Packet As String

    Packet = "CHATWINDOW" & SEP_CHAR & Inviter & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendArrows(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(Index, i)
    Next

End Sub

Sub SendChars(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "ALLCHARS" & SEP_CHAR

    For i = 1 To MAX_CHARS
        Packet = Packet & Trim$(Player(Index).Char(i).Name) & SEP_CHAR & Trim$(Class(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR

    For i = 1 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)

    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data

        DoEvents
    End If

End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
    Dim Packet As String

    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
    Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Emoticons(EmoNum).Type & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & Emoticons(EmoNum).sound & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Packet As String
    Dim i As Long

    'Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).str & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & END_CHAR
    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).Magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR

    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    Dim i As Long, z As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR

    For i = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetValue & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpeechTo(ByVal Index As Long, ByVal SpcNum As Long)
    Dim Packet As String
    Dim i, o As Long

    Packet = "EDITSPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR

    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR

        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEmoticons(ByVal Index As Long)
    Dim i As Long

    For i = 0 To MAX_EMOTICONS

        If Trim$(Emoticons(i).Command) <> vbNullString Then
            Call SendUpdateEmoticonTo(Index, i)
        End If

    Next

End Sub

Sub SendFriendListTo(ByVal Index As Long)
    Dim Packet As String
    Dim N As Long

    Packet = "FRIENDLIST" & SEP_CHAR

    For N = 1 To MAX_FRIENDS

        If FindPlayer(Player(Index).Char(Player(Index).CharNum).Friends(N)) And Player(Index).Char(Player(Index).CharNum).Friends(N) <> vbNullString Then
            Packet = Packet & Player(Index).Char(Player(Index).CharNum).Friends(N) & SEP_CHAR
        End If

    Next

    Packet = Packet & NEXT_CHAR & SEP_CHAR

    For N = 1 To MAX_FRIENDS
        Packet = Packet & Player(Index).Char(Player(Index).CharNum).Friends(N) & SEP_CHAR
    Next

    Packet = Packet & NEXT_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendFriendListToNeeded(ByVal Name As String)
    Dim i, o As Long

    For i = i To MAX_PLAYERS

        If IsPlaying(i) Then

            For o = 1 To MAX_FRIENDS

                If Trim$(Player(i).Char(Player(i).CharNum).Friends(i)) = Name Then
                    Call SendFriendListTo(i)
                End If

            Next

        End If

    Next

End Sub

Sub SendHP(ByVal Index As Long)
    Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInfo(ByVal Index As Long)
    Dim Packet As String

    Packet = "INFO" & SEP_CHAR & TotalOnlinePlayers & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR

    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Dim Packet As String

    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Trim$(Item(i).Name) <> vbNullString Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

End Sub

Sub SendJoinMap(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = vbNullString

    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index And GetPlayerMap(i) = GetPlayerMap(Index) Then
            Packet = "PLAYERDATA" & SEP_CHAR
            Packet = Packet & i & SEP_CHAR
            Packet = Packet & GetPlayerName(i) & SEP_CHAR
            Packet = Packet & GetPlayerSprite(i) & SEP_CHAR
            Packet = Packet & GetPlayerMap(i) & SEP_CHAR
            Packet = Packet & GetPlayerX(i) & SEP_CHAR
            Packet = Packet & GetPlayerY(i) & SEP_CHAR
            Packet = Packet & GetPlayerDir(i) & SEP_CHAR
            Packet = Packet & GetPlayerAccess(i) & SEP_CHAR
            Packet = Packet & GetPlayerPK(i) & SEP_CHAR
            Packet = Packet & GetPlayerGuild(i) & SEP_CHAR
            Packet = Packet & GetPlayerGuildAccess(i) & SEP_CHAR
            Packet = Packet & GetPlayerClass(i) & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(Index, Packet)

            If Player(i).Pet.Alive = YES Then
                Packet = "PETDATA" & SEP_CHAR
                Packet = Packet & i & SEP_CHAR
                Packet = Packet & Player(i).Pet.Alive & SEP_CHAR
                Packet = Packet & Player(i).Pet.Map & SEP_CHAR
                Packet = Packet & Player(i).Pet.x & SEP_CHAR
                Packet = Packet & Player(i).Pet.y & SEP_CHAR
                Packet = Packet & Player(i).Pet.Dir & SEP_CHAR
                Packet = Packet & Player(i).Pet.Sprite & SEP_CHAR
                Packet = Packet & Player(i).Pet.HP & SEP_CHAR
                Packet = Packet & Player(i).Pet.Level * 5 & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataTo(Index, Packet)
            End If
        End If

    Next

    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

    If Player(Index).Pet.Alive = YES Then
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
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If

End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)

    If Player(Index).Pet.Alive = YES Then
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
        Call SendDataToMapBut(Index, MapNum, Packet)
    End If

End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & vbNullString & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & vbNullString & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(Index, Packet)
    Packet = "PETDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(Index, Packet)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim o As Long
    Dim p1 As String, p2 As String

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX

            With Map(MapNum).Tile(x, y)
                i = 0
                o = 0

                If .Ground <> 0 Then i = 0
                If .GroundSet <> -1 Then i = 1
                If .Mask <> 0 Then i = 2
                If .MaskSet <> -1 Then i = 3
                If .Anim <> 0 Then i = 4
                If .AnimSet <> -1 Then i = 5
                If .Fringe <> 0 Then i = 6
                If .FringeSet <> -1 Then i = 7
                If .Type <> 0 Then i = 8
                If .Data1 <> 0 Then i = 9
                If .Data2 <> 0 Then i = 10
                If .Data3 <> 0 Then i = 11
                If .String1 <> vbNullString Then i = 12
                If .String2 <> vbNullString Then i = 13
                If .String3 <> vbNullString Then i = 14
                If .Mask2 <> 0 Then i = 15
                If .Mask2Set <> -1 Then i = 16
                If .M2Anim <> 0 Then i = 17
                If .M2AnimSet <> -1 Then i = 18
                If .FAnim <> 0 Then i = 19
                If .FAnimSet <> -1 Then i = 20
                If .Fringe2 <> 0 Then i = 21
                If .Fringe2Set <> -1 Then i = 22
                If .Light <> 0 Then i = 23
                If .F2Anim <> 0 Then i = 24
                If .F2AnimSet <> -1 Then i = 25
                Packet = Packet & .Ground & SEP_CHAR

                If o < i Then
                    o = o + 1
                    Packet = Packet & .GroundSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Mask & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .MaskSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Anim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .AnimSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Fringe & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .FringeSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Type & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Data1 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Data2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Data3 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .String1 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .String2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .String3 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Mask2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Mask2Set & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .M2Anim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .M2AnimSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .FAnim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .FAnimSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Fringe2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Fringe2Set & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Light & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .F2Anim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .F2AnimSet & SEP_CHAR
                End If

                Packet = Packet & NEXT_CHAR & SEP_CHAR
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(x) & SEP_CHAR
        Packet = Packet & Map(MapNum).NpcSpawn(x).Used & SEP_CHAR & Map(MapNum).NpcSpawn(x).x & SEP_CHAR & Map(MapNum).NpcSpawn(x).y & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    x = Int(Len(Packet) / 2)
    p1 = Mid$(Packet, 1, x)
    p2 = Mid$(Packet, x + 1, Len(Packet) - x)
    Call SendDataTo(Index, Packet)
    
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR

    For i = 1 To MAX_MAP_ITEMS

        If MapNum > 0 Then
            Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
        End If

    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR

    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR

    For i = 1 To MAX_MAP_NPCS

        If MapNum > 0 Then
            Packet = Packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
        End If

    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
    Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim i As Long
    Dim Packet As String

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR

    For i = 1 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Trim$(Npc(i).Name) <> vbNullString Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

End Sub

Sub SendOnlineList()
    Dim Packet As String
    Dim i As Long
    Dim N As Long

    Packet = vbNullString
    N = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
            N = N + 1
        End If

    Next

    Packet = "ONLINELIST" & SEP_CHAR & N & Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim Packet As String

    ' Send index's player data to everyone including himself on the map
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

    If Player(Index).Pet.Alive = YES Then
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
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If

End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "SPELLS" & SEP_CHAR

    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, i) & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Trim$(Shop(i).Name) <> vbNullString Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

End Sub

Sub SendSP(ByVal Index As Long)
    Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpeech(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPEECH

        If Trim$(Speech(i).Name) <> vbNullString Then
            Call SendSpeechTo(Index, i)
        End If

    Next

End Sub

Sub SendSpeechTo(ByVal Index As Long, ByVal SpcNum As Long)
    Dim Packet As String
    Dim i, o As Long

    Packet = "SPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR

    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR

        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpeechToAll(ByVal SpcNum As Long)
    Dim Packet As String
    Dim i, o As Long

    Packet = "SPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR

    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR

        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Trim$(Spell(i).Name) <> vbNullString Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next

End Sub

Sub SendStats(ByVal Index As Long)
    Dim Packet As String

    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerstr(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeTo(ByVal Index As Long)
    Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeToAll()
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If

    Next

    Call SpawnAllMapNpcs
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    Dim i As Long, x As Long, y As Long, z As Long, XX As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR

    For i = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetValue & SEP_CHAR

            ' Item #
            x = Shop(ShopNum).TradeItem(i).Value(XX).GetItem

            If Item(x).Type = ITEM_TYPE_SPELL Then

                ' Spell class requirement
                y = Spell(Item(x).Data1).ClassReq

                If y = 0 Then
                    Call PlayerMsg(Index, Trim$(Item(x).Name) & " pode ser usado por todas as classes.", Yellow)
                Else
                    Call PlayerMsg(Index, Trim$(Item(x).Name) & " pode apenas ser usado por um " & GetClassName(y) & ".", Yellow)
                End If
            End If

            If x < 1 Then
                z = z + 1
            End If

        Next
    Next

    Packet = Packet & END_CHAR

    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(Index, "Este shop não tem nada para vender!", BrightRed)
    Else
        Call SendDataTo(Index, Packet)
    End If

End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
    Dim Packet As String

    Packet = "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
    Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Speech & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Speech & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
    Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
    Dim Packet As String

    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherToAll()
    Dim i As Long
    Dim Weather As String

    Select Case GameWeather

        Case 0
            Weather = "Nenhum"

        Case 1
            Weather = "Chuva"

        Case 2
            Weather = "Neve"

        Case 3
            Weather = "Trovão"
    End Select

    frmServer.Label5.Caption = "Clima Atual: " & Weather

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If

    Next

End Sub

Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim N As Long, i As Long

    s = vbNullString
    N = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index Then
            s = s & GetPlayerName(i) & ", "
            N = N + 1
        End If

    Next

    If N = 0 Then
        s = "Não há jogadores online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "Existem " & N & " jogadores online: " & s & "."
    End If

    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim Packet As String

    If IsPlaying(Index) Then
        Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)

    If Index <> 0 Then

        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
        If Not IsBanned(GetPlayerIP(Index)) Then
            Call TextAdd(frmServer.txtText(0), "Recebeu uma conexão de " & GetPlayerIP(Index) & ".", True)
        Else
            Call AlertMsg(Index, "Você foi banido do " & GAME_NAME & ", e não pode mais jogar.")
        End If

    End If

End Sub

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " - Servidor"
    frmServer.lblIP.Caption = "Endereço de IP: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.Caption = "Porta: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.Caption = "Total de Jogadores Online: " & TotalOnlinePlayers
    Exit Sub
End Sub

' MODO DE SEGURANÇA -- "Descomente" para LIGÁ-LO, comente para DESLIGÁ-LO (whole function)
'Function Parse(ByVal index As Long) As String
'    If index > NumParse Then
'        Call HackingAttempt(ParseIndex, "Subscript out of range, " & ZePacket(0))
'        Exit Function
'    End If
'
'    Parse = ZePacket(index)

'End Function
