Attribute VB_Name = "modGameLoop"
Option Explicit

Public Sub GameLoop()
Dim Tick As Long
Dim tmr1000 As Long, tmr500 As Long, tmr250 As Long, tmr50 As Long, tmr30 As Long, tmr25 As Long
' 1 second, half a second, a fourth of a second, a 20th of a second
Dim SoundTick(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As Long, HighX As Long, LowX As Long, HighY As Long, LowY As Long, SoundRandom As Long
Dim i As Long, d As Byte, L As Long
Dim X As Long, Y As Long
Dim LastOnMap(1 To MAX_MAPS) As Long, Mapnum As Long
Dim MapItemTick(1 To MAX_MAP_X, 1 To MAX_MAP_Y, 0 To MAX_MAP_ITEM_LAYERS) As Long
Dim MapNpcRespawn(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As Long, NpcNum As Long
Dim MapResourceRespawn(1 To MAX_MAPS, 1 To MAX_MAP_X, 1 To MAX_MAP_Y) As Long
Dim Range As Byte

Dim RegenTick As Long

    frmMain.txtChat.text = vbNullString
    Call AddText("Welcome to Tequ!", Yellow, False)

    Do While Running = True
        Tick = timeGetTime
        
        ' Map Tile Sounds
        For X = 1 To MAX_MAP_X
            For Y = 1 To MAX_MAP_Y
                If Map(Player(MyIndex).Map).Tile(X, Y).Attribute = Attributes.Soundtile Then
                    If Tick > SoundTick(X, Y) Then
                        With Map(Player(MyIndex).Map).Tile(X, Y)
                            SoundRandom = .LongValue(3)
                            If RAND(1, SoundRandom) = 1 Then
                                HighY = Y + .LongValue(1)
                                LowY = Y - .LongValue(1)
                                HighX = X + .LongValue(2)
                                LowX = X - .LongValue(2)
                                If Player(MyIndex).X >= LowX And Player(MyIndex).X <= HighX Then
                                    If Player(MyIndex).Y >= LowY And Player(MyIndex).Y <= HighY Then
                                        Call PlaySound(.StringValue(1))
                                        SoundTick(X, Y) = Tick + .LongValue(4) * 1000
                                    End If
                                End If
                            Else
                                SoundTick(X, Y) = Tick + .LongValue(4) * 1000
                            End If
                        End With
                    End If
                End If
            Next
        Next
                            
        
        If Tick > tmr1000 Then
            
            If TempPlayer(MyIndex).DrawCPS = True Then frmMain.Caption = CPS
            CPS = 0
        
            If Options.OnlineMode = False Then
                For Mapnum = 1 To MAX_MAPS
                    'If we were on the map recently, then update the logic.
                    If LastOnMap(Mapnum) <= 60000 Then ' 60 second
                    
                    
                        '[UPDATING MAP ITEM TICKS]'
                        For X = 1 To MAX_MAP_X
                            For Y = 1 To MAX_MAP_Y
                                With MapItem(Mapnum).Tile(X, Y)
                                    For L = 1 To MAX_MAP_ITEM_LAYERS
                                        If .Layer(L).Num > 0 Then
                                            If Tick > MapItemTick(X, Y, L) Then
                                                Select Case .Layer(L).MapItemState
                                                    Case 0
                                                        MapItemTick(X, Y, L) = Tick + Default_Map_Item_Appear
                                                        .Layer(L).MapItemState = MAPITEMSTATE_Me
                                                    Case MAPITEMSTATE_Me
                                                        .Layer(L).MapItemState = MAPITEMSTATE_All
                                                        MapItemTick(X, Y, L) = Tick + Default_Map_Item_Despawn
                                                    Case MAPITEMSTATE_All
                                                        .Layer(L).Num = 0
                                                        .Layer(L).Tick = 0
                                                        .Layer(L).Value = 0
                                                        .Layer(L).MapItemState = 0
                                                End Select
                                            End If
                                        End If
                                    Next
                                End With
                            Next
                        Next
                        LastOnMap(Mapnum) = LastOnMap(Mapnum) + 1000
                        LastOnMap(Player(MyIndex).Map) = 0
                        
                        '[MAP ITEM ATTRIBUTE RESPAWNING]'
                        For X = 1 To MAX_MAP_X
                            For Y = 1 To MAX_MAP_Y
                                With Map(Mapnum).Tile(X, Y)
                                    If .Attribute = Attributes.ItemTile Then
                                        If Tick > MapItemTick(X, Y, 0) Then
                                            If .LongValue(1) > 0 Then
                                                MapItem(Mapnum).Tile(X, Y).Layer(0).Num = .LongValue(1)
                                                MapItem(Mapnum).Tile(X, Y).Layer(0).Value = .LongValue(2)
                                                MapItemTick(X, Y, 0) = Tick + .LongValue(3) * 1000
                                            End If
                                        End If
                                    End If
                                End With
                            Next
                        Next
                        
                        '[MAP NPC RESPAWNING]'
                        For X = 1 To MAX_MAP_NPCS
                            With TempNpc(Mapnum).NpcNum(X)
                                NpcNum = Map(Mapnum).MapNpc(X).Num
                                If NpcNum > 0 Then
                                    If .Alive = False Then
                                        If MapNpcRespawn(Mapnum, X) = 0 Then
                                            MapNpcRespawn(Mapnum, X) = Tick + Npc(NpcNum).Respawn
                                        Else
                                            If Tick > MapNpcRespawn(Mapnum, X) Then
                                                Call RespawnNpc(Mapnum, X)
                                                MapNpcRespawn(Mapnum, X) = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End With
                        Next
                        
                        '[MAP RESOURCE RESPAWNING]'
                        For X = 1 To MAX_MAP_X
                            For Y = 1 To MAX_MAP_Y
                                With MapResource(Mapnum).Tile(X, Y)
                                    If .Num > 0 Then
                                        If .Alive = False Then
                                            If MapResourceRespawn(Mapnum, X, Y) = 0 Then
                                                MapResourceRespawn(Mapnum, X, Y) = Tick + Resource(.Num).RespawnRate
                                            Else
                                                If Tick > MapResourceRespawn(Mapnum, X, Y) Then
                                                    Call RespawnResource(Mapnum, X, Y)
                                                    MapResourceRespawn(Mapnum, X, Y) = 0
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            Next
                        Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If TempNpc(Mapnum).NpcNum(i).CombatTimer >= 1 Then
                            TempNpc(Mapnum).NpcNum(i).CombatTimer = TempNpc(Mapnum).NpcNum(i).CombatTimer - 1
                        End If
                    Next
                        
                    End If
                Next 'END UPDATING MAP LOGIC ON RECENT MAPS'
                
                For i = 1 To MAX_PLAYERS
                    If TempPlayer(i).CombatTimer >= 1 Then
                        TempPlayer(i).CombatTimer = TempPlayer(i).CombatTimer - 1
                    End If
                Next
                
                '[VITALS REGEN]'
                RegenTick = RegenTick - 1000
                If RegenTick <= 0 Then
                    Call UpdatePlayerVitals(MyIndex, Player(MyIndex).Vital(Vitals.Health) * 0.1, Player(MyIndex).Vital(Vitals.Spirit) * 0.1)
                    RegenTick = 6000
                End If
                
            End If
                        
            tmr1000 = Tick + 1000
        End If
        
        If Tick > tmr500 Then
            
            tmr500 = Tick + 500
        End If
        
        If Tick > tmr250 Then
            
            For i = 1 To MAX_PLAYERS
                If TempPlayer(i).AttackTimer > 0 Then
                    TempPlayer(i).AttackTimer = TempPlayer(i).AttackTimer - 250
                    If TempPlayer(i).AttackTimer < 0 Then TempPlayer(i).AttackTimer = 0
                    TempPlayer(i).Step = 0
                End If
            Next
            
            For i = 1 To MAX_MAPS
                If LastOnMap(i) <= 60000 Then
                    For X = 1 To MAX_MAP_NPCS
                        If TempNpc(i).NpcNum(X).AttackTimer > 0 Then
                            TempNpc(i).NpcNum(X).AttackTimer = TempNpc(i).NpcNum(X).AttackTimer - 250
                            TempNpc(i).NpcNum(X).Step = 0
                        End If
                        If TempNpc(i).NpcNum(X).StunDuration > 0 Then
                            TempNpc(i).NpcNum(X).StunDuration = TempNpc(i).NpcNum(X).StunDuration - 250
                        End If
                    Next
                End If
            Next
            
            tmr250 = Tick + 250
        End If
        
        If Tick > tmr50 Then
            For X = 1 To MAX_PLAYERS
                For i = 1 To MAX_PLAYER_SPELLS
                    If Player(X).PlayerSpell(i).CoolDownTimer > 0 Then
                        Player(X).PlayerSpell(i).CoolDownTimer = Player(X).PlayerSpell(i).CoolDownTimer - 50
                        If Player(X).PlayerSpell(i).CoolDownTimer < 0 Then Player(X).PlayerSpell(i).CoolDownTimer = 0
                        Call BltSpells
                    End If
                Next
            Next
            
            tmr50 = Tick + 50
        End If

        If Tick > tmr25 Then
        
            Call DisableKeys ' Disable any key values that were enabled
        
            ' Set the focus to the picscreen only if...
            If frmMain.hWnd = GetActiveWindow() Then
                Call CheckKeys
                If frmMain.txtMyChat.Visible = False Then
                    If CreatingCharacter = False Then frmMain.picScreen.SetFocus
                Else
                    frmMain.txtMyChat.SetFocus
                End If
            End If
        
            tmr25 = Tick + 25
        End If
        
        If Tick > tmr30 Then
            For i = 1 To MAX_PLAYERS
                'If IsPlaying(i) Then
                    Call ProcessMovement(i)
                'End If
            Next
            
            ' NPC movement
            If Options.OnlineMode = False Then
                For Mapnum = 1 To MAX_MAPS
                    If LastOnMap(Mapnum) <= 60000 Then ' 60 seconds
                    
                    
                        For i = 1 To MAX_MAP_PROJECTILES
                            If MapProjectile(Mapnum).MapProjectile(i).Range > 0 Then
                                Call UpdateProjectileLogic(Mapnum, i)
                            End If
                        Next
                    
                    
                        For i = 1 To MAX_MAP_NPCS
                            If Map(Mapnum).MapNpc(i).Num > 0 Then
                                With TempNpc(Mapnum).NpcNum(i)
                                    If .Moving = 0 And .Alive = True Then
                                        If Npc(Map(Mapnum).MapNpc(i).Num).Type <> NPC_TYPE_STATIONARY Then
                                        
                                            ' Let's see if they want to get a target. If so, get it.
                                            If Npc(Map(Mapnum).MapNpc(i).Num).Type = NPC_TYPE_ATTACK_ON_SIGHT Then
                                                If .Target = 0 Then
                                                    Range = Npc(Map(Mapnum).MapNpc(i).Num).Range
                                                
                                                    For L = 1 To MAX_PLAYERS
                                                        If Player(L).Map = Mapnum Then
                                                            If Player(L).X >= .X - Range And Player(L).X <= .X + Range Then
                                                                If Player(L).Y >= .Y - Range And Player(L).Y <= .Y + Range And Player(L).Access <> ACCESS_ADMIN Then
                                                                    If L = MyIndex Then Call AddText("You sense something approaching...", Grey)
                                                                    .Target = L
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If
                                        
                                        
                                            If .Target > 0 Then
                                                If IsNpcNextToTarget(Mapnum, i, .Target) Then
                                                    ' attack the player if the attack timer says it's m'k
                                                    If .AttackTimer <= 0 Then
                                                        If CanNpcAttackPlayer(Mapnum, i, .Target) Then
                                                            Call NpcAttackPlayer(Mapnum, i, .Target)
                                                        End If
                                                    End If
                                                Else
                                                    .Moving = NpcMoveToTarget(Mapnum, i, .Target)
                                                    If .Moving <> 0 Then Call InitiateNPCMovement(i, .Moving, Mapnum)
                                                End If
                                            Else
                                                X = RAND(1, 100)
                                                If X > 98 Then
                                                    X = RAND(1, 4)
                                                    .Moving = X
                                                    Call InitiateNPCMovement(i, X, Mapnum)
                                                End If
                                                'find a dir
                                            End If
                                        End If
                                    End If
                                End With
                                Call ProcessNPCMovement(i, Mapnum)
                            End If
                        Next
                        
                        
                        
                    End If
                Next
                
                
                
            End If
            tmr30 = Tick + 30
        End If
            
        Call Render_Graphics
        
        If Options.OnlineMode = True Then
            If frmMain.socket.State <> 7 Then
                Call DestroyGame
                Call MsgBox("Disconnected from server.", vbCritical)
            End If
        End If
        
        CPS = CPS + 1
        DoEvents
        Sleep (1)
    Loop
        

End Sub
