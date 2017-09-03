Attribute VB_Name = "modCombat"
Option Explicit

Public Function GetPlayerDamage(ByVal Itemnum As Long, Optional ByVal Index As Long = 0) As Long
Dim Damage As Long

    If Index > 0 Then
        Damage = ((Item(Itemnum).Damage * 9.5) + (Player(Index).Stat(Stats.Strength) / 2 + Player(Index).Stat(Stats.Agility) / 3) * 3.5)
    Else
        Damage = ((Item(Itemnum).Damage * 9.5) + (100 / 2 + 100 / 3) * 3.5)
    End If
    
    GetPlayerDamage = Damage

End Function

Public Function GetNpcDamage(ByVal NpcNum As Long)
    GetNpcDamage = ((Npc(NpcNum).Damage * 9.5) + (Npc(NpcNum).Stat(Stats.Strength) / 2 + Npc(NpcNum).Stat(Stats.Agility) / 3) * 3.5)
End Function

Public Function CheckDied(ByVal Index As Long) As Boolean

    CheckDied = False
    If Player(Index).Vital(Vitals.Health) <= 0 Then
        Call OnDeath(Index)
        CheckDied = True
    End If
    
End Function

Public Sub CheckNpcDied(ByVal Index As Long)

    If Map(Player(MyIndex).Map).MapNpc(Index).Vital(Vitals.Health) <= 0 Then Call KillNpc(Player(MyIndex).Map, Index)

End Sub

Public Sub KillNpc(ByVal MapIndex As Long, ByVal NpcIndex As Long)

    With TempNpc(MapIndex).NpcNum(NpcIndex)
        .Y = 0
        .X = 0
        .Alive = False
    End With

End Sub

Public Sub OnDeath(ByVal Index As Long)

    With Player(Index)
        'check for boot map
        Call WarpPlayer(Index, START_MAP, START_X, START_Y)
        
        .Vital(Vitals.Health) = GetPlayerMaxVital(Index, Vitals.Health)
        .Vital(Vitals.Spirit) = GetPlayerMaxVital(Index, Vitals.Spirit)
        
        ' Make a grave with the items. Dropping on the floor won't work.
    End With
    
    Call UpdatePlayerVitals(Index)
    Call AddText("Oh dear! You died!", BrightRed)

End Sub

Public Function GetPlayerMaxVital(ByVal Index As Long, ByVal VitalType As Byte)
    Select Case VitalType
        Case Vitals.Health
            GetPlayerMaxVital = (((Player(Index).Combat.level / 2) + Player(Index).Stat(Stats.Defense) / 1.25 + Player(Index).Stat(Stats.Agility) / 2) + 20) * 15
        Case Vitals.Spirit
            GetPlayerMaxVital = (((Player(Index).Combat.level / 2) + Player(Index).Stat(Stats.Sagacity) / 1.25 + Player(Index).Stat(Stats.Agility) / 2) + 20) * 15
    End Select
End Function

Public Sub PlayerAction()
Dim TileX As Byte, TileY As Byte
Dim Mapnum As Long
Dim I As Long

    With Player(MyIndex)
        TileX = .X
        TileY = .Y
        Mapnum = .Map
        Select Case .Dir
            Case DIR_UP
                TileY = TileY - 1
            Case DIR_DOWN
                TileY = TileY + 1
            Case DIR_RIGHT
                TileX = TileX + 1
            Case DIR_LEFT
                TileX = TileX - 1
        End Select
        
        If TileX < 1 Or TileX > MAX_MAP_X Then Exit Sub
        If TileY < 1 Or TileY > MAX_MAP_Y Then Exit Sub
    End With
    
    ' Check for NPC's
    For I = 1 To MAX_MAP_NPCS
        If TempNpc(Mapnum).NpcNum(I).X = TileX And TempNpc(Mapnum).NpcNum(I).Y = TileY Then
            If CanPlayerAttackNpc(Mapnum, I) Then
                TempPlayer(MyIndex).Step = 2
                Call PlayerAttackNpc(MyIndex, Mapnum, I)
                Exit Sub
            End If
        End If
    Next
    
    ' Check for Resources
    If MapResource(Mapnum).Tile(TileX, TileY).Num > 0 Then
        Call AttackResource(Mapnum, TileX, TileY)
        Exit Sub
    End If
    
    Select Case Map(Mapnum).Tile(TileX, TileY).Attribute
        Case Attributes.ChestTile
            If OpenChest(TileX, TileY) = False Then Call PopulateChest(Mapnum, TileX, TileY)
            frmMain.picChest.Visible = True
            Call RenderChest(Map(Mapnum).Tile(TileX, TileY).LongValue(1), TileX, TileY)
    End Select
    
    ' Last resort. Try projectiles!
    If Player(MyIndex).Equipment(Equipment.Weapon).Num > 0 Then
        If Item(Player(MyIndex).Equipment(Equipment.Weapon).Num).Projectile.Range > 0 Then
            With Item(Player(MyIndex).Equipment(Equipment.Weapon).Num)
                Call CreateProjectile(Mapnum, .Projectile.Image, .Projectile.Speed, .Projectile.Range, Player(MyIndex).Dir, TileX, TileY)
                TempPlayer(MyIndex).AttackTimer = .Speed
                TempPlayer(MyIndex).CombatTimer = 10
            End With
        End If
    End If
    
End Sub

Public Function CanPlayerAttackNpc(ByVal Mapnum As Long, ByVal Num As Long) As Boolean
Dim NpcNum As Long

    CanPlayerAttackNpc = False
    NpcNum = Map(Mapnum).MapNpc(Num).Num
    
    If Npc(NpcNum).Type = NPC_TYPE_FRIENDLY Then
        Exit Function
    ElseIf Npc(NpcNum).Type = NPC_TYPE_STATIONARY Then
        Exit Function
    End If

    CanPlayerAttackNpc = True
    Exit Function

End Function

Public Function CanNpcAttackPlayer(ByVal Mapnum As Long, ByVal Num As Long, ByVal Index As Long) As Boolean
Dim NpcNum As Long

    CanNpcAttackPlayer = False
    NpcNum = Map(Mapnum).MapNpc(Num).Num
    If Mapnum <> Player(Index).Map Then
        TempNpc(Mapnum).NpcNum(Num).Target = 0
        Exit Function
    End If
    
    If Npc(NpcNum).Type = NPC_TYPE_FRIENDLY Or Npc(NpcNum).Type = NPC_TYPE_STATIONARY Then Exit Function
    
    CanNpcAttackPlayer = True
    Exit Function
    
End Function

Public Sub RespawnNpc(ByVal MapIndex As Long, ByVal NpcIndex As Long)
Dim NpcNum As Long

    NpcNum = Map(MapIndex).MapNpc(NpcIndex).Num

    With TempNpc(MapIndex).NpcNum(NpcIndex)
        .Y = Map(MapIndex).MapNpc(NpcIndex).SpawnY
        .X = Map(MapIndex).MapNpc(NpcIndex).SpawnX
        Map(MapIndex).MapNpc(NpcIndex).Dir = DIR_DOWN
        .YOffset = 0
        .XOffset = 0
        .Moving = 0
        .CombatTimer = 0
        .Alive = True
        .Target = 0
    End With
    
    Map(MapIndex).MapNpc(NpcIndex).Vital(Vitals.Spirit) = Npc(NpcNum).Vital(Vitals.Spirit)
    Map(MapIndex).MapNpc(NpcIndex).Vital(Vitals.Health) = Npc(NpcNum).Vital(Vitals.Health)
    
End Sub

Public Sub PlayerAttackNpc(ByVal PlayerIndex As Long, ByVal Mapnum As Long, ByVal Index As Long)
Dim Damage As Long
Dim Tick As Long
Dim NpcNum As Long
Dim Offense As Long, Defense As Long, COA As Long
Dim I As Long, Chance As Long, Layer As Long

    NpcNum = Map(Mapnum).MapNpc(Index).Num
    TempNpc(Mapnum).NpcNum(Index).CombatTimer = 10
    TempNpc(Mapnum).NpcNum(Index).Target = PlayerIndex
    TempPlayer(PlayerIndex).CombatTimer = 10
    
      If Player(MyIndex).Equipment(Equipment.Weapon).Num > 0 Then
            Damage = GetPlayerDamage(Player(MyIndex).Equipment(Equipment.Weapon).Num, PlayerIndex)
            Tick = Item(Player(MyIndex).Equipment(Equipment.Weapon).Num).Speed
        Else
            Damage = 100
            Tick = 1000
        End If
    
    TempPlayer(PlayerIndex).AttackTimer = Tick
    
    With Player(PlayerIndex)
        If .Equipment(Equipment.Weapon).Num > 0 Then
            Select Case Item(.Equipment(Equipment.Weapon).Num).CombatType
                Case Combat.Melee
                    Offense = GetPlayerOffense(Combat.Melee)
                    Defense = Npc(NpcNum).Defense(Combat.Melee)
                Case Combat.Ranged
                    Offense = GetPlayerOffense(Combat.Ranged)
                    Defense = Npc(NpcNum).Defense(Combat.Ranged)
                Case Combat.Magic
                    Offense = GetPlayerOffense(Combat.Magic)
                    Defense = Npc(NpcNum).Defense(Combat.Magic)
            End Select
        Else
            Offense = 0
            Defense = Npc(NpcNum).Defense(Combat.Melee)
        End If
        COA = (.Stat(Stats.Attack) / 4) + (Offense / 6.5)
        Defense = (Npc(NpcNum).Stat(Stats.Defense) / 4) + (Defense / 6.5)
        COA = RAND(1, COA)
        Defense = RAND(1, Defense)
        If COA < Defense Then
            Call ActionMsg("Blocked!", TempNpc(Mapnum).NpcNum(Index).X, TempNpc(Mapnum).NpcNum(Index).Y, BrightRed)
            Exit Sub
        End If
        
        If CanNpcDodge(NpcNum) Then
            Call ActionMsg("Dodge!", TempNpc(Mapnum).NpcNum(Index).X, TempNpc(Mapnum).NpcNum(Index).Y, Pink)
            Exit Sub
        End If
        
        If CanNpcParry(NpcNum) Then
            Call ActionMsg("Parry!", TempNpc(Mapnum).NpcNum(Index).X, TempNpc(Mapnum).NpcNum(Index).Y, Pink)
            Exit Sub
        End If
        
        Damage = RAND(0, Damage)
        
        If CanPlayerCrit Then
            Damage = Damage * 1.5
            Call ActionMsg("Critical!", TempNpc(Mapnum).NpcNum(Index).X + 1, TempNpc(Mapnum).NpcNum(Index).Y, BrightCyan)
        End If
        
        If Damage > 0 Then
            Call ActionMsg(Damage, TempNpc(Mapnum).NpcNum(Index).X, TempNpc(Mapnum).NpcNum(Index).Y, BrightRed)
            Call UpdateNpcVitals(Mapnum, Index, -Damage)
            
            ' We killed the NPC
            If Map(Mapnum).MapNpc(Index).Vital(Vitals.Health) <= 0 Then
                Call KillNpc(Mapnum, Index)
                If Npc(NpcNum).XP > 0 Then
                    Call ActionMsg("+ " & Npc(NpcNum).XP, Player(PlayerIndex).X, Player(PlayerIndex).Y, White)
                    Call GiveXP(Npc(NpcNum).XP)
                End If
                
                For I = 1 To MAX_DROPS
                    If Npc(NpcNum).Drop(I).Item > 0 Then
                        Chance = RAND(0, 100)
                        If Chance <= Npc(NpcNum).Drop(I).Chance Then
                            Layer = FindOpenLayerSlot(TempNpc(Mapnum).NpcNum(Index).X, TempNpc(Mapnum).NpcNum(Index).Y)
                            If Layer > 0 Then
                                With MapItem(Mapnum).Tile(TempNpc(Mapnum).NpcNum(Index).X, TempNpc(Mapnum).NpcNum(Index).Y).Layer(Layer)
                                    .Num = Npc(NpcNum).Drop(I).Item
                                    .Value = Npc(NpcNum).Drop(I).Value
                                    .Tick = Default_Map_Item_Appear
                                End With
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Sub NpcAttackPlayer(ByVal Mapnum As Long, ByVal Num As Long, ByVal Index As Long)
Dim Damage As Long
Dim Tick As Long
Dim NpcNum As Long
Dim Offense As Long, Defense As Long
Dim COA As Long

    ' Something screwed up, so clear the relevant info, and forget about it.
    If Index = 0 Then
        TempNpc(Mapnum).NpcNum(Num).Target = 0
        TempNpc(Mapnum).NpcNum(Num).AttackTimer = 0
        Call AddText("A bug just happened. What did you just do?....", BrightRed)
        Exit Sub
    End If

    TempNpc(Mapnum).NpcNum(Num).CombatTimer = 10
    TempPlayer(Index).CombatTimer = 10
    NpcNum = Map(Mapnum).MapNpc(Num).Num
    
    With Npc(NpcNum)
        Damage = GetNpcDamage(NpcNum)
        Tick = .AttackSpeed * 1000
        TempNpc(Mapnum).NpcNum(Num).AttackTimer = Tick
        TempNpc(Mapnum).NpcNum(Num).Step = 2
        
        Select Case Npc(NpcNum).AttackType
            Case Combat.Melee
                Offense = Npc(NpcNum).Offense(Combat.Melee)
                Defense = GetPlayerDefense(Combat.Melee)
            Case Combat.Ranged
                Offense = Npc(NpcNum).Offense(Combat.Ranged)
                Defense = GetPlayerDefense(Combat.Ranged)
            Case Combat.Magic
                Offense = Npc(NpcNum).Offense(Combat.Magic)
                Defense = GetPlayerDefense(Combat.Magic)
        End Select
        COA = (.Stat(Stats.Attack) / 4) + (Offense / 6.5)
        Defense = (Player(Index).Stat(Stats.Defense) / 4) + (Defense / 6.5)
        COA = RAND(1, COA)
        Defense = RAND(1, Defense)
        If COA <= Defense Then
            Call ActionMsg("Blocked!", Player(MyIndex).X, Player(MyIndex).Y, BrightRed)
            Exit Sub
        End If
        
        If CanPlayerDodge Then
            Call ActionMsg("Dodged!", Player(MyIndex).X, Player(MyIndex).Y, Pink)
            Exit Sub
        End If
        
        If CanPlayerParry Then
            Call ActionMsg("Parry!", Player(MyIndex).X, Player(MyIndex).Y, Pink)
            Exit Sub
        End If
        
        Damage = RAND(0, Damage)
        
        If CanNpcCrit(NpcNum) Then
            Damage = Damage * 1.5
            Call ActionMsg("Crit!", Player(MyIndex).X + 1, Player(MyIndex).Y, BrightCyan)
        End If
        
        If Damage > 0 Then
            Call ActionMsg(Damage, Player(Index).X, Player(Index).Y, BrightRed)
            Call UpdatePlayerVitals(Index, -Damage)
            If CheckDied(Index) Then TempNpc(Mapnum).NpcNum(Num).Target = 0
        End If
    End With
        
    
End Sub

Public Sub UpdatePlayerVitals(ByVal Index As Long, Optional ByVal HP As Long, Optional MP As Long)
Dim MaxHP As Long
Dim MaxMP As Long
Dim MaxXP As Long
Dim v As Byte
    
    Player(Index).Vital(Vitals.Health) = Player(Index).Vital(Vitals.Health) + HP
    Player(Index).Vital(Vitals.Spirit) = Player(Index).Vital(Vitals.Spirit) + MP
    If Player(MyIndex).Vital(Vitals.Health) > GetPlayerMaxVital(MyIndex, Vitals.Health) Then Player(MyIndex).Vital(Vitals.Health) = GetPlayerMaxVital(MyIndex, Vitals.Health)
    If Player(MyIndex).Vital(Vitals.Spirit) > GetPlayerMaxVital(MyIndex, Vitals.Spirit) Then Player(MyIndex).Vital(Vitals.Spirit) = GetPlayerMaxVital(MyIndex, Vitals.Spirit)

    For v = 1 To Vitals.Vital_Count - 1
        If Player(Index).Vital(v) > GetPlayerMaxVital(Index, v) Then
            Player(Index).Vital(v) = GetPlayerMaxVital(Index, v)
        End If
    Next
    
    If Index = MyIndex Then
        MaxHP = GetPlayerMaxVital(MyIndex, Vitals.Health)
        MaxMP = GetPlayerMaxVital(MyIndex, Vitals.Spirit)
        MaxXP = GetPlayerNextLevelXP(Player(MyIndex).Combat.level)
        
        frmMain.lblHp.Caption = Player(MyIndex).Vital(Vitals.Health) & " / " & MaxHP
        frmMain.lblMP.Caption = Player(MyIndex).Vital(Vitals.Spirit) & " / " & MaxMP
        frmMain.lblXP.Caption = Player(MyIndex).Combat.XP & " / " & MaxXP
        
        MaxHP = ((Player(MyIndex).Vital(Vitals.Health) / (MaxHP / 100) / 100) * Vital_Bar_Width)
        MaxMP = ((Player(MyIndex).Vital(Vitals.Spirit) / (MaxMP / 100) / 100) * Vital_Bar_Width)
        MaxXP = ((Player(MyIndex).Combat.XP / (MaxXP / 100) / 100) * Vital_Bar_Width)
        
        If MaxHP < 0 Then MaxHP = 0
        If MaxMP < 0 Then MaxMP = 0
        frmMain.imgHp.Width = MaxHP
        frmMain.imgSp.Width = MaxMP
        frmMain.imgXp.Width = MaxXP
    End If
    
End Sub

Public Sub UpdateNpcVitals(ByVal Mapnum As Long, ByVal Index As Long, Optional ByVal HP As Long = 0, Optional ByVal MP As Long = 0)
Dim MaxHP As Long, MaxMP As Long, MaxXP As Long, v As Byte
Dim NpcNum As Long

    NpcNum = Map(Mapnum).MapNpc(Index).Num

    With Map(Mapnum).MapNpc(Index)
        
        .Vital(Vitals.Health) = .Vital(Vitals.Health) + HP
        .Vital(Vitals.Spirit) = .Vital(Vitals.Spirit) + MP
        
        If .Vital(Vitals.Health) > Npc(NpcNum).Vital(Vitals.Health) Then .Vital(Vitals.Health) = Npc(NpcNum).Vital(Vitals.Health)
        If .Vital(Vitals.Spirit) > Npc(NpcNum).Vital(Vitals.Spirit) Then .Vital(Vitals.Spirit) = Npc(NpcNum).Vital(Vitals.Spirit)
        
    End With

End Sub

Public Sub AttackResource(ByVal Mapnum As Long, ByVal X As Long, ByVal Y As Long)
Dim Damage As Long
Dim Tick As Long
Dim GiveAmount As Long
Dim I As Long

    With MapResource(Mapnum).Tile(X, Y)
        If .Alive = False Then Exit Sub
        
        If Player(MyIndex).Equipment(Equipment.Weapon).Num > 0 Then
            Damage = GetPlayerDamage(Player(MyIndex).Equipment(Equipment.Weapon).Num)
            Tick = Item(Player(MyIndex).Equipment(Equipment.Weapon).Num).Speed
        Else
            Damage = 100
            Tick = 1000
        End If
        
        TempPlayer(MyIndex).AttackTimer = Tick
        
        ' Can they attack the resource?
        With Resource(.Num)
            For I = 1 To Skills.Skill_Count - 1
                If .RequiredXP(I) > 0 Then
                    If Player(MyIndex).Skill(I).level < .RequiredXP(I) Then
                        Call AddText("You do not have the " & GetSkillName(I) & " level to use this resource", BrightRed)
                        Exit Sub
                    End If
                End If
            Next
        
            If .EquipmentType > 0 Then
                If Player(MyIndex).Equipment(Equipment.Weapon).Num > 0 Then
                    If Item(Player(MyIndex).Equipment(Equipment.Weapon).Num).EquipmentType <> .EquipmentType Then
                        Call AddText("You need a different weapon to attack this resource!", BrightRed)
                        Exit Sub
                    End If
                Else
                    Call AddText("You need a weapon to attack this resource!", BrightRed)
                    Exit Sub
                End If
            End If
        End With
        
        Damage = RAND(1, Damage)
        Call ActionMsg(Damage, X, Y, BrightRed)
        .Health = .Health - Damage
        
        ' Kill the resource
        If .Health < 0 Then
            .Alive = False
            If Resource(.Num).Reward > 0 Then
                If Item(Resource(.Num).Reward).Stackable = True Then
                    GiveAmount = Resource(.Num).RewardValue
                    If GiveAmount = 0 Then GiveAmount = 1
                Else
                    GiveAmount = 1
                End If
                
                If FindOpenInvSlot(Resource(.Num).Reward) > 0 Then
                    Call GivePlayerItem(Resource(.Num).Reward, GiveAmount)
                    Call BltInventory
                End If
            End If
            
            ' XP reward
            With Resource(.Num)
                For I = 1 To Skills.Skill_Count - 1
                    If .RewardXP(I) > 0 Then
                        Call GiveSkillXp(I, .RewardXP(I))
                    End If
                Next
            End With
        End If
    End With

End Sub

Public Function IsNpcNextToTarget(ByVal Mapnum As Long, ByVal Index As Long, ByVal Target As Long) As Boolean
Dim X As Long
Dim Y As Long

    If Target = 0 Then Exit Function

    IsNpcNextToTarget = True
    
    If Player(Target).Map <> Mapnum Then
        Exit Function
    End If
    
    X = Player(Target).X
    Y = Player(Target).Y
    
    
    With TempNpc(Mapnum).NpcNum(Index)
    
        ' Is the player to the right of the npc?
        If .Y = Y And .X = X - 1 Then
            If Map(Mapnum).MapNpc(Index).Dir <> DIR_RIGHT Then Map(Mapnum).MapNpc(Index).Dir = DIR_RIGHT
            Exit Function
        End If
        
        ' Is the player to the left of the npc?
        If .Y = Y And .X = X + 1 Then
            If Map(Mapnum).MapNpc(Index).Dir <> DIR_LEFT Then Map(Mapnum).MapNpc(Index).Dir = DIR_LEFT
            Exit Function
        End If
        
        ' Is the player is above the npc?
        If .X = X And .Y = Y + 1 Then
            If Map(Mapnum).MapNpc(Index).Dir <> DIR_UP Then Map(Mapnum).MapNpc(Index).Dir = DIR_UP
            Exit Function
        End If
        
        ' Is the player under the npc?
        If .X = X And .Y = Y - 1 Then
            If Map(Mapnum).MapNpc(Index).Dir <> DIR_DOWN Then Map(Mapnum).MapNpc(Index).Dir = DIR_DOWN
            Exit Function
        End If
        
        IsNpcNextToTarget = False
    End With

End Function

Public Function NpcMoveToTarget(ByVal Mapnum As Long, ByVal Index As Long, ByVal Target As Long) As Byte
Dim X As Long, Y As Long

    If Player(Target).Map <> Mapnum Then
        TempNpc(Mapnum).NpcNum(Index).Target = 0
        Exit Function
    End If
    
    X = Player(Target).X
    Y = Player(Target).Y
    
    With TempNpc(Mapnum).NpcNum(Index)
    
        If .X < X Then
            If CheckNPCTile(Index, DIR_RIGHT, Mapnum) Then
                NpcMoveToTarget = DIR_RIGHT
                Exit Function
            End If
        End If
        
        If .X > X Then
            If CheckNPCTile(Index, DIR_LEFT, Mapnum) Then
                NpcMoveToTarget = DIR_LEFT
                Exit Function
            End If
        End If
        
        If .Y < Y Then
            If CheckNPCTile(Index, DIR_DOWN, Mapnum) Then
                NpcMoveToTarget = DIR_DOWN
                Exit Function
            End If
        End If
        
        If .Y > Y Then
            If CheckNPCTile(Index, DIR_UP, Mapnum) Then
                NpcMoveToTarget = DIR_UP
                Exit Function
            End If
        End If
    End With
    
End Function

Public Function GetPlayerOffense(ByVal CombatType As Byte) As Long
Dim I As Long, Amount As Long
    Amount = 0
    For I = 1 To Equipment.Equipment_Count - 1
        If Player(MyIndex).Equipment(I).Num > 0 Then
            Amount = Amount + Item(Player(MyIndex).Equipment(I).Num).Offense(CombatType)
        End If
    Next
    GetPlayerOffense = Amount
End Function

Public Function GetPlayerDefense(ByVal CombatType As Byte) As Long
Dim I As Long, Amount As Long
    Amount = 0
    For I = 1 To Equipment.Equipment_Count - 1
        If Player(MyIndex).Equipment(I).Num > 0 Then
            Amount = Amount + Item(Player(MyIndex).Equipment(I).Num).Defense(CombatType)
        End If
    Next
    GetPlayerDefense = Amount
End Function

Public Sub CastSpell(ByVal PlayerIndex As Long, ByVal Slot As Long)
Dim SpellNum As Long, Mapnum As Long
Dim PlayerX As Long, PlayerY As Long
Dim NPCX As Long, NPCY As Long
Dim DifX As Long, DifY As Long
Dim Target As Long, Range As Long
Dim MinX As Long, MinY As Long
Dim MaxX As Long, MaxY As Long
Dim I As Long, AOE As Long
Dim VitalChange(1 To 2) As Long

    Mapnum = Player(PlayerIndex).Map
    PlayerX = Player(PlayerIndex).X
    PlayerY = Player(PlayerIndex).Y
    SpellNum = Player(PlayerIndex).PlayerSpell(Slot).Num
    
    ' Has it cooled down yet?
    If Player(PlayerIndex).PlayerSpell(Slot).CoolDownTimer > 0 Then
        Call AddText("Spell hasn't cooled down yet!", BrightRed)
        Exit Sub
    End If

    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_VITAL_AFFECT
            ' Do we need a target?
            Range = Spell(SpellNum).Range
            If Range > 0 Then
                If TempPlayer(PlayerIndex).Target = 0 Then
                    Call AddText("You need a target to cast this spell!", BrightRed)
                    Exit Sub
                Else ' We have a target
                    Target = TempPlayer(PlayerIndex).Target
                    NPCX = TempNpc(Mapnum).NpcNum(Target).X
                    NPCY = TempNpc(Mapnum).NpcNum(Target).Y
                    ' Can we attack it?
                    Select Case Npc(Map(Mapnum).MapNpc(Target).Num).Type
                        Case NPC_TYPE_FRIENDLY, NPC_TYPE_STATIONARY ' Nope.
                            Call AddText("You cannot attack this npc!", BrightRed)
                            Exit Sub
                        Case Else ' Yeah we can.
                        
                            ' This might seem stupid, but is our target dead?
                            If TempNpc(Mapnum).NpcNum(TempPlayer(PlayerIndex).Target).Alive = False Then
                                Exit Sub
                            End If
                            
                            ' Are we within range?
                            DifX = PlayerX - NPCX
                            DifY = PlayerY - NPCY
                            
                            DifX = DifX ^ 2
                            DifY = DifY ^ 2
                            DifX = Sqr(DifX)
                            DifY = Sqr(DifY)
                            
                            If DifX <= Range And DifY <= Range Then
                                VitalChange(Vitals.Health) = Spell(SpellNum).VitalAffect(Vitals.Health)
                                VitalChange(Vitals.Spirit) = Spell(SpellNum).VitalAffect(Vitals.Spirit)
                                
                                Call UpdateNpcVitals(Mapnum, Target, VitalChange(Vitals.Health), VitalChange(Vitals.Spirit))
                                
            If Map(Mapnum).MapNpc(Target).Vital(Vitals.Health) <= 0 Then
                Call KillNpc(Mapnum, Target)
                If Npc(Map(Mapnum).MapNpc(Target).Num).XP > 0 Then
                    Call ActionMsg("+ " & Npc(Map(Mapnum).MapNpc(Target).Num).XP, Player(PlayerIndex).X, Player(PlayerIndex).Y, White)
                    Call GiveXP(Npc(Map(Mapnum).MapNpc(Target).Num).XP)
                End If
            End If
                                
                                TempNpc(Mapnum).NpcNum(Target).CombatTimer = 10
                                TempPlayer(PlayerIndex).CombatTimer = 10
                                TempNpc(Mapnum).NpcNum(Target).StunDuration = Spell(SpellNum).StunDuration
                                If VitalChange(Vitals.Health) < 0 Or VitalChange(Vitals.Spirit) < 0 Then
                                    TempNpc(Mapnum).NpcNum(Target).Target = PlayerIndex
                                    Call ActionMsg("-" & VitalChange(Vitals.Health), NPCX, NPCY, BrightRed)
                                End If
                                
                                ' AOE
                                If Spell(SpellNum).AOE > 0 Then
                                    AOE = Spell(SpellNum).AOE
                                    MinX = NPCX - AOE
                                    MaxX = NPCX + AOE
                                    MinY = NPCY - AOE
                                    MaxY = NPCY + AOE
                                    For I = 1 To MAX_MAP_NPCS
                                        If I <> Target Then
                                        If TempNpc(Mapnum).NpcNum(I).X >= MinX And TempNpc(Mapnum).NpcNum(I).X <= MaxX Then
                                            If TempNpc(Mapnum).NpcNum(I).Y >= MinY And TempNpc(Mapnum).NpcNum(I).Y <= MaxY Then
                                                Call UpdateNpcVitals(Mapnum, I, VitalChange(Vitals.Health), VitalChange(Vitals.Spirit))
                                                TempNpc(Mapnum).NpcNum(I).CombatTimer = 10
                                                TempNpc(Mapnum).NpcNum(I).StunDuration = Spell(SpellNum).StunDuration
                                                If VitalChange(Vitals.Health) < 0 Or VitalChange(Vitals.Spirit) < 0 Then
                                                    TempNpc(Mapnum).NpcNum(I).Target = PlayerIndex
                                                    Call ActionMsg("-" & VitalChange(Vitals.Health), TempNpc(Mapnum).NpcNum(I).X, TempNpc(Mapnum).NpcNum(I).Y, BrightRed)
                                                End If
                                            End If
                                        End If
                                        End If
                                    Next
                                End If
                            Else
                                Call AddText("Not in range! " & DifX & " " & DifY, BrightRed)
                                Exit Sub
                            End If
                    End Select
                End If
            Else ' Self cast
            
            End If
        Case SPELL_TYPE_WARP
            If Spell(SpellNum).Map <> 0 And Spell(SpellNum).X <> 0 And Spell(SpellNum).Y <> 0 Then
                Call WarpPlayer(PlayerIndex, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y)
                Call AddText("You are wisked away by the magic of the spell.", BrightGreen)
            End If
        Case SPELL_TYPE_CUSTOM
            ' todo lol
    End Select
    
    Player(PlayerIndex).PlayerSpell(Slot).CoolDownTimer = Spell(SpellNum).CoolDown
    Call BltSpells
    
End Sub

Public Function CheckProjHit(ByVal Mapnum As Long, ByVal Num As Long) As Boolean
Dim X As Long, Y As Long
Dim I As Long

    With MapProjectile(Mapnum).MapProjectile(Num)
        X = .X
        Y = .Y
        For I = 1 To MAX_MAP_NPCS
            If TempNpc(Mapnum).NpcNum(I).X = X And TempNpc(Mapnum).NpcNum(I).Y = Y Then
                If Npc(Map(Mapnum).MapNpc(I).Num).Type <> NPC_TYPE_FRIENDLY Or Npc(Map(Mapnum).MapNpc(I).Num).Type <> NPC_TYPE_STATIONARY Then
                    TempPlayer(MyIndex).Step = 2
                    Call PlayerAttackNpc(MyIndex, Mapnum, I)
                    CheckProjHit = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Public Sub GiveXP(ByVal Amount As Long)
    
    With Player(MyIndex).Combat
        If .XP + Amount >= GetPlayerNextLevelXP(.level) Then
            If .level < MAX_COMBAT_LEVEL Then
                .XP = (GetPlayerNextLevelXP(.level) - .XP - Amount) * -1
                .level = .level + 1
                Player(MyIndex).Points = Player(MyIndex).Points + 3
                Call AddText("You leveled up! Congrats!", White)
            End If
        Else
            .XP = .XP + Amount
        End If
        Call UpdatePlayerVitals(MyIndex)
    End With
End Sub

Public Function CanNpcParry(ByVal NpcNum As Long) As Boolean
Dim Chance As Long
    Chance = Npc(NpcNum).Stat(Stats.Strength) * 0.25
    If RAND(1, 100) <= Chance Then CanNpcParry = True
End Function
Public Function CanPlayerParry() As Boolean
Dim Chance As Long
    Chance = Player(MyIndex).Stat(Stats.Strength) * 0.25
    If RAND(1, 100) <= Chance Then CanPlayerParry = True
End Function
Public Function CanNpcDodge(ByVal NpcNum As Long) As Boolean
Dim Chance As Long
    Chance = Npc(NpcNum).Stat(Stats.Agility) * 0.25
    If RAND(1, 100) <= Chance Then CanNpcDodge = True
End Function
Public Function CanPlayerDodge() As Boolean
Dim Chance As Long
    Chance = Player(MyIndex).Stat(Stats.Agility) * 0.25
    If RAND(1, 100) <= Chance Then CanPlayerDodge = True
End Function
Public Function CanNpcCrit(ByVal NpcNum As Long) As Boolean
Dim Chance As Long
    With Npc(NpcNum)
        Chance = (.Stat(Stats.Strength) + .Stat(Stats.Agility) + .Stat(Stats.Attack) + .Stat(Stats.Sagacity)) / 4
        Chance = Chance * 0.25
        If RAND(1, 100) <= Chance Then CanNpcCrit = True
    End With
End Function
Public Function CanPlayerCrit() As Boolean
Dim Chance As Long
    Chance = (Player(MyIndex).Stat(Stats.Strength) + Player(MyIndex).Stat(Stats.Agility) + Player(MyIndex).Stat(Stats.Attack) + Player(MyIndex).Stat(Stats.Sagacity)) / 4
    Chance = Chance * 0.25
    If RAND(1, 100) <= Chance Then CanPlayerCrit = True
End Function
