Attribute VB_Name = "modGameLogic"
Option Explicit

Public Function CheckProjecTile(ByVal X As Long, ByVal Y As Long, ByVal Dir As Long) As Boolean
Dim Mapnum As Long
Dim i As Long

    Mapnum = Player(MyIndex).Map

    Select Case Dir
        Case DIR_UP
            Y = Y - 1
        Case DIR_DOWN
            Y = Y + 1
        Case DIR_LEFT
            X = X - 1
        Case DIR_RIGHT
            X = X + 1
    End Select
    
    With Map(Mapnum).Tile(X, Y)
        Select Case .Attribute
            Case Attributes.BlockedTile, Attributes.BankTile, Attributes.ChestTile, Attributes.ResourceTile
                CheckProjecTile = False
        End Select
    End With
    
End Function

Public Function CheckTile(ByVal Index As Long, ByVal Dir As Long) As Boolean
Dim X As Long, Y As Long
Dim Mapnum As Long
Dim Health As Long, Differ As Long
Dim i As Long

    CheckTile = True
    
    X = Player(Index).X
    Y = Player(Index).Y
    Mapnum = Player(Index).Map
    
    If frmMain.picChest.Visible = True Then frmMain.picChest.Visible = False
    frmMain.picBank.Visible = False
    Game.InBank = False
    frmMain.picShop.Visible = False
    Game.InShop = False
    
    Select Case Dir
        Case DIR_UP
            Y = Y - 1
        Case DIR_DOWN
            Y = Y + 1
        Case DIR_LEFT
            X = X - 1
        Case DIR_RIGHT
            X = X + 1
    End Select
    
    ' Check if out of bounds
    If X > MAX_MAP_X Or X < 1 Then
        If X > MAX_MAP_X Then
            If Map(Mapnum).RightWarp > 0 And Map(Mapnum).RightWarp <= MAX_MAPS Then
                Call WarpPlayer(Index, Map(Mapnum).RightWarp, 1, Y)
                CheckTile = True
                Exit Function
            End If
        ElseIf X < MAX_MAP_X Then
            If Map(Mapnum).LeftWarp > 0 And Map(Mapnum).LeftWarp <= MAX_MAPS Then
                Call WarpPlayer(Index, Map(Mapnum).LeftWarp, MAX_MAP_X, Y)
                CheckTile = True
                Exit Function
            End If
        End If
        CheckTile = False
        Exit Function
    End If
    If Y > MAX_MAP_Y Or Y < 1 Then
        If Y > MAX_MAP_Y Then
            If Map(Mapnum).DownWarp > 0 And Map(Mapnum).DownWarp <= MAX_MAPS Then
                Call WarpPlayer(Index, Map(Mapnum).DownWarp, X, 1)
                CheckTile = True
                Exit Function
            End If
        ElseIf Y < 1 Then
            If Map(Mapnum).UpWarp > 0 And Map(Mapnum).UpWarp <= MAX_MAPS Then
                Call WarpPlayer(Index, Map(Mapnum).UpWarp, X, MAX_MAP_Y)
                CheckTile = True
                Exit Function
            End If
        End If
        CheckTile = False
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_NPCS
        If TempNpc(Mapnum).NpcNum(i).X = X And TempNpc(Mapnum).NpcNum(i).Y = Y Then
            CheckTile = False
            Exit Function
        End If
    Next
    
    Select Case Map(Mapnum).Tile(X, Y).Attribute
        Case Attributes.BlockedTile ' blocked
            CheckTile = False
        Case Attributes.ChestTile ' blocked
            CheckTile = False
        Case Attributes.ResourceTile ' blocked
            CheckTile = False
        Case Attributes.HealTile
            Health = Player(Index).Vital(Vitals.Health)
            Differ = Map(Mapnum).Tile(X, Y).LongValue(1)
            Player(Index).Vital(Vitals.Health) = Health + Differ
            Call AddText("You feel healed.", BrightGreen)
            Call UpdatePlayerVitals(Index)
        Case Attributes.TrapTile
            Health = Player(Index).Vital(Vitals.Health)
            Differ = Map(Mapnum).Tile(X, Y).LongValue(1)
            Player(Index).Vital(Vitals.Health) = Health - Differ
            Call AddText("You stepped on a trap!", BrightRed)
            Call CheckDied(Index)
            Call UpdatePlayerVitals(Index)
        Case Attributes.KeyTile
            If TempPlayer(Index).UnlockedTile(X, Y) = False Then
                For i = 1 To MAX_INV
                    If Player(Index).Inv(i).Num = Map(Mapnum).Tile(X, Y).LongValue(1) Then
                        Call TakeInvItem(Index, i)
                        Call AddText("You unlock a tile.", Yellow)
                        TempPlayer(Index).UnlockedTile(X, Y) = True
                        Exit Function
                    End If
                Next
                CheckTile = False
                Exit Function
            End If
        Case Attributes.BankTile
            frmMain.picBank.Visible = True
            Call RenderBank
            Game.InBank = True
        Case Attributes.ShopTile
            Game.InShop = True
            Call RenderShop(Map(Mapnum).Tile(X, Y).LongValue(1))
    End Select
        
End Function

Public Function CheckNPCTile(ByVal Index As Long, ByVal Dir As Long, ByVal Mapnum As Long) As Boolean
Dim X As Long, Y As Long
Dim i As Long

    CheckNPCTile = True
    
    ' Check if stunned
    If TempNpc(Mapnum).NpcNum(Index).StunDuration > 0 Then
        CheckNPCTile = False
        Exit Function
    End If
    
    X = TempNpc(Mapnum).NpcNum(Index).X
    Y = TempNpc(Mapnum).NpcNum(Index).Y
    
    Select Case Dir
        Case DIR_UP
            Y = Y - 1
        Case DIR_DOWN
            Y = Y + 1
        Case DIR_LEFT
            X = X - 1
        Case DIR_RIGHT
            X = X + 1
    End Select
    
    ' Check if out of bounds
    If X > MAX_MAP_X Or X < 1 Then
        CheckNPCTile = False
        Exit Function
    End If
    If Y > MAX_MAP_Y Or Y < 1 Then
        CheckNPCTile = False
        Exit Function
    End If
    
    Select Case Map(Mapnum).Tile(X, Y).Attribute
        Case Attributes.BlockedTile, Attributes.HealTile, Attributes.KeyTile, Attributes.NpcAvoidTile, Attributes.TrapTile, Attributes.warptile, Attributes.ResourceTile
            CheckNPCTile = False
            Exit Function
    End Select
    
    For i = 1 To MAX_PLAYERS
        If Player(i).Map = Mapnum Then
            If Player(i).X = X And Player(i).Y = Y Then
                CheckNPCTile = False
                Exit Function
            End If
        End If
    Next
    
    For i = 1 To MAX_MAP_NPCS
        If i <> Index Then
            If TempNpc(Mapnum).NpcNum(i).X = X And TempNpc(Mapnum).NpcNum(i).Y = Y Then
                CheckNPCTile = False
                Exit Function
            End If
        End If
    Next

End Function

Sub ProcessMovement(ByVal Index As Long)
Dim Speed As Long
Dim X As Long

    With TempPlayer(Index)
        If .Moving > 0 Then
        
            If False = True Then
                Speed = 3
            Else
                Speed = 4
            End If
        
            Select Case .Moving
                Case DIR_UP
                    .YOffset = .YOffset - Speed
                    If .YOffset <= 0 Then
                        .Moving = 0
                        .YOffset = 0
                    End If
                Case DIR_DOWN
                    .YOffset = .YOffset + Speed
                    If .YOffset >= 0 Then
                        .Moving = 0
                        .YOffset = 0
                    End If
                Case DIR_RIGHT
                    .XOffset = .XOffset + Speed
                    If .XOffset >= 0 Then
                        .Moving = 0
                        .XOffset = 0
                    End If
                Case DIR_LEFT
                    .XOffset = .XOffset - Speed
                    If .XOffset <= 0 Then
                        .Moving = 0
                        .XOffset = 0
                    End If
            End Select

            If Index = MyIndex Then
                ' We don't want to initiate attributes that aren't for us.
                If .Moving = 0 Then
                    If Map(Player(MyIndex).Map).Tile(Player(MyIndex).X, Player(MyIndex).Y).Attribute > 0 Then
                        Call MapAttribute(Map(Player(MyIndex).Map).Tile(Player(MyIndex).X, Player(MyIndex).Y).Attribute)
                    End If
                End If
            End If
            
            If .Moving = DIR_UP Or .Moving = DIR_DOWN Then
                X = .YOffset
            ElseIf .Moving = DIR_LEFT Or .Moving = DIR_RIGHT Then
                X = .XOffset
            End If
            If X < 0 Then X = X * -1
            
            If X < 10 Then
                Select Case .Step
                    Case 2
                        .Step = 3
                    Case 4
                        .Step = 1
                End Select
            End If
            
            Exit Sub
            
            If .Moving = 0 Then
                Select Case .Step
                    Case 2
                        .Step = 3
                    Case 4
                        .Step = 1
                End Select
            End If
        End If
    End With
  
End Sub

Sub ProcessNPCMovement(ByVal Index As Long, ByVal Mapnum As Long)
Dim Speed As Long
Dim X As Long

    If Map(Mapnum).MapNpc(Index).Num = 0 Then Exit Sub

    Speed = Npc(Map(Mapnum).MapNpc(Index).Num).Speed

    With TempNpc(Mapnum).NpcNum(Index)
        If .Moving > 0 Then
        
            Select Case .Moving
                Case DIR_UP
                    .YOffset = .YOffset - Speed
                    If .YOffset <= 0 Then
                        .Moving = 0
                        .YOffset = 0
                    End If
                Case DIR_DOWN
                    .YOffset = .YOffset + Speed
                    If .YOffset >= 0 Then
                        .Moving = 0
                        .YOffset = 0
                    End If
                Case DIR_RIGHT
                    .XOffset = .XOffset + Speed
                    If .XOffset >= 0 Then
                        .Moving = 0
                        .XOffset = 0
                    End If
                Case DIR_LEFT
                    .XOffset = .XOffset - Speed
                    If .XOffset <= 0 Then
                        .Moving = 0
                        .XOffset = 0
                    End If
            End Select
            
            If .Moving = DIR_UP Or .Moving = DIR_DOWN Then
                X = .YOffset
            ElseIf .Moving = DIR_LEFT Or .Moving = DIR_RIGHT Then
                X = .XOffset
            End If
            If X < 0 Then X = X * -1
            
            If X < 10 Then
                Select Case .Step
                    Case 2
                        .Step = 3
                    Case 4
                        .Step = 1
                End Select
            End If
            
            Exit Sub
            
            If .Moving = 0 Then
                Select Case .Step
                    Case 2
                        .Step = 3
                    Case 4
                        .Step = 1
                End Select
            End If
        End If
    End With
  
End Sub

Public Sub SetOffset(ByVal Index As Long, ByVal Dir As Byte)

    With TempPlayer(Index)
        Select Case Dir
            Case DIR_DOWN
                .YOffset = -32
            Case DIR_UP
                .YOffset = 32
            Case DIR_RIGHT
                .XOffset = -32
            Case DIR_LEFT
                .XOffset = 32
        End Select
        
        If .Moving > 0 Then
            If .Step = 0 Then .Step = 1
            Select Case .Step
                Case 1
                    .Step = 2
                    Exit Sub
                Case 3
                    .Step = 4
                    Exit Sub
            End Select
        End If
    End With
End Sub

Public Sub SetNPCOffset(ByVal Index As Long, ByVal Dir As Byte, ByVal Mapnum As Long)

    With TempNpc(Mapnum).NpcNum(Index)
        Select Case Map(Mapnum).MapNpc(Index).Dir
            Case DIR_DOWN
                .YOffset = -32
            Case DIR_UP
                .YOffset = 32
            Case DIR_RIGHT
                .XOffset = -32
            Case DIR_LEFT
                .XOffset = 32
        End Select
        
        If .Moving > 0 Then
            If .Step = 0 Then .Step = 1
            Select Case .Step
                Case 1
                    .Step = 2
                Case 3
                    .Step = 4
            End Select
        End If
    End With

End Sub

Public Sub InitiateMovement()

    If CheckTile(MyIndex, Player(MyIndex).Dir) Then
        With Player(MyIndex)
            ' If we can move on the tile, update the quards
            Select Case .Dir
                Case DIR_UP
                    .Y = .Y - 1
                Case DIR_DOWN
                    .Y = .Y + 1
                Case DIR_LEFT
                    .X = .X - 1
                Case DIR_RIGHT
                    .X = .X + 1
            End Select
            Call SetOffset(MyIndex, .Dir)
        End With
    Else
        TempPlayer(MyIndex).Moving = 0
    End If
End Sub

Public Sub InitiateNPCMovement(ByVal Index As Long, ByVal CheckDir As Byte, ByVal Mapnum As Long)

    With TempNpc(Mapnum).NpcNum(Index)
        If CheckNPCTile(Index, CheckDir, Mapnum) Then
            Map(Mapnum).MapNpc(Index).Dir = CheckDir
            ' If we can move on the tile, update the quards
            Select Case Map(Mapnum).MapNpc(Index).Dir
                Case DIR_UP
                    .Y = .Y - 1
                Case DIR_DOWN
                    .Y = .Y + 1
                Case DIR_LEFT
                    .X = .X - 1
                Case DIR_RIGHT
                    .X = .X + 1
            End Select
            Call SetNPCOffset(Index, Map(Mapnum).MapNpc(Index).Dir, Mapnum)
        Else
            .Moving = 0
        End If
    End With
End Sub

Public Sub MapAttribute(ByVal Attr As Long)
Dim Mapnum As Long, NewMapNum As Long
Dim X As Long, Y As Long
Dim NewX As Long, NewY As Long

Mapnum = Player(MyIndex).Map
X = Player(MyIndex).X
Y = Player(MyIndex).Y

    Select Case Attr
        Case Attributes.warptile
            NewMapNum = Map(Mapnum).Tile(X, Y).LongValue(1)
            If NewMapNum < 0 Or NewMapNum > MAX_MAPS Then NewMapNum = Mapnum
            NewX = Map(Mapnum).Tile(X, Y).LongValue(2)
            If NewX < 1 Or NewX > MAX_MAP_X Then NewX = X
            NewY = Map(Mapnum).Tile(X, Y).LongValue(3)
            If NewY < 1 Or NewY > MAX_MAP_Y Then NewY = Y
            Call WarpPlayer(MyIndex, NewMapNum, NewX, NewY)
    End Select
End Sub

Public Sub WarpPlayer(ByVal Index As Long, ByVal Mapnum As Long, ByVal X As Long, ByVal Y As Long)
Dim TileX As Long, TileY As Long
Dim TempMap As Long
Dim i As Long

    ' Is it a dungeon map?
    If Mapnum = 0 Then
        Do While Mapnum = 0
            TempMap = RAND(MIN_DUNGEON_MAP, MAX_DUNGEON_MAP)
            If TempMap <> Player(MyIndex).Map Then
                Mapnum = TempMap
                For TileX = 1 To MAX_MAP_X
                    For TileY = 1 To MAX_MAP_Y
                        With Map(Player(MyIndex).Map)
                            For i = 1 To MAX_MAP_NPCS
                                If .MapNpc(i).Num > 0 Then Call RespawnNpc(Player(MyIndex).Map, i)
                            Next
                            If .Tile(TileX, TileY).Attribute = ResourceTile Then
                                Call RespawnResource(Player(MyIndex).Map, TileX, TileY)
                            End If
                        End With
                    Next
                Next
            End If
        Loop
    End If

    If LoadedMap(Mapnum) = False Then Call LoadMap(Mapnum)
    Player(Index).Map = Mapnum
    Player(Index).X = X
    Player(Index).Y = Y
    TempPlayer(Index).Target = 0
    
    For TileX = 1 To MAX_MAP_X
        For TileY = 1 To MAX_MAP_Y
            TempPlayer(Index).UnlockedTile(TileX, TileY) = False
            OpenChest(TileX, TileY) = False
        Next
    Next

End Sub

Public Function FindOpenInvSlot(ByVal Itemnum As Long) As Byte
Dim i As Long

    If Itemnum > 0 Then
        If Item(Itemnum).Stackable = True Then
            For i = 1 To MAX_INV
                If Player(MyIndex).Inv(i).Num = Itemnum Then
                    FindOpenInvSlot = i
                    Exit Function
                End If
            Next
        End If
    End If

    For i = 1 To MAX_INV
        If Player(MyIndex).Inv(i).Num = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
    
    Call AddText("Your inventory is full.", BrightRed)
    
End Function

Public Function FindOpenLayerSlot(ByVal X As Long, ByVal Y As Long) As Byte
Dim i As Long

    For i = 1 To MAX_MAP_ITEM_LAYERS
        If MapItem(Player(MyIndex).Map).Tile(X, Y).Layer(i).Num = 0 Then
            FindOpenLayerSlot = i
            Exit Function
        End If
    Next
    
    FindOpenLayerSlot = 0
    
End Function

Public Sub PickUpMapItem()
Dim Mapnum As Long
Dim X As Long, Y As Long, i As Long
Dim Slot As Long

    Mapnum = Player(MyIndex).Map
    X = Player(MyIndex).X
    Y = Player(MyIndex).Y
    
        For i = 0 To MAX_MAP_ITEM_LAYERS
            If MapItem(Mapnum).Tile(X, Y).Layer(i).Num > 0 Then
                Slot = FindOpenInvSlot(MapItem(Mapnum).Tile(X, Y).Layer(i).Num)
                If Slot > 0 Then
                    Player(MyIndex).Inv(Slot).Num = MapItem(Mapnum).Tile(X, Y).Layer(i).Num
                    Player(MyIndex).Inv(Slot).Value = Player(MyIndex).Inv(Slot).Value + MapItem(Mapnum).Tile(X, Y).Layer(i).Value
                    MapItem(Mapnum).Tile(X, Y).Layer(i).Num = 0
                    MapItem(Mapnum).Tile(X, Y).Layer(i).Value = 0
                    Call BltInventory
                    Exit Sub
                End If
            End If
        Next
    
End Sub

Public Sub GivePlayerItem(ByVal Itemnum As Long, Optional Value As Long = 0)
Dim Slot As Long

    Slot = FindOpenInvSlot(Itemnum)
    If Slot > 0 Then
        Player(MyIndex).Inv(Slot).Num = Itemnum
        Player(MyIndex).Inv(Slot).Value = Player(MyIndex).Inv(Slot).Value + Value
    End If
    
End Sub

Public Sub DropItem(ByVal Slot As Long)
Dim Mapnum As Long
Dim Layer As Long
Dim X As Long
Dim Y As Long
    
    X = Player(MyIndex).X
    Y = Player(MyIndex).Y
    Mapnum = Player(MyIndex).Map

    Layer = FindOpenLayerSlot(X, Y)
    If Layer = 0 Then
        AddText "This tile seems to be piled with items...", BrightRed
        Exit Sub
    End If

    MapItem(Mapnum).Tile(X, Y).Layer(Layer).Num = Player(MyIndex).Inv(Slot).Num
    MapItem(Mapnum).Tile(X, Y).Layer(Layer).Value = Player(MyIndex).Inv(Slot).Value
    MapItem(Mapnum).Tile(X, Y).Layer(Layer).Tick = Default_Map_Item_Appear
    Player(MyIndex).Inv(Slot).Num = 0
    Player(MyIndex).Inv(Slot).Value = 0
    Call BltInventory
    
End Sub

Public Sub UseItem(ByVal Slot As Long)
Dim Itemnum As Long
Dim TakeAwayItem As Boolean
Dim i As Long

    Itemnum = Player(MyIndex).Inv(Slot).Num
    
    If Options.OnlineMode = True Then
        'send the packet
        Exit Sub
    End If
    
    If Game.InBank = True Then
        Call InsertBankItem(Slot)
        Exit Sub
    End If
    
    If Game.InShop = True Then
        Exit Sub
    End If
    
    For i = 1 To Stats.Stat_Count - 1
        If Item(Itemnum).StatReq(i) > Player(MyIndex).Stat(i) Then
            Call AddText("You don't have the stat requirements to use this item.", BrightRed)
            Exit Sub
        End If
    Next
    
    With Player(MyIndex)
        If Item(Itemnum).Type > ITEM_TYPE_NONE And Item(Itemnum).Type <= ITEM_TYPE_BOOTS Then
            If Item(Itemnum).Type <> ITEM_TYPE_NULL1 Or Item(Itemnum).Type <> ITEM_TYPE_NULL2 Then
                Call EquipItem(Slot, Item(Itemnum).Type)
            End If
        End If
        Select Case Item(Itemnum).Type
            Case ITEM_TYPE_CONSUME
                If Item(Itemnum).Spell > 0 Then
                    If LearnSpell(MyIndex, Item(Itemnum).Spell) Then
                        TakeAwayItem = True
                        If Item(Itemnum).addHP > 0 Then .Vital(Vitals.Health) = .Vital(Vitals.Health) + Item(Itemnum).addHP
                        If Item(Itemnum).addSP > 0 Then .Vital(Vitals.Spirit) = .Vital(Vitals.Spirit) + Item(Itemnum).addSP
                    End If
                Else
                    If Item(Itemnum).addHP > 0 Then .Vital(Vitals.Health) = .Vital(Vitals.Health) + Item(Itemnum).addHP
                    If Item(Itemnum).addSP > 0 Then .Vital(Vitals.Spirit) = .Vital(Vitals.Spirit) + Item(Itemnum).addSP
                End If
        End Select
    End With
    
    If Item(Itemnum).CustomScript > 0 Then
        Call ItemCustomScript(Item(Itemnum).CustomScript)
    End If
    
    If TakeAwayItem = True Then
        Player(MyIndex).Inv(Slot).Value = Player(MyIndex).Inv(Slot).Value - 1
        If Player(MyIndex).Inv(Slot).Value <= 0 Then Player(MyIndex).Inv(Slot).Num = 0
    End If
    
    If Item(Itemnum).GiveBack > 0 Then
        Call GivePlayerItem(Item(Itemnum).GiveBack)
    End If
    
    Call UpdatePlayerVitals(MyIndex)
    Call BltInventory
    Call BltCharacterScreen
End Sub

Public Sub EquipItem(ByVal Slot As Long, ByVal EqSlot As Long)
Dim OldEqItem As Long, OldEqValue As Long
Dim Itemnum As Long

    Itemnum = Player(MyIndex).Inv(Slot).Num
    
    With Item(Player(MyIndex).Inv(Slot).Num)
        Select Case .Type
            Case ITEM_TYPE_WEAPON
                If .IsTwoHanded = True Then
                    If Player(MyIndex).Equipment(Equipment.Shield).Num > 0 Then
                        If FindOpenInvSlot(Player(MyIndex).Inv(Slot).Num) > 0 Then
                            Call UnequipItem(Equipment.Shield)
                        End If
                    End If
                End If
                        
            Case ITEM_TYPE_SHIELD
                If Player(MyIndex).Equipment(Equipment.Weapon).Num > 0 Then
                    If Item(Player(MyIndex).Equipment(Equipment.Weapon).Num).IsTwoHanded = True Then
                        Dim UnEquipWeapon As Boolean
                        UnEquipWeapon = True
                    End If
                End If
        End Select
    End With
    
    With Player(MyIndex)
        OldEqItem = .Equipment(EqSlot).Num
        OldEqValue = .Equipment(EqSlot).Value
        
        If Item(Itemnum).Stackable = True Then
            .Equipment(EqSlot).Num = .Inv(Slot).Num
            .Equipment(EqSlot).Value = .Equipment(EqSlot).Value + .Inv(Slot).Value
            .Inv(Slot).Num = 0
            .Inv(Slot).Value = 0
        Else
            .Equipment(EqSlot).Num = .Inv(Slot).Num
            .Equipment(EqSlot).Value = .Inv(Slot).Value
            .Inv(Slot).Num = OldEqItem
            .Inv(Slot).Value = OldEqValue
        End If
    End With
    
    Call UpdatePlayerStance(MyIndex)
    
    If UnEquipWeapon = True Then Call UnequipItem(Equipment.Weapon)
    
End Sub

Public Sub UnequipItem(ByVal EqSlot As Long)
Dim Slot As Long

    With Player(MyIndex)
        If .Equipment(EqSlot).Num > 0 Then
            Slot = FindOpenInvSlot(.Equipment(EqSlot).Num)
            If Slot > 0 Then
                .Inv(Slot).Num = .Equipment(EqSlot).Num
                .Inv(Slot).Value = .Inv(Slot).Value + .Equipment(EqSlot).Value
                .Equipment(EqSlot).Num = 0
                .Equipment(EqSlot).Value = 0
                Call BltCharacterScreen
                Call BltInventory
                Call UpdatePlayerStance(MyIndex)
            End If
        End If
    End With
End Sub

Public Sub LevelUpPlayer(ByVal Index As Long)
    
    If Player(Index).Combat.level >= MAX_COMBAT_LEVEL Then Exit Sub
    
    Player(MyIndex).Combat.level = Player(MyIndex).Combat.level + 1
    Call AddText(Trim$(Player(MyIndex).Name) & " has leveled up!", Pink)
    
End Sub

Public Sub TrainStat(ByVal Index As Long, ByVal StatIndex As Long)
    If Player(Index).Points = 0 Or Player(Index).Stat(StatIndex) = 100 Then Exit Sub
    Player(Index).Stat(StatIndex) = Player(Index).Stat(StatIndex) + 1
    Player(Index).Points = Player(Index).Points - 1
    Call frmMain.TabCharacterInit
    Call UpdatePlayerVitals(MyIndex)
End Sub

Public Function GetPlayerNextLevelXP(ByVal level As Long)
    GetPlayerNextLevelXP = (level * 4756)
End Function

Public Sub ItemCustomScript(ByVal Script As Long)

End Sub

Public Sub UpdatePlayerStance(ByVal Index As Long)
    
    With Player(Index)
        Select Case .Graphics.Gender
            Case GENDER_MALE
                Call SetPlayerStance(Index, Stance.MNorm)
                If .Equipment(Equipment.Weapon).Num > 0 Then
                    If Item(.Equipment(Equipment.Weapon).Num).IsTwoHanded = True Then
                        Call SetPlayerStance(Index, Stance.MTwoHand)
                    End If
                End If
                If .Equipment(Equipment.Shield).Num > 0 Then
                    Call SetPlayerStance(Index, Stance.MShield)
                End If
            Case GENDER_FEMALE
                Call SetPlayerStance(Index, Stance.FNorm)
                If .Equipment(Equipment.Weapon).Num > 0 Then
                    If Item(.Equipment(Equipment.Weapon).Num).IsTwoHanded = True Then
                        Call SetPlayerStance(Index, Stance.FTwoHand)
                    End If
                End If
                If .Equipment(Equipment.Shield).Num > 0 Then
                    Call SetPlayerStance(Index, Stance.FShield)
                End If
        End Select
    End With
End Sub

Public Sub SetPlayerStance(ByVal Index As Long, ByVal Stan As Integer)

    With Player(Index).Graphics
        Player(Index).Stance = Stan
        ' Skin
        Select Case .Skin
            Case 1, 5, 9
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Skin = 1
                    Case Stance.MShield, Stance.FShield
                        .Skin = 5
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Skin = 9
                End Select
            Case 2, 6, 10
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Skin = 2
                    Case Stance.MShield, Stance.FShield
                        .Skin = 6
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Skin = 10
                End Select
            Case 3, 7, 11
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Skin = 3
                    Case Stance.MShield, Stance.FShield
                        .Skin = 7
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Skin = 11
                End Select
            Case 4, 8, 12
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Skin = 4
                    Case Stance.MShield, Stance.FShield
                        .Skin = 8
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Skin = 12
                End Select
        End Select
        ' hair
        Select Case .Hair
            Case 1, 5, 9
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Hair = 1
                    Case Stance.MShield, Stance.FShield
                        .Hair = 5
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Hair = 9
                End Select
            Case 2, 6, 10
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Hair = 2
                    Case Stance.MShield, Stance.FShield
                        .Hair = 6
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Hair = 10
                End Select
            Case 3, 7, 11
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Hair = 3
                    Case Stance.MShield, Stance.FShield
                        .Hair = 7
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Hair = 11
                End Select
            Case 4, 8, 12
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Hair = 4
                    Case Stance.MShield, Stance.FShield
                        .Hair = 8
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Hair = 12
                End Select
        End Select
        ' body
        Select Case .Body
            Case 1, 5, 9
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Body = 1
                    Case Stance.MShield, Stance.FShield
                        .Body = 5
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Body = 9
                End Select
            Case 2, 6, 10
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Body = 2
                    Case Stance.MShield, Stance.FShield
                        .Body = 6
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Body = 10
                End Select
            Case 3, 7, 11
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Body = 3
                    Case Stance.MShield, Stance.FShield
                        .Body = 7
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Body = 11
                End Select
            Case 4, 8, 12
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Body = 4
                    Case Stance.MShield, Stance.FShield
                        .Body = 8
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Body = 12
                End Select
        End Select
        ' legs
        Select Case .Legs
            Case 1, 5, 9
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Legs = 1
                    Case Stance.MShield, Stance.FShield
                        .Legs = 5
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Legs = 9
                End Select
            Case 2, 6, 10
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Legs = 2
                    Case Stance.MShield, Stance.FShield
                        .Legs = 6
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Legs = 10
                End Select
            Case 3, 7, 11
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Legs = 3
                    Case Stance.MShield, Stance.FShield
                        .Legs = 7
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Legs = 11
                End Select
            Case 4, 8, 12
                Select Case Stan
                    Case Stance.MNorm, Stance.FNorm
                        .Legs = 4
                    Case Stance.MShield, Stance.FShield
                        .Legs = 8
                    Case Stance.MTwoHand, Stance.MTwoHand
                        .Legs = 12
                End Select
        End Select
    End With

End Sub

Public Sub TakeInvItem(ByVal Index As Long, ByVal invSlot As Long, Optional ByVal Value As Long)

    With Player(Index).Inv(invSlot)
        If Value > 0 Then
            If .Value - Value <= 0 Then
                .Value = 0
                .Num = 0
            Else
                .Value = .Value - Value
            End If
        Else
            .Value = 0
            .Num = 0
        End If
    End With
    Call BltInventory
End Sub

Public Sub WithdrawBankItem(ByVal BankSlot As Long)
Dim Slot As Long, Itemnum As Long
Dim TempValue As Long
Dim i As Long

    TempValue = Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot).Value
    If WIMultiplier < TempValue Then TempValue = WIMultiplier
    Itemnum = Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot).Num
    
    If Item(Itemnum).Stackable = True Then
        If FindOpenInvSlot(Itemnum) > 0 Then
            Call GivePlayerItem(Itemnum, TempValue)
            With Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot)
                .Value = .Value - TempValue
                If .Value <= 0 Then .Num = 0
            End With
        End If
    Else
        For i = 1 To TempValue
            If FindOpenInvSlot(Itemnum) > 0 Then
                Call GivePlayerItem(Itemnum, 1)
                With Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot)
                    .Value = .Value - 1
                    If .Value <= 0 Then
                        .Num = 0
                        .Value = 0
                    End If
                End With
            Else
                Call BltInventory
                Call RenderBank
                Exit For
            End If
        Next
    End If
    
    Call RenderBank
    Call BltInventory
    
End Sub

Public Sub InsertBankItem(ByVal invSlot As Long)
Dim BankSlot As Long, Itemnum As Long
Dim TempValue As Long, TempMulti As Long
Dim i As Long, Slot As Long

    Itemnum = Player(MyIndex).Inv(invSlot).Num

    TempValue = Player(MyIndex).Inv(invSlot).Value
    If WIMultiplier < TempValue Then TempValue = WIMultiplier
    
    If Item(Itemnum).Stackable = True Then
    
        If FindBankSlot(Itemnum) > 0 Then
            BankSlot = FindBankSlot(Itemnum)
            With Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot)
                .Num = Itemnum
                .Value = .Value + TempValue
            End With
            With Player(MyIndex).Inv(invSlot)
                .Value = .Value - TempValue
                If .Value <= 0 Then .Num = 0
            End With
        End If
    Else
        TempMulti = WIMultiplier
        If TempMulti > 35 Then TempMulti = 35
        TempValue = TempMulti
        For i = 1 To TempMulti
            For Slot = 1 To MAX_INV
                With Player(MyIndex).Inv(Slot)
                    If .Num = Itemnum Then
                        If FindBankSlot(Itemnum) > 0 Then
                            BankSlot = FindBankSlot(Itemnum)
                            .Num = 0
                            .Value = 0
                            With Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot)
                                .Num = Itemnum
                                .Value = .Value + 1
                            End With
                            TempValue = TempValue - 1
                            If TempValue <= 0 Then
                                Call BltInventory
                                Call RenderBank
                                Exit Sub
                            End If
                        End If
                    End If
                End With
            Next
        Next
    End If
    Call BltInventory
    Call RenderBank
    
End Sub

Public Function FindBankSlot(ByVal Itemnum As Long) As Long
Dim i As Long

    For i = 1 To MAX_BANK_ITEMS
        If Bank(MyIndex).BankTab(CurTab).BankItem(i).Num = Itemnum Then
            FindBankSlot = i
            Exit Function
        End If
    Next
    
    For i = 1 To MAX_BANK_ITEMS
        If Bank(MyIndex).BankTab(CurTab).BankItem(i).Num = 0 Then
            FindBankSlot = i
            Exit Function
        End If
    Next
    
    Call AddText("This bank page is full!", BrightRed)
End Function

Public Function TransformAmount(ByVal Amount As Long) As String

    If Int(Amount) < 10000 Then
        TransformAmount = Amount
    ElseIf Int(Amount) < 999999 Then
        TransformAmount = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        TransformAmount = Int(Amount / 1000000) & "m"
    Else
        TransformAmount = Int(Amount / 1000000000) & "b"
    End If
    
End Function

Public Sub BuyItem(ByVal ShopSlot As Byte)
Dim ShopNum As Long
Dim i As Byte, X As Byte, Y As Long
Dim TakeAwaySlot(1 To 10) As Long, TakeAwayValue(1 To 10) As Long
Dim BoughtItem(1 To 10) As Boolean

    ShopNum = Game.ShopNum
    If Shop(ShopNum).ShopItem(ShopSlot).StockItem > 0 Then
        With Shop(ShopNum).ShopItem(ShopSlot)
            If .NumberofCosts = 0 Then
                Call GivePlayerItem(.StockItem, 1)
                Call BltInventory
            Else
                If .AddXP = True Then
                    If Not CanBuyItem(.StockItem) Then
                        Exit Sub
                    End If
                End If
                For i = 1 To .NumberofCosts
                    For X = 1 To MAX_INV
                        If .ItemCost(i).ItemCostNum = Player(MyIndex).Inv(X).Num Then
                            If .ItemCost(i).ItemCostValue <= Player(MyIndex).Inv(X).Value Then
                                For Y = 1 To .NumberofCosts
                                    If X <> TakeAwaySlot(Y) Then
                                        BoughtItem(i) = True
                                        If .ItemCost(i).UseUpItem = True Then
                                            TakeAwaySlot(i) = X
                                            TakeAwayValue(i) = .ItemCost(i).ItemCostValue
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        If BoughtItem(.NumberofCosts) = True Then Exit For
                    Next
                Next
                
                For i = 1 To .NumberofCosts
                    If BoughtItem(i) = False Then
                        Call AddText("You couldn't buy the item.", BrightRed)
                        Exit Sub
                    End If
                Next
                
                For Y = 1 To .NumberofCosts
                    If TakeAwaySlot(Y) > 0 Then
                        Call TakeInvItem(MyIndex, TakeAwaySlot(Y), TakeAwayValue(Y))
                        If .AddXP = True Then
                            Call GiveShopItemXP(.StockItem)
                        End If
                    End If
                Next
                Call GivePlayerItem(.StockItem, .StockValue)
                Call BltInventory
                
            End If
        End With
    End If
End Sub

Public Sub PopulateChest(ByVal Mapnum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Byte
Dim ChestNum

    If Map(Mapnum).Tile(X, Y).Attribute = Attributes.ChestTile Then
        ChestNum = Map(Mapnum).Tile(X, Y).LongValue(1)
    Else
        Exit Sub
    End If
    
    For i = 1 To MAX_CHEST_ITEMS
        With MapChest(Mapnum).Tile(X, Y).ChestItem(i)
            If RAND(1, 100) <= Chest(ChestNum).ChestItem(i).Chance Then
                .Itemnum = Chest(ChestNum).ChestItem(i).Itemnum
                .ItemValue = Chest(ChestNum).ChestItem(i).ItemValue
            Else
                .Itemnum = 0
                .ItemValue = 0
            End If
        End With
    Next
    
    OpenChest(X, Y) = True


End Sub

Public Function CanBuyItem(ByVal Itemnum As Long) As Boolean
Dim i As Byte

    CanBuyItem = True

    With Player(MyIndex)
        For i = 1 To Skills.Skill_Count - 1
            If .Skill(i).level < Item(Itemnum).ReqXP(i) Then
                Call AddText("You need a level of " & Item(Itemnum).ReqXP(i) & " in " & GetSkillName(i) & " to make this item.", BrightRed)
                CanBuyItem = False
            End If
        Next
    End With

End Function

Public Sub GiveShopItemXP(ByVal Itemnum As Long)
Dim i As Byte

    For i = 1 To Skills.Skill_Count - 1
        Call GiveSkillXp(i, Item(Itemnum).RewXP(i))
    Next
    
End Sub

Public Sub GiveSkillXp(ByVal SkillNum As Byte, ByVal GiveAmount As Long)

    With Player(MyIndex).Skill(SkillNum)
        .XP = .XP + GiveAmount
        If .XP >= GetPlayerNextLevelXP(.level) Then
            .XP = (GetPlayerNextLevelXP(.level) - .XP) * (-1)
            .level = .level + 1
            If frmMain.picSkills.Visible = True Then frmMain.TabSkillsInit
            Call AddText("You advanced a level in " & GetSkillName(SkillNum), White)
        End If
    End With

End Sub

Public Function GetSkillName(ByVal Num As Byte) As String

    Select Case Num
        Case Skills.Woodcutting
            GetSkillName = "Woodcutting"
        Case Skills.Mining
            GetSkillName = "Mining"
        Case Skills.Cooking
            GetSkillName = "Cooking"
        Case Skills.Crafting
            GetSkillName = "Crafting"
        Case Skills.Fishing
            GetSkillName = "Fishing"
        Case Skills.PotionBrewing
            GetSkillName = "Potion Brewing"
        Case Skills.Smithing
            GetSkillName = "Smithing"
        Case Skills.Fletching
            GetSkillName = "Fletching"
    End Select
End Function

Public Sub RespawnResource(ByVal Mapnum As Long, ByVal X As Long, ByVal Y As Long)

    With MapResource(Mapnum).Tile(X, Y)
        .Alive = True
        .Health = Resource(.Num).Health
    End With

End Sub

Public Sub ActionMsg(ByVal message As String, ByVal X As Long, ByVal Y As Long, ByVal color As Integer)
Dim i As Long
Dim ActionmsgIndex As Long

    For i = 1 To 255
        If TempActionMsg(i).Created = 0 Then
            ActionmsgIndex = i
            Exit For
        End If
    Next

    With TempActionMsg(ActionmsgIndex)
        .message = message
        .color = color
        .Created = timeGetTime
        .Scroll = 1
        .X = X * 32
        .Y = Y * 32
    End With
    
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long
    
    TempActionMsg(Index).message = vbNullString
    TempActionMsg(Index).Created = 0
    TempActionMsg(Index).color = 0
    TempActionMsg(Index).Scroll = 0
    TempActionMsg(Index).X = 0
    TempActionMsg(Index).Y = 0

End Sub

Public Function LearnSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long, Slot As Long

    LearnSpell = False
    
    For i = MAX_PLAYER_SPELLS To 1 Step -1
        With Player(Index).PlayerSpell(i)
            If .Num = SpellNum Then
                LearnSpell = False
                Call AddText("You already know this spell!", BrightRed)
                Exit Function
            End If
            If .Num = 0 Then
                Slot = i
            End If
        End With
    Next
    
    Player(Index).PlayerSpell(Slot).Num = SpellNum
    Player(Index).PlayerSpell(Slot).CoolDownTimer = 0
    Call AddText("You now know the glorious power of " & Trim$(Spell(SpellNum).Name), Yellow)
    Call BltSpells
    
    LearnSpell = True

End Function

Public Function ForgetSpell(ByVal SpellNum As Long)

    Player(MyIndex).PlayerSpell(SpellNum).Num = 0
    Player(MyIndex).PlayerSpell(SpellNum).CoolDownTimer = 0
    Call BltSpells
End Function

Public Sub UpdateProjectileLogic(ByVal Mapnum As Long, ByVal Num As Long)
Dim i As Long
Dim PSquare As Long
Dim ClearProjectile As Boolean
Dim Speed As Double

    With MapProjectile(Mapnum).MapProjectile(Num)
        PSquare = .Distance
        Speed = .Speed
        Select Case .Dir
            Case DIR_DOWN
                .YOffset = .YOffset - Speed
                If .YOffset <= -32 Then
                    .YOffset = 0
                    .Y = .Y + 1
                    .Distance = .Distance + 1
                End If
            Case DIR_UP
                .YOffset = .YOffset + Speed
                If .YOffset >= 32 Then
                    .YOffset = 0
                    .Y = .Y - 1
                    .Distance = .Distance + 1
                End If
            Case DIR_RIGHT
                .XOffset = .XOffset - Speed
                If .XOffset <= -32 Then
                    .XOffset = 0
                    .X = .X + 1
                    .Distance = .Distance + 1
                End If
            Case DIR_LEFT
                .XOffset = .XOffset + Speed
                If .XOffset >= 32 Then
                    .XOffset = 0
                    .X = .X - 1
                    .Distance = .Distance + 1
                End If
        End Select
        
        If .Distance >= .Range Or .X > MAX_MAP_X Or .X < 1 Or .Y < 1 Or .Y > MAX_MAP_Y Then
            Call ClearMapProjectile(Mapnum, Num)
            Exit Sub
        End If
        
        ' We moved a square
        If PSquare <> .Distance Then
            ' See what we can do.
            If CheckProjHit(Mapnum, Num) Then
                ClearProjectile = True
            End If
            
            ' Can we travel to the next square?
            If CheckProjecTile(.X, .Y, .Dir) Then
                ClearProjectile = True
            End If
        End If
        
        If ClearProjectile Then
            Call ClearMapProjectile(Mapnum, Num)
        End If
    End With

End Sub

Public Sub ClearMapProjectile(ByVal Mapnum As Long, ByVal Num As Long)
    With MapProjectile(Mapnum).MapProjectile(Num)
        .Dir = 0
        .Picture = 0
        .Picture = 0
        .Range = 0
        .Speed = 0
        .X = 0
        .XOffset = 0
        .Y = 0
        .YOffset = 0
        .Distance = 0
        .Range = 0
    End With
End Sub

Public Sub CreateProjectile(ByVal Mapnum As Long, ByVal Image As Long, ByVal Speed As Long, ByVal Range As Long, ByVal Dir As Byte, ByVal X As Long, ByVal Y As Long)
Dim i As Long
    
    For i = 1 To MAX_MAP_PROJECTILES
        With MapProjectile(Mapnum).MapProjectile(i)
            If .Range = 0 Then
                .Dir = Dir
                .Speed = Speed
                .Picture = Image
                .Range = Range
                .X = X
                .Y = Y
                Exit For
            End If
        End With
    Next

End Sub
