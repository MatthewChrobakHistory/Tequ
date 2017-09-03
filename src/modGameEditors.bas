Attribute VB_Name = "modGameEditors"
Option Explicit

' Map Editor
Public Const EDITING_LAYERS As Byte = 1
Public Const EDITING_ATTRIBUTES As Byte = 2
Public Const EDITING_MAP_ITEMS As Byte = 3
Public CurrentlyEditing As Byte
Public EditingLayer As Byte
Public EditingAttribute As Byte
Public CurTileX As Long
Public CurTileY As Long
Public EditingPage As Byte

Public EItemNum As Long
Public ENpcNum As Long
Public EResourceNum As Long
Public EShopNum As Long
Public EChestNum As Long
Public ESpellNum As Long
Public EInterfaceNum As Long

Public Sub NewResourceIndex(ByVal Index As Long)
Dim i As Byte

    If Index = 0 Then
        EResourceNum = 1
        Index = 1
    End If
    
    With frmEditor_Resource
        .txtName.text = Trim$(Resource(Index).Name)
        .txtRespawn.text = Resource(Index).RespawnRate
        .txtHealth.text = Resource(Index).Health
        .scrlAlive.Value = Resource(Index).AliveGFX
        .scrlDead.Value = Resource(Index).DeadGFX
        .scrlItem.Value = Resource(Index).Reward
        .scrlRewardValue.Value = Resource(Index).RewardValue
        .cmbEquipmentType.ListIndex = Resource(Index).EquipmentType
        
        For i = 1 To Skills.Skill_Count - 1
            .txtSkillReq(i).text = Resource(Index).RequiredXP(i)
            .txtSkillReward(i).text = Resource(Index).RewardXP(i)
        Next
        
    End With
End Sub

Public Sub NewNpcIndex(ByVal Index As Long)
Dim i As Long

    If Index = 0 Then
        EItemNum = 1
        Index = 1
    End If
    
    With frmEditor_Npc
        .txtName.text = Trim$(Npc(Index).Name)
        .txtRespawn.text = Npc(Index).Respawn
        .scrlSprite = Npc(Index).Sprite
        .cmbType.ListIndex = Npc(Index).Type
        If Npc(Index).Speed = 0 Then Npc(Index).Speed = 1
        .scrlSpeed.Value = Npc(Index).Speed * 10
        If Npc(Index).AttackSpeed = 0 Then Npc(Index).AttackSpeed = 1
        .scrlAttackSpeed.Value = Npc(Index).AttackSpeed * 10
        .txtRange.text = Npc(Index).Range
        If Npc(Index).Vital(Vitals.Health) > 0 Then
            .txtHealth.text = Npc(Index).Vital(Vitals.Health)
        Else
            .txtHealth.text = "1"
        End If
        If Npc(Index).Vital(Vitals.Spirit) > 0 Then
            .txtSpirit.text = Npc(Index).Vital(Vitals.Spirit)
        Else
            .txtSpirit.text = "1"
        End If
        For i = 1 To Combat.Combat_Count - 1
            .txtDefense(i).text = Npc(ENpcNum).Defense(i)
            .txtOffense(i).text = Npc(ENpcNum).Offense(i)
        Next
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Item(Index).Stat(i)
        Next
        .txtXP.text = Npc(ENpcNum).XP
        .cmbAttackType.ListIndex = Npc(ENpcNum).AttackType
        .txtDamage.text = Npc(ENpcNum).Damage
        .scrlDrop.Value = 1
    End With
End Sub

Public Sub NewChestIndex(ByVal Index As Long)

    If Index = 0 Then
        EChestNum = 1
        Index = 1
    End If
    
    With frmEditor_Chest
        .txtName.text = Trim$(Chest(Index).Name)
        If Chest(Index).Picture = 0 Then Chest(Index).Picture = 1
        .scrlPicture.Value = Chest(Index).Picture
        .scrlChestItem.Value = 1
        .txtChance.text = Chest(Index).ChestItem(1).Chance
    End With
    
End Sub

Public Sub NewSpellIndex(ByVal Index As Long)
Dim i As Long

    If Index = 0 Then
        ESpellNum = 1
        Index = 1
    End If
    
    With frmEditor_Spell
        .txtName.text = Trim$(Spell(Index).Name)
        .scrlPicture.Value = Spell(Index).Picture
        .scrlRange.Value = Spell(Index).Range
        .cmbType.ListIndex = Spell(Index).Type
        .txtCoolDown.text = Spell(Index).CoolDown
        .txtStunDuration.text = Spell(Index).StunDuration
        .txtY.text = Spell(Index).Y
        .txtX.text = Spell(Index).X
        For i = 1 To 2
            .txtVital(i).text = Spell(Index).VitalAffect(i)
        Next
    End With
End Sub

Public Sub NewItemIndex(ByVal Index As Long)
Dim i As Long
    If Index = 0 Then
        EItemNum = 1
        Index = 1
    End If
    
    With frmEditor_Item
        .txtName.text = Trim$(Item(Index).Name)
        If Item(Index).Picture > 0 Then
            .picItem.Picture = Nothing
            If FileExist(App.Path & "\graphics\items\" & Item(Index).Picture & ".bmp") = True Then
                .picItem.Picture = LoadPicture(App.Path & "\graphics\items\" & Item(Index).Picture & ".bmp")
            End If
        End If
        .scrlPicNum.Value = Item(Index).Picture
        .txtPrice.text = Item(Index).Price
        .lstType.ListIndex = Item(Index).Type
        For i = 1 To Stance.Stance_Count - 1
            .scrlStance(i).Value = Item(Index).Paperdoll(i)
        Next
        If Item(Index).Type = ITEM_TYPE_WEAPON Then
            .chkTwoHanded.Visible = True
            If Item(Index).IsTwoHanded = True Then
                .chkTwoHanded.Value = 1
            Else
                .chkTwoHanded.Value = 0
            End If
        Else
            .chkTwoHanded.Visible = False
        End If
        
        .txtInfo.text = Trim$(Item(Index).info)
        If Item(Index).BltPlayerGraphics = True Then
            .chkBltGraphics.Value = 1
        Else
            .chkBltGraphics.Value = 0
        End If
        
        .scrladdHP.Value = Item(Index).addHP
        .scrladdSP.Value = Item(Index).addSP
        
        .scrlGiveBack.Value = Item(Index).GiveBack
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Item(Index).Stat(i)
            .scrlStatReq(i).Value = Item(Index).StatReq(i)
        Next
        
        For i = 1 To Combat.Combat_Count - 1
            .txtDefense(i).text = Item(Index).Defense(i)
            .txtOffense(i).text = Item(Index).Offense(i)
        Next
        
        For i = 1 To Skills.Skill_Count - 1
            .txtSkillMake(i).text = Item(Index).ReqXP(i)
            .txtSkillReq(i).text = Item(Index).WReqXP(i)
            .txtSkillReward(i).text = Item(Index).RewXP(i)
        Next
        
        .txtDamage.text = Item(Index).Damage
        .txtSpeed.text = Item(Index).Speed
        
        If Item(Index).Stackable = True Then
            .chkStackable.Value = 1
        Else
            .chkStackable.Value = 0
        End If
        
        .scrlLearnSpell.Value = Item(Index).Spell
        
        .scrlPImage.Value = Item(Index).Projectile.Image
        .scrlPRange.Value = Item(Index).Projectile.Range
        .scrlPSpeed.Value = Item(Index).Projectile.Speed * 10
        
        .cmbCombatType.ListIndex = Item(Index).CombatType
        
    End With
    
End Sub

Public Sub NewShopIndex(ByVal Index As Long)
Dim i As Long

    If Index = 0 Then
        EShopNum = 1
        Index = 1
    End If
    
    With frmEditor_Shop
        .txtName.text = Trim$(Shop(Index).Name)
        .lstStock.Clear
        .Caption = .lstStock.ListIndex
        For i = 1 To MAX_SHOP_ITEMS
            If Shop(EShopNum).ShopItem(i).StockItem > 0 Then
                .lstStock.AddItem (i & ": " & Trim$(Item(Shop(EShopNum).ShopItem(i).StockItem).Name))
            Else
                .lstStock.AddItem (i & ": ")
            End If
            .Caption = .lstStock.ListIndex
        Next
        .cmbItemCosts.ListIndex = Shop(EShopNum).ShopItem(1).NumberofCosts
        .txtVerb.text = Trim$(Shop(EShopNum).ShopItem(1).Verb)
        If Shop(EShopNum).ShopItem(1).AddXP = True Then
            .chkXP.Value = 1
        Else
            .chkXP.Value = 0
        End If
        
        .scrlStockValue.Value = Shop(EShopNum).ShopItem(1).StockValue
        .scrlStockValue.Visible = False
        .lblStockValue.Visible = False
        If Shop(EShopNum).ShopItem(1).StockItem > 0 Then
            If Item(Shop(EShopNum).ShopItem(1).StockItem).Stackable = True Then
                .scrlStockValue.Visible = True
                .lblStockValue.Visible = True
            Else
                Shop(EShopNum).ShopItem(1).StockValue = 1
            End If
        End If
        
    End With
End Sub

Public Sub InitShopEditor()
Dim i As Long, X As Long, Y As Long
    
    With frmEditor_Shop
        .lstIndex.Clear
        .lstItems.Clear
        .lstStock.Clear
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem (i & ": " & Trim$(Shop(i).Name))
            For X = 1 To MAX_SHOP_ITEMS
                If Shop(i).ShopItem(X).StockItem > 0 Then
                    .lstStock.AddItem (i & ": " & Trim$(Item(Shop(i).ShopItem(X).StockItem).Name))
                Else
                    .lstStock.AddItem (i & ": ")
                End If
            Next
        Next
        For i = 1 To MAX_ITEMS
            .lstItems.AddItem (i & ": " & Trim$(Item(i).Name))
        Next
        If EShopNum = 0 Then
            .lstIndex.ListIndex = 0
            EShopNum = .lstIndex.ListIndex + 1
        Else
            .lstIndex.ListIndex = EShopNum - 1
        End If
        .lstItems.ListIndex = 0
        .lstStock.ListIndex = 0
        .txtName.text = Trim$(Shop(EShopNum).Name)
        .txtVerb.text = Trim$(Shop(EShopNum).ShopItem(.lstStock.ListIndex + 1).Verb)
        
    .scrlStockValue.Value = Shop(EShopNum).ShopItem(.lstStock.ListIndex + 1).StockValue
    .scrlStockValue.Visible = False
    .lblStockValue.Visible = False
    If Shop(EShopNum).ShopItem(.lstStock.ListIndex + 1).StockItem > 0 Then
        If Item(Shop(EShopNum).ShopItem(.lstStock.ListIndex + 1).StockItem).Stackable = True Then
            .scrlStockValue.Visible = True
            .lblStockValue.Visible = True
        Else
            Shop(EShopNum).ShopItem(.lstStock.ListIndex + 1).StockValue = 1
        End If
    End If
        
        .Show
        .scrlPicture = Shop(EShopNum).Picture
    End With
End Sub

Public Sub InitResourceEditor()
Dim i As Long

    With frmEditor_Resource
        .lstIndex.Clear
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem (i & ": " & Trim$(Resource(i).Name))
        Next
        
        If EResourceNum = 0 Then
            .lstIndex.ListIndex = 0
            EResourceNum = .lstIndex.ListIndex + 1
        Else
            .lstIndex.ListIndex = EResourceNum - 1
        End If
        
        .scrlAlive.max = NumResources
        .scrlDead.max = NumResources
        .txtName.MaxLength = NAME_LENGTH
        .Show
    End With
    
End Sub

Public Sub InitChestEditor()
Dim i As Long

    With frmEditor_Chest
        .lstIndex.Clear
        For i = 1 To MAX_CHESTS
            .lstIndex.AddItem (i & ": " & Trim$(Chest(i).Name))
        Next
        
        If EChestNum = 0 Then
            .lstIndex.ListIndex = 0
            EChestNum = .lstIndex.ListIndex + 1
        Else
            .lstIndex.ListIndex = EItemNum - 1
        End If
        .txtName.MaxLength = NAME_LENGTH
        .scrlPicture.max = NumChests
        .scrlChestItem.max = MAX_CHEST_ITEMS
        .scrlItem.max = MAX_ITEMS
        .Show
    End With
    
End Sub

Public Sub InitSpellEditor()
Dim i As Long

    With frmEditor_Spell
        .lstIndex.Clear
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem (i & ": " & Trim$(Spell(i).Name))
        Next
        
        If ESpellNum = 0 Then
            .lstIndex.ListIndex = 0
            ESpellNum = .lstIndex.ListIndex + 1
        Else
            .lstIndex.ListIndex = ESpellNum - 1
        End If
        .txtName.MaxLength = NAME_LENGTH
        .scrlPicture.max = NumSpells
        .scrlMap.max = MAX_MAPS
        .Show
    End With
End Sub

Public Sub InitItemEditor()
Dim i As Long

    With frmEditor_Item
        .lstIndex.Clear
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem (i & ": " & Trim$(Item(i).Name))
        Next
        
        If EItemNum = 0 Then
            .lstIndex.ListIndex = 0
            EItemNum = .lstIndex.ListIndex + 1
        Else
            .lstIndex.ListIndex = EItemNum - 1
        End If
        .txtInfo.MaxLength = INFO_LENGTH
        .txtName.MaxLength = NAME_LENGTH
        .scrlLearnSpell.max = MAX_SPELLS
        .scrlPImage.max = NumProjectiles
        .scrlPicNum.max = NumItems
        .Show
    End With
End Sub

Public Sub InitNpcEditor()
Dim i As Long

    With frmEditor_Npc
        .lstIndex.Clear
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem (i & ": " & Trim$(Npc(i).Name))
        Next
        
        If ENpcNum = 0 Then
            .lstIndex.ListIndex = 0
            ENpcNum = .lstIndex.ListIndex + 1
        Else
            .lstIndex.ListIndex = ENpcNum - 1
        End If
        .scrlSpeed.max = NumSprites
        .txtName.MaxLength = NAME_LENGTH
        .scrlItem.max = MAX_ITEMS
        .scrlDrop.max = MAX_DROPS
        .scrlDrop.Value = 1
        .Show
    End With
End Sub

Public Sub InitMapEditor()
Dim i As Long

    With frmEditor_Map
        .scrlTileset.max = NumTilesets
        .picTileset.Picture = Nothing
        .picTileset.Picture = LoadPicture(App.Path & "\graphics\tilesets\" & .scrlTileset.Value & ".bmp")
        .optEdit(1).Value = True
        CurrentlyEditing = EDITING_LAYERS
        EditingLayer = Layers.Ground
        EditingAttribute = Attributes.BlockedTile
        .txtMusic = Trim$(Map(Player(MyIndex).Map).Music)
        .txtName = Trim$(Map(Player(MyIndex).Map).Name)
        .cmdMoral.ListIndex = Map(Player(MyIndex).Map).Moral
        .txtUpWarp.text = Map(Player(MyIndex).Map).UpWarp
        .txtDownWarp.text = Map(Player(MyIndex).Map).DownWarp
        .txtLeftWarp.text = Map(Player(MyIndex).Map).LeftWarp
        .txtRightWarp.text = Map(Player(MyIndex).Map).RightWarp
        
        .lstMapNpcs.Clear
        .lstNpcs.Clear
        For i = 1 To MAX_MAP_NPCS
            If Map(Player(MyIndex).Map).MapNpc(i).Num <> 0 Then
                .lstMapNpcs.AddItem (i & ": " & Trim$(Npc(Map(Player(MyIndex).Map).MapNpc(i).Num).Name))
            Else
                .lstMapNpcs.AddItem (i & ": ")
            End If
        Next
        For i = 1 To MAX_NPCS
            .lstNpcs.AddItem (i & ": " & Trim$(Npc(i).Name))
        Next
        .lstMapNpcs.ListIndex = 0
        .lstNpcs.ListIndex = 0
        .Show
    End With

End Sub

Public Sub EditMap(ByVal X As Long, ByVal Y As Long, ByVal Clear As Boolean)
Dim MouseX As Single
Dim MouseY As Single
Dim ConvX As Byte
Dim ConvY As Byte
Dim i As Long
    
    If Options.FullScreen = True Then Exit Sub
    
    MouseX = X / 32
    MouseY = Y / 32

    ' Exit out if out of parameters
    If MouseX < 0 Or MouseX > 480 / 32 Then Exit Sub
    If MouseY < 0 Or MouseY > 384 / 32 Then Exit Sub
    
    ConvX = MouseX
    ConvY = MouseY
    
    ' If the rounded number is bigger than the original number, we must have rounded up. Deduct one
    If ConvX - MouseX > 0 Then ConvX = ConvX - 1
    If ConvY - MouseY > 0 Then ConvY = ConvY - 1
    
    ConvX = ConvX + 1
    ConvY = ConvY + 1
    
    If ConvY > MAX_MAP_Y Or ConvX > MAX_MAP_X Then Exit Sub

    If Clear = False Then
        With Map(Player(MyIndex).Map).Tile(ConvX, ConvY)
            Select Case CurrentlyEditing
                Case EDITING_LAYERS
                    .Layer(EditingLayer).X = CurTileX
                    .Layer(EditingLayer).Y = CurTileY
                    .Layer(EditingLayer).Tileset = frmEditor_Map.scrlTileset.Value
                Case EDITING_ATTRIBUTES
                    .Attribute = EditingAttribute
                    TempPlayer(MyIndex).UnlockedTile(ConvX, ConvY) = False
                    For i = 1 To 4
                        .LongValue(i) = frmEditor_Map.scrlLong(i).Value
                        .StringValue(i) = frmEditor_Map.txtString(i).text
                    Next
                    Select Case .Attribute
                        Case Attributes.NpcSpawnTile
                            With Map(Player(MyIndex).Map).MapNpc(frmEditor_Map.lstMapNpcs.ListIndex + 1)
                                .SpawnX = ConvX
                                .SpawnY = ConvY
                                .Dir = 1
                                .Vital(Vitals.Health) = Npc(frmEditor_Map.lstNpcs.ListIndex + 1).Vital(Vitals.Health)
                                .Vital(Vitals.Spirit) = Npc(frmEditor_Map.lstNpcs.ListIndex + 1).Vital(Vitals.Spirit)
                                If frmEditor_Map.scrlLong(1).Value < 5 And frmEditor_Map.scrlLong(1).Value <> 0 Then .Dir = frmEditor_Map.scrlLong(1).Value
                            End With
                            With TempNpc(Player(MyIndex).Map).NpcNum(frmEditor_Map.lstMapNpcs.ListIndex + 1)
                                .X = ConvX
                                .Y = ConvY
                                .Alive = True
                            End With
                        Case Attributes.ResourceTile
                            With MapResource(Player(MyIndex).Map).Tile(ConvX, ConvY)
                                .Num = frmEditor_Map.scrlLong(1)
                                .Health = Resource(.Num).Health
                                .Alive = True
                            End With
                    End Select
            End Select
        End With
    Else
        With Map(Player(MyIndex).Map).Tile(ConvX, ConvY)
            Select Case CurrentlyEditing
                Case EDITING_LAYERS
                    If EditingLayer = Layers.Ground Then
                        .Layer(EditingAttribute).X = 1
                    Else
                        .Layer(EditingLayer).X = 0
                    End If
                    .Layer(EditingLayer).Y = 0
                    .Layer(EditingLayer).Tileset = 1
                Case EDITING_ATTRIBUTES
                    MapItem(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(0).Num = 0
                    MapItem(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(0).Value = 0
                    MapItem(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(0).Tick = 0
                    MapItem(Player(MyIndex).Map).Tile(ConvX, ConvY).Layer(0).MapItemState = 0
                    .Attribute = 0
                    
                    With MapResource(Player(MyIndex).Map).Tile(ConvX, ConvY)
                        If .Num > 0 Then
                            .Health = 0
                            .Num = 0
                            .Alive = False
                        End If
                    End With
            End Select
        End With
    End If
End Sub

Public Sub InitInterfaceEditor()
Dim i As Long

    With frmEditor_Interface
        .scrlPage.max = MAX_INTERFACE_PAGES
        .scrlPage.Value = 1
        .Show
    End With
End Sub
