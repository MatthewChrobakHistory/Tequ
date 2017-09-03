Attribute VB_Name = "modData"
Option Explicit

Public Sub SaveData()
Dim i As Long

Call SaveDebugOptions

If Player(MyIndex).Access <> ACCESS_ADMIN Then Exit Sub

For i = 1 To MAX_MAPS
    If LoadedMap(i) = True Then
        Call SaveMap(i)
    End If
Next

For i = 1 To MAX_ITEMS
    Call SaveItem(i)
Next

For i = 1 To MAX_NPCS
    Call SaveNpc(i)
Next

For i = 1 To MAX_RESOURCES
    Call SaveResource(i)
Next

For i = 1 To MAX_SHOPS
    Call SaveShop(i)
Next

For i = 1 To MAX_CHESTS
    Call SaveChest(i)
Next

For i = 1 To MAX_SPELLS
    Call SaveSpell(i)
Next

End Sub

Public Sub ClearData()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
        Call ClearMapItem(i)
        Call ClearMapNpc(i)
        LoadedMap(i) = False
    Next
    
    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
    
    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next
    
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next
    
    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
    
    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
    
    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
    
End Sub

Public Sub LoadData()
Dim i As Long

Call LoadDebugOptions

For i = 1 To MAX_ITEMS
    Call LoadItem(i)
Next

For i = 1 To MAX_NPCS
    Call LoadNpc(i)
Next

For i = 1 To MAX_RESOURCES
    Call LoadResource(i)
Next

For i = 1 To MAX_CHESTS
    Call LoadChest(i)
Next

For i = 1 To MAX_SHOPS
    Call LoadShop(i)
Next

For i = 1 To MAX_SPELLS
    Call LoadSpell(i)
Next

Call LoadMap(Player(MyIndex).Map)

End Sub

Public Sub LoadOptions()
Dim FileName As String

    ' Get the filename
    FileName = App.Path & "\options.ini"
    
    ' If the file doesn't exist, save it and then it will continue as normal.
    If FileExist(FileName) = False Then
        SaveOptions (True)
    End If
    
    Options.Debug = GetVar(FileName, "Options", "Debug")
    Options.IP = GetVar(FileName, "Options", "IP")
    Options.Port = GetVar(FileName, "Options", "Port")
    Options.InstallRuntimes = GetVar(FileName, "Options", "InstallRuntimes")
    Options.GameFont = GetVar(FileName, "Options", "GameFont")
    Options.Username = GetVar(FileName, "Options", "Username")
    Options.Password = GetVar(FileName, "Options", "Password")
    Options.Music = GetVar(FileName, "Options", "Music")
    Options.Sound = GetVar(FileName, "Options", "Sound")
    Options.Voices = GetVar(FileName, "Options", "Voices")
    Options.FullScreen = GetVar(FileName, "Options", "FullScreen")
    
    frmMain.lblMusic.Caption = "Music: " & Options.Music & " at " & frmMain.scrlVolume.Value / 10 & " Volume"
    frmMain.lblSound.Caption = "Sound: " & Options.Sound & " at " & frmMain.scrlSFX.Value / 10 & " Volume"
    
End Sub

Public Sub SaveOptions(Optional ByVal NewFile As Boolean = False)
Dim FileName As String
    
    FileName = App.Path & "\options.ini"
    
    If NewFile = True Then
        Options.Debug = 1
        Options.IP = "localhost"
        Options.Port = 7001
        Options.InstallRuntimes = True
        Options.GameFont = FontStyle(8) ' Tamoha
        Options.Username = vbNullString
        Options.Password = vbNullString
        Options.Music = True
        Options.Sound = True
        Options.Voices = True
        Options.FullScreen = False
    End If
    
    Call PutVar(FileName, "Options", "Debug", Str(Options.Debug))
    Call PutVar(FileName, "Options", "IP", Trim$(Options.IP))
    Call PutVar(FileName, "Options", "Port", Str(Options.Port))
    Call PutVar(FileName, "Options", "InstallRuntimes", Str(Options.InstallRuntimes))
    Call PutVar(FileName, "Options", "GameFont", Trim$(Options.GameFont))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "Music", Str(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Str(Options.Sound))
    Call PutVar(FileName, "Options", "Voices", Str(Options.Voices))
    Call PutVar(FileName, "Options", "FullScreen", Str(Options.FullScreen))
    
End Sub

Public Sub MakeAccount(ByVal name As String)
Dim i As Long

    If FileExist(App.Path & "\data\players\" & name & ".bin") = False Then
        MyIndex = 1
        With Player(MyIndex)
            .name = name
            .Access = ACCESS_PLAYER
            .Map = START_MAP
            .X = START_X
            .Y = START_Y
            .Dir = 1
            .Stance = 1
            For i = 1 To Stats.Stat_Count - 1
                .Stat(i) = 1
            Next
            .Combat.level = 1
            For i = 1 To Vitals.Vital_Count - 1
                .Vital(i) = GetPlayerMaxVital(MyIndex, i)
            Next
            For i = 1 To Skills.Skill_Count - 1
                .Skill(i).level = 1
                .Skill(i).XP = 0
            Next
            .Graphics.Gender = GENDER_MALE
        End With
        Call SavePlayer(name)
        ' send them into the game
        Call EnterGame
    Else
        MsgBox "Player already exists!", vbCritical
        Exit Sub
    End If
    
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    ' For multiplayer purposes
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
End Sub

Sub LoadPlayer(ByVal name As String)
Dim FileName As String
Dim F As Long

Call ClearPlayer(MyIndex)

FileName = App.Path & "/data/players/" & name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , Player(MyIndex)
Close #F

Call ClearBank(MyIndex)
FileName = App.Path & "/data/players/" & name & "Bank.bin"
F = FreeFile
Open FileName For Binary As #F
Get #F, , Bank(MyIndex)
Close #F

End Sub

Sub SavePlayer(ByVal name As String)
Dim FileName As String
Dim F As Long

FileName = App.Path & "/data/players/" & name & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Player(MyIndex)
Close #F

FileName = App.Path & "/data/players/" & name & "Bank.bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Bank(MyIndex)
Close #F

End Sub

Sub ClearMap(ByVal Index As Long)
Dim i As Long, X As Long, Y As Long
    Call ZeroMemory(ByVal VarPtr(Map(Index)), LenB(Map(Index)))
    Call ZeroMemory(ByVal VarPtr(MapChest(Index)), LenB(MapChest(Index)))
    Call ZeroMemory(ByVal VarPtr(MapProjectile(Index)), LenB(MapProjectile(Index)))

    For X = 1 To MAX_MAP_X
        For Y = 1 To MAX_MAP_Y
            For i = 1 To MAX_PLAYERS
                TempPlayer(i).UnlockedTile(X, Y) = False
            Next
            
            For i = 1 To MAX_MAP_ITEM_LAYERS
                With MapItem(Index).Tile(X, Y)
                    .Layer(i).Num = 0
                    .Layer(i).MapItemState = 0
                    .Layer(i).Value = 0
                    .Layer(i).Tick = 0
                End With
            Next
            OpenChest(X, Y) = False
        Next
    Next
    
    For X = 1 To MAX_MAP_NPCS
        With TempNpc(Index).NpcNum(X)
            .Target = 0
            .CombatTimer = 0
            .Alive = False
        End With
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(TempNpc(Index)), LenB(TempNpc(Index)))
End Sub

Sub LoadMap(ByVal Index As Long)
Dim FileName As String
Dim F As Long
Dim X As Long, Y As Long

Call ClearMap(Index)

FileName = App.Path & "\data\maps\map" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Map(Index)
Close #F

For X = 1 To MAX_MAP_X
    For Y = 1 To MAX_MAP_Y
        If Map(Index).Tile(X, Y).Attribute = Attributes.ResourceTile Then
            MapResource(Index).Tile(X, Y).Num = Map(Index).Tile(X, Y).LongValue(1)
            MapResource(Index).Tile(X, Y).Health = Resource(MapResource(Index).Tile(X, Y).Num).Health
            MapResource(Index).Tile(X, Y).Alive = True
        End If
        
        With Map(Index).Tile(X, Y).Layer(Layers.Ground)
            If .X = 0 And .Y = 0 Then
                .Tileset = 1
                .X = 1
                .Y = 0
            End If
        End With
    Next
Next

For F = 1 To MAX_MAP_NPCS
    TempNpc(Index).NpcNum(F).X = Map(Index).MapNpc(F).SpawnX
    TempNpc(Index).NpcNum(F).Y = Map(Index).MapNpc(F).SpawnY
    If Map(Index).MapNpc(F).Num > 0 Then
        TempNpc(Index).NpcNum(F).Alive = True
        Map(Index).MapNpc(F).Vital(Vitals.Health) = Npc(Map(Index).MapNpc(F).Num).Vital(Vitals.Health)
        Map(Index).MapNpc(F).Vital(Vitals.Spirit) = Npc(Map(Index).MapNpc(F).Num).Vital(Vitals.Spirit)
    End If
Next

LoadedMap(Index) = True

End Sub

Sub SaveMap(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\maps\map" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Map(Index)
Close #F

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
End Sub

Sub LoadItem(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearItem(Index)

FileName = App.Path & "\data\items\item" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Item(Index)
Close #F
End Sub

Sub SaveItem(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\items\item" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Put #F, , Item(Index)
Close #F
End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
End Sub

Sub LoadNpc(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearNpc(Index)

FileName = App.Path & "\data\npcs\npc" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Npc(Index)
Close #F
End Sub

Sub SaveNpc(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\npcs\npc" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Put #F, , Npc(Index)
Close #F
End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
End Sub

Sub LoadResource(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearResource(Index)

FileName = App.Path & "\data\resources\resource" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Resource(Index)
Close #F

End Sub

Sub SaveResource(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\resources\resource" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Resource(Index)
Close #F

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
End Sub

Sub LoadShop(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearShop(Index)

FileName = App.Path & "\data\shops\shop" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Shop(Index)
Close #F

End Sub

Sub SaveShop(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\shops\shop" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Shop(Index)
Close #F

End Sub

Sub ClearChest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Chest(Index)), LenB(Chest(Index)))
End Sub

Sub LoadChest(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearChest(Index)

FileName = App.Path & "\data\chests\chest" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Chest(Index)
Close #F

End Sub

Sub SaveChest(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\chests\chest" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Chest(Index)
Close #F

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
End Sub

Sub LoadSpell(ByVal Index As Long)
Dim FileName As String
Dim F As Long

Call ClearSpell(Index)

FileName = App.Path & "\data\spells\spell" & Index & ".bin"

F = FreeFile
Open FileName For Binary As #F
Get #F, , Spell(Index)
Close #F

End Sub

Sub SaveSpell(ByVal Index As Long)
Dim FileName As String
Dim F As Long

FileName = App.Path & "\data\spells\spell" & Index & ".bin"
F = FreeFile
Open FileName For Binary As #F
Put #F, , Spell(Index)
Close #F

End Sub

