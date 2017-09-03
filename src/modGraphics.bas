Attribute VB_Name = "modGraphics"
Option Explicit

' Main DirectX Object
Dim DD As DirectDraw7
' Clipper
Public DD_Clip As DirectDrawClipper

' Primary surface
Dim DDS_Primary As DirectDrawSurface7
Dim DDSD_Primary As DDSURFACEDESC2

' Backbuffer
Dim DDS_Backbuffer As DirectDrawSurface7
Dim DDSD_Backbuffer As DDSURFACEDESC2

' Arrays
Dim DDS_Bars As DirectDrawSurface7
Dim DDS_Target As DirectDrawSurface7
Dim DDS_Tileset(1 To NumTilesets) As DirectDrawSurface7
Dim DDS_Item(1 To NumItems) As DirectDrawSurface7
Dim DDS_Sprite(1 To NumSprites) As DirectDrawSurface7
Dim DDS_Paperdoll(1 To NumPaperdolls) As DirectDrawSurface7
Dim DDS_Resource(1 To NumResources) As DirectDrawSurface7
Dim DDS_Spell(1 To NumSpells) As DirectDrawSurface7
Dim DDS_Projectile(1 To NumProjectiles) As DirectDrawSurface7
Dim DDS_Chest(1 To NumChests) As DirectDrawSurface7

Dim DDS_Male_Skin(1 To NumMaleSkin) As DirectDrawSurface7
Dim DDS_Male_NormHair(1 To NumMaleNormHair) As DirectDrawSurface7
Dim DDS_Male_NormBody(1 To NumMaleNormBody) As DirectDrawSurface7
Dim DDS_Male_NormLegs(1 To NumMaleNormLegs) As DirectDrawSurface7
Dim DDS_Female_Skin(1 To NumFemaleSkin) As DirectDrawSurface7
Dim DDS_Female_NormHair(1 To NumFemaleNormHair) As DirectDrawSurface7
Dim DDS_Female_NormBody(1 To NumFemaleNormBody) As DirectDrawSurface7
Dim DDS_Female_NormLegs(1 To NumFemaleNormLegs) As DirectDrawSurface7

' Arrays
Dim DDSD_Bars As DDSURFACEDESC2
Dim DDSD_Target As DDSURFACEDESC2
Dim DDSD_Tileset(1 To NumTilesets) As DDSURFACEDESC2
Dim DDSD_Item(1 To NumItems) As DDSURFACEDESC2
Dim DDSD_Sprite(1 To NumSprites) As DDSURFACEDESC2
Dim DDSD_Paperdoll(1 To NumPaperdolls) As DDSURFACEDESC2
Dim DDSD_Resource(1 To NumResources) As DDSURFACEDESC2
Dim DDSD_Spell(1 To NumSpells) As DDSURFACEDESC2
Dim DDSD_Projectile(1 To NumProjectiles) As DDSURFACEDESC2
Dim DDSD_Chest(1 To NumChests) As DDSURFACEDESC2

Dim DDSD_Male_Skin(1 To NumMaleSkin) As DDSURFACEDESC2
Dim DDSD_Male_NormHair(1 To NumMaleNormHair) As DDSURFACEDESC2
Dim DDSD_Male_NormBody(1 To NumMaleNormBody) As DDSURFACEDESC2
Dim DDSD_Male_NormLegs(1 To NumMaleNormLegs) As DDSURFACEDESC2
Dim DDSD_Female_Skin(1 To NumFemaleSkin) As DDSURFACEDESC2
Dim DDSD_Female_NormHair(1 To NumFemaleNormHair) As DDSURFACEDESC2
Dim DDSD_Female_NormBody(1 To NumFemaleNormBody) As DDSURFACEDESC2
Dim DDSD_Female_NormLegs(1 To NumFemaleNormLegs) As DDSURFACEDESC2


Public Const PlayerSpriteWidth As Byte = 72
Public Const PlayerSpriteHeight As Byte = 91

Public DDSD_Temp As DDSURFACEDESC2

Private LoadedDX7 As Boolean

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim Rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT

    If frmMain.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If LoadedDX7 = False Then
        Exit Sub
    End If
    
    If DD.TestCooperativeLevel <> DD_OK Then
        Call DestroyDirectX
        Call InitDirectX
        Exit Sub
    End If
    
    If CreatingCharacter = True Then
    
            ' rec_pos
        With rec_pos
            .Bottom = 416
            .Right = 512
        End With
    
        ' fill it with black
        DDS_Backbuffer.BltColorFill rec_pos, 0
        
        Call BltCC
        
        TexthDC = DDS_Backbuffer.GetDC
        
        'text here
        
        DDS_Backbuffer.ReleaseDC TexthDC
        
        ' Get rec
        With Rec
            .top = 32
            .Bottom = 416
            .Left = 32
            .Right = 512
        End With
        

        
        ' Render
        DX7.GetWindowRect frmMain.picScreen.hWnd, rec_pos
        
        With rec_pos
            frmMain.Caption = "L:" & .Left & " R:" & .Right & " T:" & .top & " B:" & .Bottom
        End With
        
        DDS_Primary.Blt rec_pos, DDS_Backbuffer, Rec, DDBLT_WAIT
        Exit Sub
    End If
    
    Call BltMap(Player(MyIndex).Map)
    
    TexthDC = DDS_Backbuffer.GetDC
    
    If frmEditor_Map.Visible = True Then
        For X = 0 To MAX_MAP_X
            For Y = 0 To MAX_MAP_Y
                Call DrawMapTileAttribute(X, Y)
            Next
        Next
    End If

    ' BLT TEXT
    If TempPlayer(MyIndex).DrawCoords = True Then
        Call DrawCoords
    End If
    
    If Len(Trim$(Map(Player(MyIndex).Map).name)) > 0 Then
        If Trim$(Map(Player(MyIndex).Map).name) <> vbNullString Then
            Call DrawMapName
        End If
    End If
    
    For X = 1 To MAX_PLAYERS
        If Player(X).Map = Player(MyIndex).Map Then
            Call DrawPlayerName(X, PlayerSpriteHeight)
        End If
    Next
    
    For X = 1 To MAX_MAP_NPCS
        If Map(Player(MyIndex).Map).MapNpc(X).Num > 0 Then
            If Npc(Map(Player(MyIndex).Map).MapNpc(X).Num).Sprite <> 0 Then
                Call DrawNpcName(X, DDSD_Sprite(Npc(Map(Player(MyIndex).Map).MapNpc(X).Num).Sprite).lWidth / 4)
            End If
        End If
    Next
    
    For X = 1 To 255
        If TempActionMsg(X).Created > 0 Then
            Call DrawActionMsg(X)
        End If
    Next

    DDS_Backbuffer.ReleaseDC TexthDC
    
    ' Get rec
    With Rec
        .top = 32
        .Bottom = 416
        .Left = 32
        .Right = 512
    End With
    
    ' rec_pos
    With rec_pos
        .Bottom = 416
        .Right = 512
        
        If Options.FullScreen = True Then
            .Bottom = .Bottom * 2
            .Right = .Right * 2
        End If
    End With
    
    ' Render
    DX7.GetWindowRect frmMain.picScreen.hWnd, rec_pos
    DDS_Primary.Blt rec_pos, DDS_Backbuffer, Rec, DDBLT_WAIT

End Sub

Public Sub DestroyDirectX()
Dim i As Long

    Set DDS_Bars = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Bars), LenB(DDSD_Bars)
    
    Set DDS_Target = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Target), LenB(DDSD_Target)
    
    For i = 1 To NumProjectiles
        Set DDS_Projectile(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Projectile(i)), LenB(DDSD_Projectile(i))
    Next
    
    For i = 1 To NumSpells
        Set DDS_Spell(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Spell(i)), LenB(DDSD_Spell(i))
    Next
    
    For i = 1 To NumTilesets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next
    
    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next
    
    For i = 1 To NumSprites
        Set DDS_Sprite(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Sprite(i)), LenB(DDSD_Sprite(i))
    Next
    
    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next
    
    Set DDS_Primary = Nothing
    Set DDS_Backbuffer = Nothing
    Set DD = Nothing
    
End Sub

Public Sub InitDirectX()
    
    ' Clear DirectX7
    Call DestroyDirectX
    
    Set DD = DX7.DirectDrawCreate(vbNullString)
    
    DD.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
    
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1 ' One Backbuffer
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    Set DD_Clip = DD.CreateClipper(0)

    DD_Clip.SetHWnd frmMain.picScreen.hWnd
    DDS_Primary.SetClipper DD_Clip
    
    Call InitSurfaces
    
End Sub

Public Sub InitSurfaces()

    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    Set DDS_Backbuffer = Nothing
    
    With DDSD_Backbuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (14 + 3) * 32
        .lHeight = (11 + 3) * 32
    End With
    Set DDS_Backbuffer = DD.CreateSurface(DDSD_Backbuffer)
    
    ' Load persistent surfaces
    
    LoadedDX7 = True
    
End Sub

Public Sub BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
    
    If Not ddS Is Nothing Then
        Call DDS_Backbuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If
    
End Sub

Public Function BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
    
    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh
    BltToDC = True
    
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, Y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
    
End Sub

Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    
    ' Set path
    FileName = App.Path & "\graphics\" & FileName & ".bmp"

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)
    
    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)

End Sub

Public Sub LoadGraphics()
Dim i As Long

    Call InitDDSurf("bars", DDSD_Bars, DDS_Bars)
    Call InitDDSurf("target", DDSD_Target, DDS_Target)
    
    ' Projectiles
    For i = 1 To NumProjectiles
        Call InitDDSurf("projectiles\" & i, DDSD_Projectile(i), DDS_Projectile(i))
    Next
    
    ' Spells
    For i = 1 To NumSpells
        Call InitDDSurf("spells\" & i, DDSD_Spell(i), DDS_Spell(i))
    Next
    
    ' Resources
    For i = 1 To NumResources
        Call InitDDSurf("resources\" & i, DDSD_Resource(i), DDS_Resource(i))
    Next
    
    ' Chests
    For i = 1 To NumChests
        Call InitDDSurf("chests\" & i, DDSD_Chest(i), DDS_Chest(i))
    Next

    ' Tilesets
    For i = 1 To NumTilesets
        Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
    Next
    
    ' Items
    For i = 1 To NumItems
        Call InitDDSurf("items\" & i, DDSD_Item(i), DDS_Item(i))
    Next
    
    ' Sprites
    For i = 1 To NumSprites
        Call InitDDSurf("sprites\" & i, DDSD_Sprite(i), DDS_Sprite(i))
    Next
    
    ' Paperdolls
    For i = 1 To NumPaperdolls
        Call InitDDSurf("paperdolls\" & i, DDSD_Paperdoll(i), DDS_Paperdoll(i))
    Next
    
    ' Male Skin
    For i = 1 To NumMaleSkin
        Call InitDDSurf("player\male\skin\" & i, DDSD_Male_Skin(i), DDS_Male_Skin(i))
    Next
    
    ' Female Skin
    For i = 1 To NumFemaleSkin
        Call InitDDSurf("player\female\skin\" & i, DDSD_Female_Skin(i), DDS_Female_Skin(i))
    Next
    
    ' Male Body
    For i = 1 To NumMaleNormBody
        Call InitDDSurf("player\male\body\norm\" & i, DDSD_Male_NormBody(i), DDS_Male_NormBody(i))
    Next
    
    ' Female Body
    For i = 1 To NumFemaleNormBody
        Call InitDDSurf("player\female\body\norm\" & i, DDSD_Female_NormBody(i), DDS_Female_NormBody(i))
    Next
    
    ' Hair Male
    For i = 1 To NumMaleNormHair
        Call InitDDSurf("player\male\hair\norm\" & i, DDSD_Male_NormHair(i), DDS_Male_NormHair(i))
    Next
    
    ' Female Hair
    For i = 1 To NumFemaleNormHair
        Call InitDDSurf("player\female\hair\norm\" & i, DDSD_Female_NormHair(i), DDS_Female_NormHair(i))
    Next
    
    ' Male Legs
    For i = 1 To NumMaleNormLegs
        Call InitDDSurf("player\male\legs\norm\" & i, DDSD_Male_NormLegs(i), DDS_Male_NormLegs(i))
    Next
    
    ' Female Legs
    For i = 1 To NumFemaleNormLegs
        Call InitDDSurf("player\female\legs\norm\" & i, DDSD_Female_NormLegs(i), DDS_Female_NormLegs(i))
    Next

End Sub

Public Sub BltInventory()
Dim i As Long, X As Long, Y As Long, Itemnum As Long, itemPic As Long
Dim Amount As Long
Dim itemPicRec As RECT, positionRec As RECT
Dim color As Long

    frmMain.picInventory.Cls
    
    For i = 1 To MAX_INV
        Itemnum = Player(MyIndex).Inv(i).Num
        
        If Itemnum > 0 Then
            itemPic = Item(Itemnum).Picture
            
            If itemPic > 0 And itemPic <= NumItems Then
                If DDSD_Item(itemPic).lWidth <= 96 Then
                
                    With itemPicRec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
                    
                    With positionRec
                        .top = 1 + (35 * ((i - 1) \ 5))
                        .Bottom = .top + 32
                        .Left = 12 + (35 * (((i - 1) Mod 5)))
                        .Right = .Left + 32
                    End With
                    
                    If DDS_Item(itemPic) Is Nothing Then
                        Call InitDDSurf("items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                    End If
                    
                    BltToDC DDS_Item(itemPic), itemPicRec, positionRec, frmMain.picInventory, False
                    

                    Amount = Player(MyIndex).Inv(i).Value
                    If Amount = 0 Then
                        Player(MyIndex).Inv(i).Value = 1
                        Amount = 1
                    End If
                    
                    If Amount < 1000000 Then
                        color = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        color = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        color = QBColor(BrightGreen)
                    End If
                    
                    If Amount > 1 Then
                        Y = positionRec.top + 22
                        X = positionRec.Left - 6
                        DrawText frmMain.picInventory.hDC, X, Y, TransformAmount(Str(Amount)), color
                    End If
                    
                End If
            End If
        End If
    Next
    
    frmMain.picInventory.Refresh
            
End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long, ByVal Layer As Long)
Dim Rec As DxVBLib.RECT
Dim Mapnum As Long
Dim TilesetNum As Byte

Mapnum = Player(MyIndex).Map
If Mapnum = 0 Then Exit Sub

    'If Map(MapNum).Tile(X, Y).Layer(Layer).X > 0 Or Map(MapNum).Tile(X, Y).Layer(Layer).Y > 0 Then ' There has to be an image set to blt
        'set the rec
        
        With Rec
            Rec.top = Map(Mapnum).Tile(X, Y).Layer(Layer).Y * 32
            Rec.Bottom = .top + 32
            Rec.Left = Map(Mapnum).Tile(X, Y).Layer(Layer).X * 32
            Rec.Right = .Left + 32
        End With
        
        TilesetNum = Map(Mapnum).Tile(X, Y).Layer(Layer).Tileset
        If TilesetNum = 0 Then Exit Sub
        
        Call BltFast(X * 32, Y * 32, DDS_Tileset(TilesetNum), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    'End If
End Sub

Public Sub BltPlayer(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
Dim Rec As DxVBLib.RECT
Dim b As Long, d As Byte
Dim Skin As Byte, Legs As Byte, Body As Byte, Hair As Byte

    Skin = Player(Index).Graphics.Skin
    Legs = Player(Index).Graphics.Legs
    Body = Player(Index).Graphics.Body
    Hair = Player(Index).Graphics.Hair

    If TempPlayer(Index).Step = 0 Then TempPlayer(Index).Step = 1
    With Rec
        .top = (Player(Index).Dir - 1) * PlayerSpriteHeight
        .Left = (TempPlayer(Index).Step - 1) * PlayerSpriteWidth
        .Right = .Left + PlayerSpriteWidth
        .Bottom = .top + PlayerSpriteHeight
    End With
    
    X = Player(Index).X * 32 + TempPlayer(Index).XOffset - ((PlayerSpriteWidth - 32) / 2)
        
    If PlayerSpriteHeight > 32 Then
        Y = Player(Index).Y * 32 + TempPlayer(Index).YOffset - (PlayerSpriteHeight - 32)
        If Y < 0 Then
            With Rec
                .top = .top - Y
            End With
            Y = 0
        End If
    Else
        Y = Player(Index).Y * 32 + TempPlayer(Index).YOffset
    End If

    Select Case Player(Index).Graphics.Gender
        Case GENDER_MALE
            Call BltFast(X, Y, DDS_Male_Skin(Skin), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Select Case Player(Index).Graphics.BodyDir
                Case "norm"
                    If Player(Index).Equipment(Equipment.Body).Num = 0 Then
                        Call BltFast(X, Y, DDS_Male_NormBody(Body), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        If Item(Player(Index).Equipment(Equipment.Body).Num).BltPlayerGraphics = True Then
                            Call BltFast(X, Y, DDS_Male_NormBody(Body), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
            End Select
            Select Case Player(Index).Graphics.LegsDir
                Case "norm"
                    If Player(Index).Equipment(Equipment.Legs).Num = 0 Then
                        Call BltFast(X, Y, DDS_Male_NormLegs(Legs), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        If Item(Player(Index).Equipment(Equipment.Legs).Num).BltPlayerGraphics = True Then
                            Call BltFast(X, Y, DDS_Male_NormLegs(Legs), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
            End Select
            Select Case Player(Index).Graphics.HairDir
                Case "norm"
                    If Player(Index).Equipment(Equipment.Head).Num = 0 Then
                        Call BltFast(X, Y, DDS_Male_NormHair(Hair), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        If Item(Player(Index).Equipment(Equipment.Head).Num).BltPlayerGraphics = True Then
                            Call BltFast(X, Y, DDS_Male_NormHair(Hair), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
            End Select
        Case GENDER_FEMALE
            Call BltFast(X, Y, DDS_Female_Skin(Skin), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Select Case Player(Index).Graphics.BodyDir
                Case "norm"
                    If Player(Index).Equipment(Equipment.Body).Num = 0 Then
                        Call BltFast(X, Y, DDS_Female_NormBody(Body), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        If Item(Player(Index).Equipment(Equipment.Body).Num).BltPlayerGraphics = True Then
                            Call BltFast(X, Y, DDS_Female_NormBody(Body), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
            End Select
            Select Case Player(Index).Graphics.LegsDir
                Case "norm"
                    If Player(Index).Equipment(Equipment.Legs).Num = 0 Then
                        Call BltFast(X, Y, DDS_Female_NormLegs(Legs), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        If Item(Player(Index).Equipment(Equipment.Legs).Num).BltPlayerGraphics = True Then
                            Call BltFast(X, Y, DDS_Female_NormLegs(Legs), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
            End Select
            Select Case Player(Index).Graphics.HairDir
                Case "norm"
                    If Player(Index).Equipment(Equipment.Head).Num = 0 Then
                        Call BltFast(X, Y, DDS_Female_NormHair(Hair), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        If Item(Player(Index).Equipment(Equipment.Head).Num).BltPlayerGraphics = True Then
                            Call BltFast(X, Y, DDS_Female_NormHair(Hair), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
            End Select
    End Select
    
    Call BltEquipment(Index, Player(Index).X, Player(Index).Y)
    
End Sub

Public Sub BltMap(ByVal Mapnum As Long)
Dim L As Long, X As Long, Y As Long, i As Long
Dim NpcNum As Long, Index As Long

If Mapnum = 0 Then Exit Sub

    For L = 1 To Layers.MaskAnim
        For X = 1 To MAX_MAP_X
            For Y = 1 To MAX_MAP_Y
                Call BltMapTile(X, Y, L)
            Next
        Next
    Next
    
    ' blt map item
    For X = 1 To MAX_MAP_X
        For Y = 1 To MAX_MAP_Y
            For i = 0 To MAX_MAP_ITEM_LAYERS
                If MapItem(Mapnum).Tile(X, Y).Layer(i).Num > 0 Then
                    Call BltMapItem(X, Y, MapItem(Mapnum).Tile(X, Y).Layer(i).Num)
                End If
            Next
        Next
    Next
    
    For Y = 1 To Player(MyIndex).Y - 1
        For i = 1 To MAX_MAP_PROJECTILES
            If MapProjectile(Mapnum).MapProjectile(i).Y = Y Then
                Call BltProjectile(i)
            End If
        Next
        For NpcNum = 1 To MAX_NPCS
            If Map(Mapnum).MapNpc(NpcNum).Num > 0 Then
                If TempNpc(Mapnum).NpcNum(NpcNum).Y = Y Then
                    Call BltNpc(NpcNum, TempNpc(Mapnum).NpcNum(NpcNum).X, Y)
                End If
            End If
        Next
        For X = 1 To MAX_MAP_X
            Select Case Map(Mapnum).Tile(X, Y).Attribute
                Case Attributes.ChestTile
                    If OpenChest(X, Y) = True Then
                        Call BltChest(Map(Mapnum).Tile(X, Y).LongValue(1), X, Y, True)
                    Else ' false
                        Call BltChest(Map(Mapnum).Tile(X, Y).LongValue(1), X, Y, False)
                    End If
                Case Attributes.ResourceTile
                    If MapResource(Mapnum).Tile(X, Y).Num > 0 Then
                        Call BltResource(MapResource(Mapnum).Tile(X, Y).Num, X, Y)
                    End If
            End Select
        Next
        For Index = 1 To MAX_PLAYERS
            If Player(Index).Map = Player(MyIndex).Map Then
                If Player(Index).Y = Y Then
                    Call BltPlayer(Index, Player(Index).X, Player(Index).Y)
                End If
            End If
        Next
    Next
    
    For Y = Player(MyIndex).Y To MAX_MAP_Y
        For NpcNum = 1 To MAX_NPCS
            If Map(Mapnum).MapNpc(NpcNum).Num > 0 Then
                If TempNpc(Mapnum).NpcNum(NpcNum).Y = Y Then
                    Call BltNpc(NpcNum, TempNpc(Mapnum).NpcNum(NpcNum).X, Y)
                End If
            End If
        Next
        For X = 1 To MAX_MAP_X
            Select Case Map(Mapnum).Tile(X, Y).Attribute
                Case Attributes.ChestTile
                    If OpenChest(X, Y) = True Then
                        Call BltChest(Map(Mapnum).Tile(X, Y).LongValue(1), X, Y, True)
                    Else ' false
                        Call BltChest(Map(Mapnum).Tile(X, Y).LongValue(1), X, Y, False)
                    End If
                Case Attributes.ResourceTile
                    If MapResource(Mapnum).Tile(X, Y).Num > 0 Then
                        Call BltResource(MapResource(Mapnum).Tile(X, Y).Num, X, Y)
                    End If
            End Select
        Next
        For Index = 1 To MAX_PLAYERS
            If Player(Index).Map = Player(MyIndex).Map Then
                If Player(Index).Y = Y Then
                    Call BltPlayer(Index, Player(Index).X, Player(Index).Y)
                End If
            End If
        Next
        For i = 1 To MAX_MAP_PROJECTILES
            If MapProjectile(Mapnum).MapProjectile(i).Y = Y Then
                Call BltProjectile(i)
            End If
        Next
    Next
    
    For L = Layers.MaskAnim To Layers.Layer_Count
        For X = 0 To MAX_MAP_X
            For Y = 0 To MAX_MAP_Y
                Call BltMapTile(X, Y, L)
            Next
        Next
    Next
    
    ' Blting Bars
    For i = 1 To MAX_PLAYERS
        If TempPlayer(i).CombatTimer > 0 Then
            Call BltBars(i)
        End If
    Next
    
    For i = 1 To MAX_MAP_NPCS
        If TempNpc(Player(MyIndex).Map).NpcNum(i).CombatTimer > 0 Then
            Call BltNpcBars(i)
        End If
    Next
    
    If TempPlayer(MyIndex).HoverTarget > 0 Then
        Call BltHoverTarget(TempPlayer(MyIndex).HoverTarget)
    End If
    If TempPlayer(MyIndex).Target > 0 Then
        Call BltTarget(TempPlayer(MyIndex).Target)
    End If
    
End Sub

Public Sub BltMapItem(ByVal X As Long, ByVal Y As Long, ByVal Num As Long)
Dim Rec As DxVBLib.RECT
Dim Mapnum As Long

Mapnum = Player(MyIndex).Map

    If Item(Num).Picture > 0 Then ' There has to be an image set to blt
        'set the rec
        
        With Rec
            Rec.top = 0
            Rec.Bottom = .top + 32
            Rec.Left = 0
            Rec.Right = .Left + 32
        End With
        
        Call BltFast(X * 32, Y * 32, DDS_Item(Item(Num).Picture), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        
    End If
End Sub

Public Sub BltCharacterScreen()
Dim i As Long, X As Long, Y As Long, Itemnum As Long, itemPic As Long
Dim Amount As Long
Dim itemPicRec As RECT, positionRec As RECT
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long
Dim TopOffset As Single, LeftOffset As Single

    frmMain.picCharacter.Cls
    
    For i = 1 To Equipment.Equipment_Count - 1
        Itemnum = Player(MyIndex).Equipment(i).Num
        
        If Itemnum > 0 Then
            itemPic = Item(Itemnum).Picture
            amountModifier = 0
            
            If itemPic > 0 And itemPic <= NumItems Then
                If DDSD_Item(itemPic).lWidth <= 96 Then
                
                    LeftOffset = i / 3
                    TopOffset = i / 3
                    If TopOffset <= 1 Then
                        TopOffset = 48
                    ElseIf TopOffset <= 2 Then
                        TopOffset = 88
                    ElseIf TopOffset <= 3 Then
                        TopOffset = 128
                    End If
                    If i = 1 Or i = 4 Or i = 7 Then
                        LeftOffset = 32
                    ElseIf i = 2 Or i = 5 Or i = 8 Then
                        LeftOffset = 80
                    ElseIf i = 3 Or i = 6 Or i = 9 Then
                        LeftOffset = 128
                    End If
                
                    With itemPicRec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
                    
                    With positionRec
                        .top = TopOffset
                        .Bottom = .top + 32
                        .Left = LeftOffset
                        .Right = .Left + 32
                    End With
                    
                    If DDS_Item(itemPic) Is Nothing Then
                        Call InitDDSurf("items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                    End If
                    BltToDC DDS_Item(itemPic), itemPicRec, positionRec, frmMain.picCharacter, False
                    
                    
                    Amount = Player(MyIndex).Equipment(i).Value
                    If Amount > 1 Then
                    
                        Y = positionRec.top + 22
                        X = positionRec.Left - 4
                        
                        DrawText frmMain.picCharacter.hDC, X, Y, Format$(Amount, "#,###,###,###"), QBColor(White)
                    End If
                    
                End If
            End If
        End If
    Next
            
End Sub

Public Sub RenderInfo(ByVal Itemnum As Long)
Dim i As Long, X As Long, Y As Long, itemPic As Long
Dim Amount As Long
Dim itemPicRec As RECT, positionRec As RECT

    If Game.InShop = False Then If Options.DIOC = False Then Exit Sub
    
    frmMain.picInfo.Visible = True
    frmMain.picInfo.Cls
    
    If Itemnum > 0 Then
        itemPic = Item(Itemnum).Picture
            If DDSD_Item(itemPic).lWidth <= 96 Then
                With itemPicRec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With
                
                With positionRec
                    .top = 4
                    .Bottom = .top + 32
                    .Left = 14
                    .Right = .Left + 32
                End With
                
                If DDS_Item(itemPic) Is Nothing Then
                    Call InitDDSurf("items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                End If
                
                BltToDC DDS_Item(itemPic), itemPicRec, positionRec, frmMain.picInfo, False
                frmMain.lblInfoName.Caption = Trim$(Item(Itemnum).name)
                frmMain.lblPrice.Caption = "Worth " & Item(Itemnum).Price & " gold"
                frmMain.lblInfoText.Caption = Trim$(Item(Itemnum).info)
            End If
    End If
            
End Sub

Public Sub BltEquipment(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
Dim Rec As DxVBLib.RECT
Dim Paperdoll(1 To Equipment.Equipment_Count - 1) As Long
Dim Order(1 To Equipment.Equipment_Count - 1) As Byte
Dim i As Long

    For i = 1 To Equipment.Equipment_Count - 1
        If Player(Index).Equipment(i).Num > 0 Then
            Paperdoll(i) = Item(Player(Index).Equipment(i).Num).Paperdoll(Player(Index).Stance)
        End If
    Next
    
    Select Case Player(Index).Dir
        Case DIR_DOWN
            Order(1) = Equipment.Shield
            Order(7) = Equipment.Weapon
        Case DIR_RIGHT
            Order(1) = Equipment.Shield
            Order(7) = Equipment.Weapon
        Case DIR_UP
            Order(7) = Equipment.Weapon
            Order(1) = Equipment.Shield
        Case DIR_LEFT
            Order(7) = Equipment.Weapon
            Order(1) = Equipment.Shield
    End Select
    
    Order(2) = Equipment.Boots
    Order(3) = Equipment.Hands
    Order(4) = Equipment.Legs
    Order(5) = Equipment.Body
    Order(6) = Equipment.Head
    
    ' Paperdoll stores the paperdoll. Index is the equipment slot.
    ' Order stores the order in which we blt. Index is the priority. 1 = first to blt
    For i = 1 To Equipment.Equipment_Count - 1
        If Order(i) > 0 Then
            If Paperdoll(Order(i)) > 0 Then
                With Rec
                    .top = (Player(Index).Dir - 1) * DDSD_Paperdoll(Paperdoll(Order(i))).lHeight / 4
                    .Left = (TempPlayer(Index).Step - 1) * DDSD_Paperdoll(Paperdoll(Order(i))).lWidth / 4
                    .Right = .Left + DDSD_Paperdoll(Paperdoll(Order(i))).lWidth / 4
                    .Bottom = .top + DDSD_Paperdoll(Paperdoll(Order(i))).lHeight / 4
                End With
                X = Player(Index).X * 32 + TempPlayer(Index).XOffset - ((DDSD_Paperdoll(Paperdoll(Order(i))).lWidth / 4 - 32) / 2)
                
                If (DDSD_Paperdoll(Paperdoll(Order(i))).lHeight) > 32 Then
                    Y = Player(Index).Y * 32 + TempPlayer(Index).YOffset - ((DDSD_Paperdoll(Paperdoll(Order(i))).lHeight / 4) - 32)
                    If Y < 0 Then
                        With Rec
                            .top = .top - Y
                        End With
                        Y = 0
                    End If
                Else
                    Y = Player(Index).Y * 32 + TempPlayer(Index).YOffset
                End If
                
                Call BltFast(X, Y, DDS_Paperdoll(Paperdoll(Order(i))), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Next
        
End Sub

Public Sub BltCC()
Dim GeneralRec As DxVBLib.RECT
Dim Skin As Byte, Legs As Byte, Hair As Byte, Body As Byte, Gender As Byte
Dim LegsDir As String, HairDir As String, BodyDir As String
Dim X As Long, Y As Long

    Skin = Player(MyIndex).Graphics.Skin
    Legs = Player(MyIndex).Graphics.Legs
    LegsDir = Player(MyIndex).Graphics.LegsDir
    Hair = Player(MyIndex).Graphics.Hair
    HairDir = Player(MyIndex).Graphics.HairDir
    Body = Player(MyIndex).Graphics.Body
    BodyDir = Player(MyIndex).Graphics.BodyDir
    Gender = Player(MyIndex).Graphics.Gender
    
    X = 237
    Y = 160
    
    With GeneralRec
        .top = 0
        .Left = 0
        .Right = 72
        .Bottom = 91
    End With
    
    ' blt the skin first
    Select Case Player(MyIndex).Graphics.Gender
        Case GENDER_MALE
            Call BltFast(X, Y, DDS_Male_Skin(Skin), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Select Case BodyDir
                Case "norm"
                    Call BltFast(X, Y, DDS_Male_NormBody(Body), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End Select
            Select Case LegsDir
                Case "norm"
                    Call BltFast(X, Y, DDS_Male_NormLegs(Legs), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End Select
            Select Case HairDir
                Case "norm"
                    Call BltFast(X, Y, DDS_Male_NormHair(Hair), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End Select
        Case GENDER_FEMALE
            Call BltFast(X, Y, DDS_Female_Skin(Skin), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Select Case BodyDir
                Case "norm"
                    Call BltFast(X, Y, DDS_Female_NormBody(Body), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End Select
            Select Case LegsDir
                Case "norm"
                    Call BltFast(X, Y, DDS_Female_NormLegs(Legs), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End Select
            Select Case HairDir
                Case "norm"
                    Call BltFast(X, Y, DDS_Female_NormHair(Hair), GeneralRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End Select
    End Select
End Sub

Public Sub BltNpc(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
Dim Rec As DxVBLib.RECT
Dim b As Long, d As Byte
Dim Sprite As Long, NpcNum As Long
Dim Mapnum As Long

    Mapnum = Player(MyIndex).Map
    NpcNum = Map(Mapnum).MapNpc(Index).Num
    If NpcNum = 0 Then Exit Sub
    Sprite = Npc(NpcNum).Sprite
    
    If Sprite = 0 Then Exit Sub

    If TempNpc(Mapnum).NpcNum(Index).Step = 0 Then TempNpc(Mapnum).NpcNum(Index).Step = 1
    If Map(Mapnum).MapNpc(Index).Dir = 0 Then
        Map(Mapnum).MapNpc(Index).Dir = 1
    End If
    With Rec
        .top = (Map(Mapnum).MapNpc(Index).Dir - 1) * DDSD_Sprite(Sprite).lHeight / 4
        .Left = (TempNpc(Mapnum).NpcNum(Index).Step - 1) * DDSD_Sprite(Sprite).lWidth / 4
        .Right = .Left + DDSD_Sprite(Sprite).lWidth / 4
        .Bottom = .top + DDSD_Sprite(Sprite).lHeight / 4
    End With
    
    X = TempNpc(Mapnum).NpcNum(Index).X * 32 + TempNpc(Mapnum).NpcNum(Index).XOffset - ((DDSD_Sprite(Sprite).lWidth / 4 - 32) / 2)
        
    If DDSD_Sprite(Sprite).lHeight / 4 > 32 Then
        Y = TempNpc(Mapnum).NpcNum(Index).Y * 32 + TempNpc(Mapnum).NpcNum(Index).YOffset - (DDSD_Sprite(Sprite).lHeight / 4 - 32)
        If Y < 0 Then
            With Rec
                .top = .top - Y
            End With
            Y = 0
        End If
    Else
        Y = TempNpc(Mapnum).NpcNum(Index).Y * 32 + TempNpc(Mapnum).NpcNum(Index).YOffset
    End If

    Call BltFast(X, Y, DDS_Sprite(Sprite), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub

Public Sub RenderBank()
Dim sRECT As RECT, dRECT As RECT
Dim Itemnum As Long, Image As Long
Dim Amount As Long, X As Long, Y As Long
Dim i As Long, color As Long

    If frmMain.picBank.Visible = True Then
        frmMain.picBank.Cls
        If CurTab = 0 Then CurTab = 1
        If WIMultiplier = 0 Then WIMultiplier = 1
        
        For i = 1 To MAX_BANK_TABS
            If Bank(MyIndex).BankTab(i).BankItem(1).Num > 0 Then
                Itemnum = Bank(MyIndex).BankTab(i).BankItem(1).Num
                If Itemnum > 0 And Itemnum <= MAX_ITEMS Then
                    Image = Item(Itemnum).Picture
                    
                    If Image = 0 Then Exit Sub
                    
                    If DDS_Item(Image) Is Nothing Then
                        Call InitDDSurf("Items\" & Image, DDSD_Item(Image), DDS_Item(Image))
                    End If
                    
                    With sRECT
                        .top = 0
                        .Bottom = .top + 32
                        .Left = (DDSD_Item(Image).lWidth / 3) * 2
                        .Right = .Left + 32
                    End With
                    
                    With dRECT
                        .Left = 34 + (i * 8) + (i - 1) + ((i - 1) * 32) - 4
                        .Right = .Left + 32
                        .top = 7
                        .Bottom = .top + 32
                    End With
                    
                    BltToDC DDS_Item(Image), sRECT, dRECT, frmMain.picBank, False
                End If
            End If
        Next
        
        For i = 1 To MAX_BANK_ITEMS
            Itemnum = Bank(MyIndex).BankTab(CurTab).BankItem(i).Num
            If Itemnum > 0 And Itemnum <= MAX_ITEMS Then
            
                Image = Item(Itemnum).Picture
                
                If Image = 0 Then Exit Sub
                
                If DDS_Item(Image) Is Nothing Then
                    Call InitDDSurf("Items\" & Image, DDSD_Item(Image), DDS_Item(Image))
                End If
            
                With sRECT
                    .top = 0
                    .Bottom = .top + 32
                    .Left = DDSD_Item(Image).lWidth / 3
                    .Right = .Left + 32
                End With
                
                With dRECT
                    .top = 38 + (35 * ((i - 1) \ 11)) + 1
                    .Bottom = .top + 32
                    .Left = 42 + (36 * (((i - 1) Mod 11)))
                    .Right = .Left + 32
                End With
                
                BltToDC DDS_Item(Image), sRECT, dRECT, frmMain.picBank, False
                
                Amount = Bank(MyIndex).BankTab(CurTab).BankItem(i).Value
                If Amount > 0 Then
            
                    If Amount < 1000000 Then
                        color = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        color = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        color = QBColor(BrightGreen)
                    End If
                    
                    Y = dRECT.top + 22
                    X = dRECT.Left - 4
                        
                    DrawText frmMain.picBank.hDC, X, Y, TransformAmount(Str(Amount)), color
                End If
                
            End If
        Next
        
        If WIMultiplier = 2147483647 Then
            DrawText frmMain.picBank.hDC, 16, frmMain.picBank.Height - 19, "Multiplier: All", QBColor(White)
        Else
            DrawText frmMain.picBank.hDC, 16, frmMain.picBank.Height - 19, "Multiplier: x" & WIMultiplier, QBColor(White)
        End If
        
        frmMain.picBank.Refresh
    End If
    
End Sub

Public Sub RenderShop(ByVal Index As Long)
Dim sRECT As RECT, dRECT As RECT
Dim Itemnum As Long, Image As Long
Dim Amount As Long, X As Long, Y As Long
Dim i As Long, color As Long
    
    If Game.InShop = True Then
        Game.ShopState = SHOP_STATE_NONE
        Game.ShopNum = Index
        frmMain.picShop.Picture = Nothing
        If Shop(Index).Picture > 0 Then frmMain.picShop.Picture = LoadPicture(App.Path & "\graphics\gui\main\shop\" & Shop(Index).Picture & ".bmp")
        frmMain.imgShopBuy.Picture = Nothing
        frmMain.imgShopSell.Picture = Nothing
        frmMain.imgShopBuy.Picture = LoadPicture(App.Path & "\graphics\gui\main\shop\" & Shop(Index).Picture & "buy.bmp")
        frmMain.imgShopSell.Picture = LoadPicture(App.Path & "\graphics\gui\main\shop\" & Shop(Index).Picture & "sell.bmp")
        frmMain.picShop.Visible = True
        frmMain.picShopItems.Cls
        
        For i = 1 To MAX_SHOP_ITEMS
            If Shop(Index).ShopItem(i).StockItem > 0 Then
                Itemnum = Shop(Index).ShopItem(i).StockItem
                If Itemnum > 0 And Itemnum <= MAX_ITEMS Then
                    Image = Item(Itemnum).Picture
                    
                    If Image = 0 Then Exit Sub
                    
                    If DDS_Item(Image) Is Nothing Then
                        Call InitDDSurf("Items\" & Image, DDSD_Item(Image), DDS_Item(Image))
                    End If
                    
                    With sRECT
                        .top = 0
                        .Bottom = .top + 32
                        .Left = (DDSD_Item(Image).lWidth / 3)
                        .Right = .Left + 32
                    End With
                    
                    With dRECT
                    .top = 5 + ((4 + 32) * ((i - 1) \ 5))
                    .Bottom = .top + 32
                    .Left = 10 + ((4 + 32) * (((i - 1) Mod 5)))
                    .Right = .Left + 32
                    End With
                    
                    BltToDC DDS_Item(Image), sRECT, dRECT, frmMain.picShopItems, False
                    
                    Amount = Shop(Index).ShopItem(i).StockValue
                    If Amount > 1 Then
                
                        If Amount < 1000000 Then
                            color = QBColor(White)
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            color = QBColor(Yellow)
                        ElseIf Amount > 10000000 Then
                            color = QBColor(BrightGreen)
                        End If
                        
                        Y = dRECT.top + 22
                        X = dRECT.Left - 4
                            
                        DrawText frmMain.picShopItems.hDC, X, Y, TransformAmount(Str(Amount)), color
                    End If
                    
                End If
            End If
        Next
        frmMain.picShopItems.Refresh
    End If
End Sub

Public Sub BltBars(ByVal Index As Long)
Dim Rec As RECT, LocRec As RECT
Dim Percentage As Single
Dim Length As Long
Dim Left As Long, top As Long

    '|| HEALTH ||'
    Percentage = Player(Index).Vital(Vitals.Health) / GetPlayerMaxVital(Index, Vitals.Health)
    Length = (DDSD_Bars.lWidth / 2) * Percentage
    top = (Player(Index).Y * 32) - 35 + TempPlayer(Index).YOffset - 8
    Left = (Player(Index).X * 32) + TempPlayer(Index).XOffset - 1
    
    With Rec
        .top = 0
        .Left = 0
        .Bottom = .top + 6
        .Right = 36
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    With Rec
        .top = 0
        .Left = 36
        .Bottom = .top + 6
        .Right = .Left + Length
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    '|| HEALTH ||'
    
    '|| SPIRIT ||'
    Percentage = Player(Index).Vital(Vitals.Spirit) / GetPlayerMaxVital(Index, Vitals.Spirit)
    Length = (DDSD_Bars.lWidth / 2) * Percentage
    top = top + 8
    
    With Rec
        .top = 6
        .Left = 0
        .Bottom = .top + 6
        .Right = 36
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    With Rec
        .top = 6
        .Left = 36
        .Bottom = .top + 6
        .Right = .Left + Length
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    '|| SPIRIT ||'
    
End Sub

Public Sub BltNpcBars(ByVal Index As Long)
Dim Rec As RECT, LocRec As RECT
Dim Percentage As Single
Dim Length As Long
Dim Left As Long, top As Long

Dim NpcNum As Long, Mapnum As Long

    Mapnum = Player(MyIndex).Map
    NpcNum = Map(Mapnum).MapNpc(Index).Num
    
    Percentage = Map(Mapnum).MapNpc(Index).Vital(Vitals.Health) / Npc(NpcNum).Vital(Vitals.Health)
    Length = (DDSD_Bars.lWidth / 2) * Percentage
    '|| HEALTH ||'
    top = (TempNpc(Mapnum).NpcNum(Index).Y * 32) - 35 + TempNpc(Mapnum).NpcNum(Index).YOffset - 8
    Left = (TempNpc(Mapnum).NpcNum(Index).X * 32) + TempNpc(Mapnum).NpcNum(Index).XOffset - 1
    
    With Rec
        .top = 0
        .Left = 0
        .Bottom = .top + 6
        .Right = 36
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    With Rec
        .top = 0
        .Left = 36
        .Bottom = .top + 6
        .Right = .Left + Length
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    '|| HEALTH ||'
    
    '|| SPIRIT ||'
    Percentage = Map(Mapnum).MapNpc(Index).Vital(Vitals.Spirit) / Npc(NpcNum).Vital(Vitals.Spirit)
    Length = (DDSD_Bars.lWidth / 2) * Percentage
    top = top + 8
    
    With Rec
        .top = 6
        .Left = 0
        .Bottom = .top + 6
        .Right = 36
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    With Rec
        .top = 6
        .Left = 36
        .Bottom = .top + 6
        .Right = .Left + Length
    End With
    
    Call BltFast(Left, top, DDS_Bars, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    '|| SPIRIT ||'
    
End Sub

Public Sub BltChest(ByVal ChestNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Opened As Boolean)
Dim Rec As RECT
Dim Picture As Long
Dim Left As Long, top As Long

    Picture = Chest(ChestNum).Picture
    
    With Rec
        .top = 1
        .Left = 0
        If Opened = True Then .Left = .Left + 32
        .Bottom = 21
        .Right = .Left + 32
    End With
    
    Left = X * 32
    top = (Y * 32) + 12
    
    Call BltFast(Left, top, DDS_Chest(Picture), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub

Public Sub RenderChest(ByVal ChestNum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long, Itemnum As Long, itemPic As Long
Dim Amount As Long
Dim itemPicRec As RECT, positionRec As RECT
Dim Left As Long, top As Long
Dim color As Long

    frmMain.picChest.Picture = Nothing
    frmMain.picChest.Picture = LoadPicture(App.Path & "\graphics\gui\main\chest.bmp")
    frmMain.picChest.Cls
    
    Game.ChestX = X
    Game.ChestY = Y
    
    For i = 1 To MAX_CHEST_ITEMS
        Itemnum = MapChest(Player(MyIndex).Map).Tile(X, Y).ChestItem(i).Itemnum
        
        If Itemnum > 0 Then
            itemPic = Item(Itemnum).Picture
            
            If itemPic > 0 And itemPic <= NumItems Then
                If DDSD_Item(itemPic).lWidth <= 96 Then
                
                    With itemPicRec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = .Left + 32
                    End With
                    
                    With positionRec
                        .Left = 32 + ((32) * (((i - 1) Mod 8)))
                        .top = (Int((i - 1) / 8) * 32) + 28 + (Int((i - 1) / 8) * 6)
                        .Bottom = .top + 32
                        .Right = .Left + 32
                    End With

                    
                    If DDS_Item(itemPic) Is Nothing Then
                        Call InitDDSurf("items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                    End If
                    
                    BltToDC DDS_Item(itemPic), itemPicRec, positionRec, frmMain.picChest, False

                    Amount = MapChest(ChestNum).Tile(X, Y).ChestItem(i).ItemValue
                    
                    If Amount < 1000000 Then
                        color = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        color = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        color = QBColor(BrightGreen)
                    End If
                    
                    If Amount > 1 Then
                        top = positionRec.top + 22
                        Left = positionRec.Left + 22
                        DrawText frmMain.picChest.hDC, Left, top, TransformAmount(Str(Amount)), color
                    End If
                    
                End If
            End If
        End If
    Next
    
    frmMain.picChest.Refresh
            
End Sub

Public Sub BltResource(ByVal Num As Long, ByVal X As Long, ByVal Y As Long)
Dim Rec As DxVBLib.RECT
Dim Mapnum As Long, Picture As Long
Dim Left As Long, top As Long

Mapnum = Player(MyIndex).Map

    With Resource(Num)
        Select Case MapResource(Mapnum).Tile(X, Y).Alive
            Case True
                Picture = .AliveGFX
            Case False
                Picture = .DeadGFX
        End Select
        If Picture > 0 Then
            With Rec
                .top = 0
                .Bottom = DDSD_Resource(Picture).lHeight
                .Left = 0
                .Right = DDSD_Resource(Picture).lWidth
            End With
            Left = (X * 32) - (DDSD_Resource(Picture).lWidth - 32) / 2
            top = (Y * 32) - (DDSD_Resource(Picture).lHeight - 32)
        
            Call BltFast(Left, top, DDS_Resource(Picture), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End With
    
End Sub

Public Sub BltTarget(ByVal NpcNum As Long)
Dim Rec As RECT
Dim Mapnum As Long
Dim Left As Long, top As Long

    Mapnum = Player(MyIndex).Map
    
    With Rec
        .top = 0
        .Bottom = DDSD_Target.lHeight / 2
        .Left = 0
        .Right = DDSD_Target.lWidth
    End With
    Left = (TempNpc(Mapnum).NpcNum(NpcNum).X * 32 + TempNpc(Mapnum).NpcNum(NpcNum).XOffset) - (DDSD_Target.lWidth - 32) / 2
    top = (TempNpc(Mapnum).NpcNum(NpcNum).Y * 32 + TempNpc(Mapnum).NpcNum(NpcNum).YOffset) - 16
    
    Call BltFast(Left, top, DDS_Target, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Public Sub BltProjectile(ByVal Num As Long)
Dim Rec As RECT
Dim Mapnum As Long
Dim Left As Long, top As Long
Dim Pic As Long

    Mapnum = Player(MyIndex).Map
    Pic = MapProjectile(Mapnum).MapProjectile(Num).Picture
    
    With Rec
        .top = 0
        .Left = MapProjectile(Mapnum).MapProjectile(Num).Dir * 32
        .Right = .Left + 32
        .Bottom = DDSD_Projectile(Pic).lHeight
    End With
    
    Left = ((MapProjectile(Mapnum).MapProjectile(Num).X * 32) - MapProjectile(Mapnum).MapProjectile(Num).XOffset)
    top = ((MapProjectile(Mapnum).MapProjectile(Num).Y * 32) - MapProjectile(Mapnum).MapProjectile(Num).YOffset) - (DDSD_Projectile(Pic).lHeight - 32)
    
    Call BltFast(Left, top, DDS_Projectile(Pic), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub


Public Sub BltHoverTarget(ByVal NpcNum As Long)
Dim Rec As RECT
Dim Mapnum As Long
Dim Left As Long, top As Long

    Mapnum = Player(MyIndex).Map
    
    With Rec
        .top = DDSD_Target.lHeight / 2
        .Bottom = DDSD_Target.lHeight
        .Left = 0
        .Right = DDSD_Target.lWidth
    End With
    Left = (TempNpc(Mapnum).NpcNum(NpcNum).X * 32 + TempNpc(Mapnum).NpcNum(NpcNum).XOffset) - (DDSD_Target.lWidth - 32) / 2
    top = (TempNpc(Mapnum).NpcNum(NpcNum).Y * 32 + TempNpc(Mapnum).NpcNum(NpcNum).YOffset) - 16
    
    Call BltFast(Left, top, DDS_Target, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Public Sub BltSpells()
Dim i As Long, X As Long, Y As Long, SpellNum As Long, SpellPic As Long
Dim Amount As Long
Dim Rec As RECT, PRec As RECT
Dim color As Long

    frmMain.picSpells.Cls
    
    For i = 1 To MAX_PLAYER_SPELLS
        SpellNum = Player(MyIndex).PlayerSpell(i).Num
        
        If SpellNum > 0 Then
            SpellPic = Spell(SpellNum).Picture
            
            If SpellPic > 0 And SpellPic <= NumSpells Then
                'If DDSD_Spell(SpellPic).lWidth <= 96 Then
                
                    If Player(MyIndex).PlayerSpell(i).CoolDownTimer > 0 Then
                        With Rec
                            .top = 0
                            .Bottom = 32
                            .Left = 32
                            .Right = 64
                        End With
                    Else
                        With Rec
                            .top = 0
                            .Bottom = 32
                            .Left = 0
                            .Right = 32
                        End With
                    End If
                    
                    With PRec
                        .top = 1 + (35 * ((i - 1) \ 5))
                        .Bottom = .top + 32
                        .Left = 12 + (35 * (((i - 1) Mod 5)))
                        .Right = .Left + 32
                    End With
                    
                    If DDS_Spell(SpellPic) Is Nothing Then
                        Call InitDDSurf("spells\" & SpellPic, DDSD_Spell(SpellPic), DDS_Spell(SpellPic))
                    End If
                    
                    BltToDC DDS_Spell(SpellPic), Rec, PRec, frmMain.picSpells, False
                    
                'End If
            End If
        End If
    Next
    
    frmMain.picSpells.Refresh
            
End Sub
