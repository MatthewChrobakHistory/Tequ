Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SMsgBox) = GetAddress(AddressOf HandleMsgBox)
    HandleDataSub(SEnterGame) = GetAddress(AddressOf HandleEnterGame)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SClientAddText) = GetAddress(AddressOf HandleClientAddText)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandleMovePlayer)
    HandleDataSub(SCanStop) = GetAddress(AddressOf HandleCanStop)
    HandleDataSub(SDropItem) = GetAddress(AddressOf HandleDropItem)
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
    
End Sub

Public Sub HandleEnterGame(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MyIndex = Buffer.ReadLong
    frmMenu.tmrEnterGame.Enabled = True
    Set Buffer = Nothing
    
End Sub

Public Sub HandlePlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PlayerIndex As Long, X As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerIndex = Buffer.ReadLong
    Player(PlayerIndex).name = Buffer.ReadString
    Player(PlayerIndex).Combat.level = Buffer.ReadLong
    Player(PlayerIndex).Combat.XP = Buffer.ReadLong
    Player(PlayerIndex).Points = Buffer.ReadLong
    'Player(PlayerIndex).Sprite = Buffer.ReadLong
    Player(PlayerIndex).Map = Buffer.ReadLong
    Player(PlayerIndex).X = Buffer.ReadLong
    Player(PlayerIndex).Y = Buffer.ReadLong
    Player(PlayerIndex).Dir = Buffer.ReadLong
    Player(PlayerIndex).Access = Buffer.ReadLong
    Set Buffer = Nothing
    ' Check if the player is the client player. If so, update relevant stuff.
    If PlayerIndex = MyIndex Then
    End If

End Sub

Public Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim MapIndex As Long, X As Long, Y As Long, L As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapIndex = Buffer.ReadLong
    Map(MapIndex).name = Buffer.ReadString
    For X = 1 To MAX_MAP_X
        For Y = 1 To MAX_MAP_Y
            For L = 1 To Layers.Layer_Count - 1
                Map(MapIndex).Tile(X, Y).Layer(L).Tileset = Buffer.ReadLong
                Map(MapIndex).Tile(X, Y).Layer(L).X = Buffer.ReadLong
                Map(MapIndex).Tile(X, Y).Layer(L).Y = Buffer.ReadLong
            Next
        Next
    Next
    Set Buffer = Nothing

End Sub

Public Sub HandleMsgBox(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call MsgBox(Buffer.ReadString)
    Set Buffer = Nothing
    
End Sub

Public Sub HandleClientAddText(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Color As Byte, text As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    text = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(text, Color)
    Set Buffer = Nothing
    
End Sub

Public Sub HandleMovePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PIndex As Long, Dir As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PIndex = Buffer.ReadLong
    Player(PIndex).Dir = Buffer.ReadLong
    Player(PIndex).X = Buffer.ReadLong
    Player(PIndex).Y = Buffer.ReadLong
    TempPlayer(PIndex).Moving = Player(PIndex).Dir
    Call SetOffset(PIndex, Player(PIndex).Dir)
    Set Buffer = Nothing
    
End Sub

Public Sub HandleCanStop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PIndex As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PIndex = Buffer.ReadLong
    TempPlayer(PIndex).X = Buffer.ReadLong
    TempPlayer(PIndex).Y = Buffer.ReadLong
    TempPlayer(PIndex).CanStop = True
    Set Buffer = Nothing
End Sub

Public Sub HandleMapLayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, L As Long
Dim MinX As Long, MaxX As Long
Dim MinY As Long, MaxY As Long
Dim MapNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MaxY = Buffer.ReadLong
    MinY = Buffer.ReadLong
    MaxX = Buffer.ReadLong
    MinX = Buffer.ReadLong
    MapNum = Buffer.ReadLong
    For X = MinX To MaxX
        For Y = MinY To MaxY
            For L = 1 To Layers.Layer_Count - 1
                With Map(MapNum).Tile(X, Y)
                    .Layer(L).Tileset = Buffer.ReadLong
                    .Layer(L).X = Buffer.ReadLong
                    .Layer(L).Y = Buffer.ReadLong
                End With
            Next
        Next
    Next
    Set Buffer = Nothing
End Sub

Public Sub HandleDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim DataIndex As Long, itemNum As Long, Slot As Long, MapNum As Long, L As Long, X As Long, Y As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    DataIndex = Buffer.ReadLong
    itemNum = Buffer.ReadLong
    Slot = Buffer.ReadLong
    MapNum = Buffer.ReadLong
    L = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    
    If DataIndex = MyIndex Then
        Call DropItem(Slot)
    Else
        With MapItem(MapNum).Tile(X, Y).Layer(L)
        
        End With
    End If
End Sub
