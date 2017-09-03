VERSION 5.00
Begin VB.Form frmEditor_Map 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMapOptions 
      Caption         =   "Map Properties"
      Height          =   6735
      Left            =   8640
      TabIndex        =   31
      Top             =   120
      Width           =   2895
      Begin VB.Frame Frame2 
         Caption         =   "Npc's"
         Height          =   2895
         Left            =   120
         TabIndex        =   54
         Top             =   2640
         Width           =   2655
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   1440
            TabIndex        =   58
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   2400
            Width           =   1095
         End
         Begin VB.ListBox lstNpcs 
            Height          =   2010
            Left            =   1440
            TabIndex        =   56
            Top             =   240
            Width           =   1095
         End
         Begin VB.ListBox lstMapNpcs 
            Height          =   2010
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtRightWarp 
         Height          =   285
         Left            =   1320
         TabIndex        =   51
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtDownWarp 
         Height          =   285
         Left            =   840
         TabIndex        =   50
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtLeftWarp 
         Height          =   285
         Left            =   360
         TabIndex        =   49
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtUpWarp 
         Height          =   285
         Left            =   840
         TabIndex        =   48
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox cmdMoral 
         Height          =   315
         ItemData        =   "frmEditor_Map.frx":0000
         Left            =   720
         List            =   "frmEditor_Map.frx":000A
         TabIndex        =   47
         Text            =   "None"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtMusic 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Moral:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Music:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame fraAttrExtra 
      BorderStyle     =   0  'None
      Caption         =   "Click on a label to find value"
      Height          =   6495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtString 
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   42
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtString 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   41
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtString 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   40
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtString 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.HScrollBar scrlLong 
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   27
         Top             =   3720
         Value           =   1
         Width           =   2055
      End
      Begin VB.HScrollBar scrlLong 
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   25
         Top             =   3120
         Value           =   1
         Width           =   2055
      End
      Begin VB.HScrollBar scrlLong 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   23
         Top             =   2520
         Value           =   1
         Width           =   2055
      End
      Begin VB.HScrollBar scrlLong 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   21
         Top             =   1920
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label lblString 
         Caption         =   "String Value"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   38
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblString 
         Caption         =   "String Value"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   37
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblString 
         Caption         =   "String Value"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblString 
         Caption         =   "String Value"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   35
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   1200
         TabIndex        =   28
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label lblLong 
         Caption         =   "Long Value: 1"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   26
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label lblLong 
         Caption         =   "Long Value: 1"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   24
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblLong 
         Caption         =   "Long Value: 1"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   22
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblLong 
         Caption         =   "Long Value: 1"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   6600
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdCopyMap 
         Caption         =   "Copy Map"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Frame fraAttributes 
         Caption         =   "Attributes"
         Height          =   4575
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.OptionButton optAttribute 
            Caption         =   "Chest"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   65
            Top             =   3240
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Shop"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   64
            Top             =   3000
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Bank"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   63
            Top             =   2760
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Resource"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   62
            Top             =   2520
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Key"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   61
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "NpcAvoid"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   60
            Top             =   2040
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Npc"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   59
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Trap"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   53
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Heal"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Item"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Sound"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Warp"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optAttribute 
            Caption         =   "Blocked"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Attributes"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   5040
         Width           =   1575
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Layers"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame fraLayers 
         Caption         =   "Layers"
         Height          =   4575
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1935
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Layer"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   4080
            Width           =   1695
         End
         Begin VB.CommandButton cmdFill 
            Caption         =   "Fill Layer"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3720
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Ground"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Mask - 1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Mask - 2"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Mask - 3"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Mask - Anim"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Fringe - 1"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   2160
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Fringe - 2"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Fringe - 3"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   6
            Top             =   2880
            Width           =   1695
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Fringe - Anim"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   5
            Top             =   3240
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox picTileset 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   1
      Top             =   0
      Width           =   6495
      Begin VB.PictureBox picSel 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6000
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   6000
         Width           =   480
      End
   End
   Begin VB.HScrollBar scrlTileset 
      Height          =   255
      Left            =   0
      Min             =   1
      TabIndex        =   0
      Top             =   6600
      Value           =   1
      Width           =   6495
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim I As Long
Dim Rem1 As Long, Rem2 As Long

    Map(Player(MyIndex).Map).MapNpc(lstMapNpcs.ListIndex + 1).Num = lstNpcs.ListIndex + 1
    
    Rem1 = lstMapNpcs.ListIndex
    Rem2 = lstNpcs.ListIndex

        lstMapNpcs.Clear
        lstNpcs.Clear
        For I = 1 To MAX_MAP_NPCS
            If Map(Player(MyIndex).Map).MapNpc(I).Num <> 0 Then
                lstMapNpcs.AddItem (I & ": " & Trim$(Npc(Map(Player(MyIndex).Map).MapNpc(I).Num).name))
            Else
                lstMapNpcs.AddItem (I & ": ")
            End If
        Next
        For I = 1 To MAX_NPCS
            lstNpcs.AddItem (I & ": " & Trim$(Npc(I).name))
        Next
        lstNpcs.ListIndex = Rem2
        lstMapNpcs.ListIndex = Rem1

End Sub

Private Sub cmdClear_Click()
Dim X As Long, Y As Long

    For X = 1 To MAX_MAP_X
        For Y = 1 To MAX_MAP_Y
            With Map(Player(MyIndex).Map).Tile(X, Y).Layer(EditingLayer)
                .Tileset = 0
                .X = 0
                .Y = 0
            End With
        Next
    Next
    
End Sub

Private Sub cmdCopyMap_Click()
Dim Mapnum As String
Dim X As Long, Y As Long, L As Long

    Mapnum = InputBox("Input the map index")
    If IsNumeric(Mapnum) Then
        If Mapnum <> 0 And Mapnum < MAX_MAPS + 1 Then
            Call ClearMap(Player(MyIndex).Map)
            With Map(Player(MyIndex).Map)
                .DownWarp = Map(Mapnum).DownWarp
                .UpWarp = Map(Mapnum).UpWarp
                .LeftWarp = Map(Mapnum).LeftWarp
                .RightWarp = Map(Mapnum).RightWarp
                .Moral = Map(Mapnum).Moral
                .Music = Map(Mapnum).Music
                .name = Trim$(Map(Mapnum).name)
                For L = 1 To MAX_MAP_NPCS
                    If Map(Mapnum).MapNpc(L).Num > 0 Then
                        .MapNpc(L).Dir = Map(Mapnum).MapNpc(L).Dir
                        .MapNpc(L).Num = Map(Mapnum).MapNpc(L).Num
                        .MapNpc(L).SpawnX = Map(Mapnum).MapNpc(L).SpawnX
                        .MapNpc(L).SpawnY = Map(Mapnum).MapNpc(L).SpawnY
                        .MapNpc(L).Vital(Vitals.Health) = Npc(Map(Mapnum).MapNpc(L).Num).Vital(Vitals.Health)
                        .MapNpc(L).Vital(Vitals.Spirit) = Npc(Map(Mapnum).MapNpc(L).Num).Vital(Vitals.Spirit)
                        With TempNpc(Player(MyIndex).Map).NpcNum(L)
                            .Alive = True
                            .Target = 0
                        End With
                    End If
                Next
                For X = 1 To MAX_MAP_X
                    For Y = 1 To MAX_MAP_Y
                        With .Tile(X, Y)
                            .Attribute = Map(Mapnum).Tile(X, Y).Attribute
                            For L = 1 To Layers.Layer_Count - 1
                                .Layer(L).Tileset = Map(Mapnum).Tile(X, Y).Layer(L).Tileset
                                .Layer(L).X = Map(Mapnum).Tile(X, Y).Layer(L).X
                                .Layer(L).Y = Map(Mapnum).Tile(X, Y).Layer(L).Y
                            Next
                            For L = 1 To 4
                                .LongValue(L) = Map(Mapnum).Tile(X, Y).LongValue(L)
                                .StringValue(L) = Trim$(Map(Mapnum).Tile(X, Y).StringValue(L))
                            Next
                        End With
                        
                        With MapResource(Mapnum).Tile(X, Y)
                            If MapResource(Mapnum).Tile(X, Y).Num > 0 Then
                                .Alive = True
                                .Health = Resource(.Num).Health
                                .Num = MapResource(Mapnum).Tile(X, Y).Num
                            End If
                        End With
                    Next
                Next
            End With
            
            ' Respawn the map
            Call SaveMap(Player(MyIndex).Map)
            Call LoadMap(Player(MyIndex).Map)
            
        End If
    End If

End Sub

Private Sub cmdFill_Click()
Dim X As Long, Y As Long

    For X = 1 To MAX_MAP_X
        For Y = 1 To MAX_MAP_Y
            With Map(Player(MyIndex).Map).Tile(X, Y).Layer(EditingLayer)
                .Tileset = scrlTileset.Value
                .X = CurTileX
                .Y = CurTileY
            End With
        Next
    Next

End Sub

Private Sub cmdMoral_Click()

    Map(Player(MyIndex).Map).Moral = cmdMoral.ListIndex

End Sub

Private Sub cmdRemove_Click()
Dim I As Long
Dim Rem1 As Long, Rem2 As Long

    Map(Player(MyIndex).Map).MapNpc(lstMapNpcs.ListIndex + 1).Num = 0
    Map(Player(MyIndex).Map).MapNpc(lstMapNpcs.ListIndex + 1).SpawnX = 0
    Map(Player(MyIndex).Map).MapNpc(lstMapNpcs.ListIndex + 1).SpawnY = 0
    With TempNpc(Player(MyIndex).Map).NpcNum(lstMapNpcs.ListIndex + 1)
        .X = 0
        .Y = 0
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
        .Step = 0
    End With
    
    Rem1 = lstMapNpcs.ListIndex
    Rem2 = lstNpcs.ListIndex
    
        lstMapNpcs.Clear
        lstNpcs.Clear
        For I = 1 To MAX_MAP_NPCS
            If Map(Player(MyIndex).Map).MapNpc(I).Num <> 0 Then
                lstMapNpcs.AddItem (I & ": " & Trim$(Npc(Map(Player(MyIndex).Map).MapNpc(I).Num).name))
            Else
                lstMapNpcs.AddItem (I & ": ")
            End If
        Next
        For I = 1 To MAX_NPCS
            lstNpcs.AddItem (I & ": " & Trim$(Npc(I).name))
        Next
        lstNpcs.ListIndex = Rem2
        lstMapNpcs.ListIndex = Rem1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Unload(Me)
    Call SaveMap(Player(MyIndex).Map)
End Sub

Private Sub lblLong_Click(Index As Integer)

    lblValue.Caption = "Null"

    Select Case EditingAttribute
        Case Attributes.warptile
            Select Case Index
                Case 1 ' Map num
                    lblValue.Caption = "This changes the map number you warp to."
                Case 2 ' X
                    lblValue.Caption = "This changes the x you warp to."
                Case 3 ' Y
                    lblValue.Caption = "This changes the y you warp to."
                Case 4 ' Dir
                    lblValue.Caption = "This changes the dir you will have."
            End Select
        Case Attributes.Soundtile
            Select Case Index
                Case 1 ' YTileRange
                    lblValue.Caption = "This changes the YTileRange."
                Case 2 ' XTileRange
                    lblValue.Caption = "This changes the XTileRange."
                Case 3 ' Chance
                    lblValue.Caption = "This changes the chance (1/value) that the sound is played."
                Case 4 ' SoundLength
                    lblValue.Caption = "This value should be the seconds the sound is played for."
            End Select
        Case Attributes.ItemTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the item index."
                Case 2
                    lblValue.Caption = "This changes the item amount value."
                Case 3
                    lblValue.Caption = "This changes the seconds til the item respawns."
            End Select
        Case Attributes.HealTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the amount that's healed when you walk on it."
            End Select
        Case Attributes.TrapTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the damage that's taken when you walk on it."
            End Select
        Case Attributes.NpcSpawnTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the dir that the npc faces in. Default = 1 [1:D 2:L 3:R 4:U]"
            End Select
        Case Attributes.KeyTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the item index needed to unlock the tile."
            End Select
        Case Attributes.ShopTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the shop index"
            End Select
        Case Attributes.ChestTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the chest index."
            End Select
        Case Attributes.ResourceTile
            Select Case Index
                Case 1
                    lblValue.Caption = "This changes the resource index."
            End Select
    End Select
    
End Sub

Private Sub lblString_Click(Index As Integer)

    lblValue.Caption = "Null"
    
    Select Case EditingAttribute
        Case Attributes.Soundtile
            Select Case Index
                Case 1 ' Sound name
                    lblValue.Caption = "This should be the name of the sound (.wav)"
            End Select
    End Select
    
End Sub

Private Sub optAttribute_Click(Index As Integer)

    EditingAttribute = Index
    
    fraAttrExtra.Visible = False
    
    Select Case Index
        Case Attributes.warptile
            fraAttrExtra.Visible = True
        Case Attributes.Soundtile
            fraAttrExtra.Visible = True
        Case Attributes.ItemTile
            fraAttrExtra.Visible = True
        Case Attributes.HealTile
            fraAttrExtra.Visible = True
        Case Attributes.TrapTile
            fraAttrExtra.Visible = True
        Case Attributes.NpcSpawnTile
            fraAttrExtra.Visible = True
        Case Attributes.KeyTile
            fraAttrExtra.Visible = True
        Case Attributes.ShopTile
            fraAttrExtra.Visible = True
        Case Attributes.ChestTile
            fraAttrExtra.Visible = True
        Case Attributes.ResourceTile
            fraAttrExtra.Visible = True
    End Select

End Sub

Private Sub optEdit_Click(Index As Integer)

    CurrentlyEditing = Index
    
    fraLayers.Visible = False
    fraAttributes.Visible = False
    
    Select Case Index
        Case EDITING_LAYERS
            fraLayers.Visible = True
            fraAttrExtra.Visible = False
        Case EDITING_ATTRIBUTES
            fraAttributes.Visible = True
    End Select
            

End Sub

Private Sub optLayer_Click(Index As Integer)
    EditingLayer = Index
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempX As Single, TempY As Single

    
    ' Convert the number, and make sure there's no rounding up or down.
    TempX = (X / 32)
    TempY = (Y / 32)
    
    
    CurTileX = TempX
    CurTileY = TempY
    If TempX - CurTileX < 0 Then CurTileX = CurTileX - 1
    If TempY - CurTileY < 0 Then CurTileY = CurTileY - 1
    
    picSel.Left = CurTileX * 32
    picSel.top = CurTileY * 32
    
    frmEditor_Map.Caption = CurTileX & ":" & CurTileY

End Sub

Private Sub scrlLong_Change(Index As Integer)

    lblLong(Index).Caption = "LongValue: " & scrlLong(Index).Value

End Sub

Private Sub scrlTileset_Change()

    picTileset.Picture = Nothing
    picTileset.Picture = LoadPicture(App.Path & "\graphics\tilesets\" & scrlTileset.Value & ".bmp")

End Sub

Private Sub txtDownWarp_Change()

    If IsNumeric(txtDownWarp.text) = False Then
        txtDownWarp.text = Map(Player(MyIndex).Map).DownWarp
    End If
    
    Map(Player(MyIndex).Map).DownWarp = txtDownWarp.text

End Sub

Private Sub txtLeftWarp_Change()

    If IsNumeric(txtLeftWarp.text) = False Then
        txtLeftWarp.text = Map(Player(MyIndex).Map).LeftWarp
    End If

    Map(Player(MyIndex).Map).LeftWarp = txtLeftWarp.text

End Sub

Private Sub txtMusic_Change()

    Map(Player(MyIndex).Map).Music = txtMusic.text

End Sub

Private Sub txtName_Change()

    Map(Player(MyIndex).Map).name = txtName.text

End Sub

Private Sub txtRightWarp_Change()

    If IsNumeric(txtRightWarp.text) = False Then
        txtRightWarp.text = Map(Player(MyIndex).Map).RightWarp
    End If
    
    Map(Player(MyIndex).Map).RightWarp = txtRightWarp.text

End Sub

Private Sub txtUpWarp_Change()

    If IsNumeric(Player(MyIndex).Map) = False Then
        txtUpWarp.text = Map(Player(MyIndex).Map).UpWarp
    End If
    
    Map(Player(MyIndex).Map).UpWarp = txtUpWarp.text

End Sub
