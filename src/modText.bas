Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Function GetFontNumber(ByVal name As String) As Byte
    Select Case name
        Case "Calibri"
            GetFontNumber = 1
        Case "Cambria"
            GetFontNumber = 2
        Case "Candara"
            GetFontNumber = 3
        Case "Courier New"
            GetFontNumber = 4
        Case "News Gothic"
            GetFontNumber = 5
        Case "Palantino Linotype"
            GetFontNumber = 6
        Case "Pescadero"
            GetFontNumber = 7
        Case "Tahoma"
            GetFontNumber = 8
        Case "Trajan Pro"
            GetFontNumber = 9
        Case "Trebuchet MS"
            GetFontNumber = 10
    End Select
            
End Function

Public Sub LoadFonts()

FontStyle(1) = "Calibri"
FontStyle(2) = "Cambria"
FontStyle(3) = "Candara"
FontStyle(4) = "Courier New"
FontStyle(5) = "News Gothic"
FontStyle(6) = "Palatino Linotype"
FontStyle(7) = "Pescadero"
FontStyle(8) = "Tahoma"
FontStyle(9) = "Trajan Pro"
FontStyle(10) = "Trebuchet MS"

End Sub

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
    frmMain.Font = Font
    frmMain.FontSize = Size - 5
    
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
    
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    
End Sub

Public Sub DrawCoords()
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim text As String

    ' Set the color
    color = QBColor(Yellow)
    ' Set the text you want to render
    
    ' calc pos
    TextX = 32
    TextY = 32

    ' Draw Stuff
    text = "X: " & Player(MyIndex).X & "  XOffSet: " & TempPlayer(MyIndex).XOffset
    Call DrawText(TexthDC, TextX, TextY, text, color)
    text = "Y: " & Player(MyIndex).Y & "  YOffSet: " & TempPlayer(MyIndex).YOffset
    Call DrawText(TexthDC, TextX, TextY + 14, text, color)
    text = "Map: " & Player(MyIndex).Map
    Call DrawText(TexthDC, TextX, TextY + 26, text, color)
    text = "Dir: " & Player(MyIndex).Dir
    Call DrawText(TexthDC, TextX, TextY + 40, text, color)
    
End Sub

Public Sub DrawMapName()
Dim color As Byte
    
    Select Case Map(Player(MyIndex).Map).Moral
        Case MAP_MORAL_NONE
            color = White
        Case MAP_MORAL_DUNGEON
            color = Yellow
    End Select
    
    Call DrawText(TexthDC, (frmMain.picScreen.Width / 2) - Len(Trim$(Map(Player(MyIndex).Map).name)) / 2, 40, (Trim$(Map(Player(MyIndex).Map).name)), QBColor(color))
End Sub

Public Sub DrawMapTileAttribute(ByVal X As Long, ByVal Y As Long)
Dim color As Byte
Dim text As String

    Select Case Map(Player(MyIndex).Map).Tile(X, Y).Attribute
        Case 0
            Exit Sub
        Case Attributes.BlockedTile ' blocked
            text = "B"
            color = BrightRed
        Case Attributes.warptile ' Warp
            text = "W"
            color = BrightBlue
        Case Attributes.Soundtile
            text = "S"
            color = Cyan
        Case Attributes.ItemTile
            text = "I"
            color = White
        Case Attributes.HealTile
            text = "H"
            color = BrightGreen
        Case Attributes.TrapTile
            text = "T"
            color = BrightRed
        Case Attributes.NpcSpawnTile
            text = "N"
            color = Yellow
        Case Attributes.NpcAvoidTile
            text = "A"
            color = White
        Case Attributes.KeyTile
            text = "K"
            color = White
        Case Attributes.ResourceTile
            text = "R"
            color = Green
        Case Attributes.BankTile
            text = "B"
            color = Blue
        Case Attributes.ShopTile
            text = "S"
            color = BrightBlue
        Case Attributes.ChestTile
            text = "C"
            color = White
        End Select
    
    Call DrawText(TexthDC, (X * 32) + 12, (Y * 32) + 8, text, QBColor(color))

End Sub

Public Sub AddText(ByVal Msg As String, ByVal color As Integer, Optional NewLine As Boolean = True)
Dim s As String
    
    If NewLine = True Then
        s = vbNewLine & Msg
    Else
        s = Msg
    End If
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(color)
    frmMain.txtChat.SelText = s
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1

End Sub

Public Sub DrawPlayerName(ByVal Index As Long, ByVal spriteheight As Long)
Dim color As Byte

    Select Case Player(Index).Access
        Case ACCESS_PLAYER
            color = Grey
        Case ACCESS_MEMBER
            color = White
        Case ACCESS_MODERATOR
            color = Cyan
        Case ACCESS_ADMIN
            color = Red
        Case ACCESS_OWNER
            color = Yellow
    End Select
    
    With Player(Index)
        Call DrawText(TexthDC, (.X * 32) - (Len("Level: " & Player(Index).Combat.level) / 1.05) + TempPlayer(Index).XOffset, .Y * 32 - spriteheight + TempPlayer(Index).YOffset + 32, "Level: " & Player(Index).Combat.level, QBColor(color))
        Call DrawText(TexthDC, .X * 32 - Len(Trim$(Player(Index).name)) + TempPlayer(Index).XOffset, .Y * 32 + 16 - spriteheight + TempPlayer(Index).YOffset, Trim$(Player(Index).name), QBColor(color))
    End With
    
End Sub

Public Sub DrawNpcName(ByVal Index As Long, ByVal spriteheight As Long)
Dim color As Byte
Dim NpcNum As Long
Dim Left As Long, top As Long

    NpcNum = Map(Player(MyIndex).Map).MapNpc(Index).Num

    Select Case Npc(NpcNum).Type
        Case NPC_TYPE_FRIENDLY
            color = BrightGreen
        Case NPC_TYPE_STATIONARY
            color = BrightGreen
        Case NPC_TYPE_ATTACK_WHEN_ATTACKED
            color = Yellow
        Case NPC_TYPE_ATTACK_ON_SIGHT
            color = BrightRed
    End Select
    
    With TempNpc(Player(MyIndex).Map).NpcNum(Index)
        Left = .X * 32 - Len(Trim$(Npc(NpcNum).name)) + .XOffset
        top = .Y * 32 + 16 - spriteheight + .YOffset
        Call DrawText(TexthDC, Left, top - 15, Trim$(Npc(NpcNum).name), QBColor(color))
        Call DrawText(TexthDC, Left, top, "Level: " & "0", QBColor(color))
    End With
End Sub

Sub DrawActionMsg(ByVal Index As Long)
Dim time As Long
    
    ' does it exist
    If TempActionMsg(Index).Created = 0 Then Exit Sub
    
        time = 900
        
        TempActionMsg(Index).Y = TempActionMsg(Index).Y - 0.25

    If timeGetTime < TempActionMsg(Index).Created + time Then
        Call DrawText(TexthDC, TempActionMsg(Index).X, TempActionMsg(Index).Y, TempActionMsg(Index).message, QBColor(TempActionMsg(Index).color))
    Else
        ClearActionMsg Index
    End If
End Sub
