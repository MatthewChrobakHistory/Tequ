Attribute VB_Name = "modInput"
Option Explicit

Public Sub DisableKeys()

    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False

End Sub

Public Sub CheckKeys()
Dim PressedKey As Integer

    PressedKey = 0

    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
        PressedKey = vbKeyShift
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
        PressedKey = vbKeyControl
    Else
        ControlDown = False
    End If

    'Move Up
    If GetKeyState(vbKeyUp) < 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        PressedKey = vbKeyUp
    ElseIf GetKeyState(vbKeyRight) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        PressedKey = vbKeyRight
    ElseIf GetKeyState(vbKeyDown) < 0 Then
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        PressedKey = vbKeyDown
    ElseIf GetKeyState(vbKeyLeft) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        PressedKey = vbKeyLeft
    End If
    
    If PressedKey > 0 Then Call HandleKeyPresses(PressedKey)
    
End Sub

Public Sub HandleKeyPresses(ByVal KeyCode As Integer)
Dim ChatText As String
Dim TempMult As String
Dim X As Long

    If CreatingCharacter = True Then Exit Sub

    ' Non chat related stuff
    If ChatFocus = False Then
    
        If Game.InBank = True Then
            X = WIMultiplier
            Select Case KeyCode
                Case vbKey1
                    WIMultiplier = 1
                Case vbKey2
                    WIMultiplier = 5
                Case vbKey3
                    WIMultiplier = 10
                Case vbKey4
                    WIMultiplier = 2147483647
                Case vbKey5
                    TempMult = InputBox("Insert a custom multiplier.", "Custom Multiplier")
                    If IsNumeric(TempMult) Then
                        If TempMult > 2147483647 Then TempMult = 2147483647
                        WIMultiplier = TempMult
                    Else
                        WIMultiplier = 1
                    End If
            End Select
            If X <> WIMultiplier Then RenderBank
        End If
        
        ' The player isn't moving
        With TempPlayer(MyIndex)
            If .Moving = 0 Then
                Select Case KeyCode
                    Case vbKeyShift
                        '.Running = True
                    Case vbKeyDown
                        .Moving = DIR_DOWN
                        Player(MyIndex).Dir = .Moving
                    Case vbKeyUp
                        .Moving = DIR_UP
                        Player(MyIndex).Dir = .Moving
                    Case vbKeyRight
                        .Moving = DIR_RIGHT
                        Player(MyIndex).Dir = .Moving
                    Case vbKeyLeft
                        .Moving = DIR_LEFT
                        Player(MyIndex).Dir = .Moving
                End Select
                If .Moving > 0 Then
                    If Options.OnlineMode = True Then
                        Call SendPlayerMove
                    Else
                        Call InitiateMovement
                    End If
                End If
            End If
        End With
            
        Select Case KeyCode
            Case vbKeyEscape
                ' Admin panel
                If Player(MyIndex).Access <> ACCESS_ADMIN Then
                    If InputBox("In order to proceed, you must enter the admin key.", "Admin Validation") = AdminPassword Then
                        IsLegitAdmin = True
                        Call frmAdminPanel.PanelInit
                        Player(MyIndex).Access = ACCESS_ADMIN
                    End If
                Else
                    frmAdminPanel.PanelInit
                End If
        End Select
    End If

    If KeyCode = vbKeySpace Then
        If ChatFocus = False Then
            ChatFocus = True
            frmMain.txtMyChat.Visible = True
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If ChatFocus = True Then
            ChatText = Trim$(frmMain.txtMyChat.text)
            If Len(ChatText) > 0 Then
                Call SayMsg(ChatText)
                frmMain.txtMyChat.text = vbNullString
            Else
                frmMain.txtMyChat.Visible = False
                ChatFocus = False
            End If
        Else
            If Options.OnlineMode = True Then
            Else
                Call PickUpMapItem
            End If
        End If
    End If
    
    If KeyCode = vbKeyControl Then
        If Options.OnlineMode = True Then
        Else
            If TempPlayer(MyIndex).AttackTimer = 0 Then
                Call PlayerAction
            End If
        End If
    End If

End Sub

Public Sub HandleKeyReleases(ByVal KeyCode As Integer)

    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Call SendPlayerStop
    End If
    
    If KeyCode = vbKeyShift Then
        'TempPlayer(MyIndex).Running = False
    End If

End Sub

Public Sub SayMsg(ByVal text As String)
Dim color As Byte

    If CheckForCommands(text) Then Exit Sub
    
    If Options.OnlineMode = True Then
        Call SendServerMessage(MyIndex, text)
        Exit Sub
    End If

    Select Case Player(MyIndex).Access
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
    
    Call AddText(Trim$(Player(MyIndex).Name) & ": " & text, color)
End Sub

Public Function CheckForCommands(ByVal text As String) As Boolean

    CheckForCommands = True
    
    If text = "/loc" Or text = "/Loc" Then
        If TempPlayer(MyIndex).DrawCoords = False Then
            TempPlayer(MyIndex).DrawCoords = True
            Exit Function
        Else
            TempPlayer(MyIndex).DrawCoords = False
            Exit Function
        End If
    End If
    
    If text = "/cps" Or text = "/Cps" Then
        If TempPlayer(MyIndex).DrawCPS = False Then
            TempPlayer(MyIndex).DrawCPS = True
            Exit Function
        Else
            TempPlayer(MyIndex).DrawCPS = False
            Exit Function
        End If
    End If
    
    If text = "/debug" Or text = "/Debug" Then
        'Call InitDebugGUI
        Exit Function
    End If
    
    CheckForCommands = False

End Function
