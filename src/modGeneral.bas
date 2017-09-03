Attribute VB_Name = "modGeneral"
Option Explicit

' Main timer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Halts thread of execution
Public Declare Sub Sleep Lib "Kernel32" (ByVal wMilliseconds As Long)

' API declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' Clearing UDT's
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Shell Executing for opening up files
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Loading and Saving Text
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' Figuring out whether or not the forum is the focus
Public Declare Function GetActiveWindow Lib "user32" () As Long

' Keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Master Object
Public DX7 As New DirectX7
Private Form As ClsFormBorder ' Form Stuff

'///////////////////////////////////////////////////////////////////
'/////////////////////// GAME FUNCTIONS ///////////////////////////
'///////////////////////////////////////////////////////////////////

Sub Main()
Dim Answer As Integer
    
On Error GoTo errorhandler ' We haven't loaded the options yet, so there can't be the check yet.

    Set Form = New ClsFormBorder
    Set Form.Client = frmMain

    Call LoadFonts
    Call LoadOptions
    Call InitSound
    Call PlayMusic("Theme.mp3")
    
    Norm(1) = App.Path & "\graphics\gui\menu\buttons\singleplayer.bmp"
    Norm(2) = App.Path & "\graphics\gui\menu\buttons\multiplayer.bmp"
    Norm(3) = App.Path & "\graphics\gui\menu\buttons\options.bmp"
    Norm(4) = App.Path & "\graphics\gui\menu\buttons\exit.bmp"
    Hover(1) = App.Path & "\graphics\gui\menu\buttons\singleplayer_hover.bmp"
    Hover(2) = App.Path & "\graphics\gui\menu\buttons\multiplayer_hover.bmp"
    Hover(3) = App.Path & "\graphics\gui\menu\buttons\options_hover.bmp"
    Hover(4) = App.Path & "\graphics\gui\menu\buttons\exit_hover.bmp"
    
    Call SetupMenuGui
    
    frmMenu.Visible = True
    
    If Options.InstallRuntimes = True Then
       Answer = MsgBox("Is this your first time using a Tequ program? If so, you'll need to download and register some DLL's." & vbCrLf & vbCrLf & "Don't worry! You can use the Origins Runtimes installer to install and register all the necessary files. Would you like to run the installer?", vbYesNo, "DLL Installer Prompt")
       If Answer = vbYes Then
            MsgBox "If the game doesn't work after installing the runtimes, then run the Runtimes.exe file located in the same folder as the client.", , "DLL Installer Prompt"
            Call Shell(App.Path & "\runtimes.exe", vbNormalFocus)
            Options.InstallRuntimes = False
            SaveOptions
        ElseIf Answer = vbNo Then
            Options.InstallRuntimes = False
            SaveOptions
        End If
     End If
    
Exit Sub
errorhandler:
    Call HandleError(Err.Number, Err.Description, Erl, "Sub Main")
End Sub

Sub EnterGame()

If Options.Debug = True Then On Error GoTo errorhandler

    ' All this needs to be setup whether or not it's online
    frmMenu.Hide
    Call LoadData
    Call InitDirectX
    Call LoadGraphics
    Call SetupGUI
    ' Fonts
    Call SetFont(Options.GameFont, FONT_SIZE)
    Running = True
    frmMain.Show
    
    If Trim$(Map(Player(MyIndex).Map).Music) <> "Theme.mp3" Then
       Call StopMusic
       If Options.Music = True Then
           If Map(Player(MyIndex).Map).Music <> vbNullString Then
               Call PlayMusic(Map(Player(MyIndex).Map).Music)
           End If
       End If
    End If
    
    If Player(MyIndex).Map >= MIN_DUNGEON_MAP And Player(MyIndex).Map <= MAX_DUNGEON_MAP Then 'And Player(MyIndex).Access <> ACCESS_ADMIN Then
        Call WarpPlayer(MyIndex, 1, 8, 12)
    End If

    Call UpdatePlayerVitals(MyIndex)
    Call BltInventory
    Call BltCharacterScreen
    
    'player editor
    If Player(MyIndex).Graphics.Body = 0 Then
        If Player(MyIndex).Graphics.Hair = 0 Then
            If Player(MyIndex).Graphics.Legs = 0 Then
                If Player(MyIndex).Graphics.Skin = 0 Then
                    CreatingCharacter = True
                    Call SetPlayerGraphics
                End If
            End If
        End If
    End If
    Call GameLoop

Exit Sub
errorhandler:
    Call HandleError(Err.Number, Err.Description, Erl, "Sub EnterGame")
End Sub

Sub EmergencyShutDown()

If Options.Debug = True Then On Error GoTo errorhandler

    Call StopMusic
    Running = False
    Call DestroyDirectX
    Call DestroyClient

Exit Sub
errorhandler:
    Call HandleError(Err.Number, Err.Description, Erl, "Sub EmergencyShutDown")
End Sub

Sub DestroyGame()

If Options.Debug = True Then On Error GoTo errorhandler

    frmMain.Visible = False
    
    If Trim$(Map(Player(MyIndex).Map).Music) <> "Theme.mp3" Then
        Call StopMusic
        Call PlayMusic("Theme.mp3")
    End If

    Running = False
    If Options.OnlineMode = True Then frmMain.socket.close
    If Options.OnlineMode = False Then Call SavePlayer(Trim$(Player(MyIndex).name))
    Call DestroyDirectX
    Call SaveData
    Call ClearData
    MyIndex = 0
    CreatingCharacter = False
    
    Call SetupMenuGui
    
    frmMenu.Visible = True
    frmMenu.picOptions.Visible = False
    frmMenu.picPlayer.Visible = False

Exit Sub
errorhandler:
    Call HandleError(Err.Number, Err.Description, Erl, "Sub DestroyGame")
End Sub

Sub DestroyClient()
    Call CloseSound
    Call SaveOptions
    End
End Sub

' Checks if a directory exists, if it doesn't, it makes it
Public Sub CheckDir(ByVal Path As String, ByVal Directory As String)
    If LCase$(Dir$(Path & Directory, vbDirectory)) <> Directory Then Call MkDir(Path & Directory)
End Sub

' Checks to see if a file exists
Public Function FileExist(ByVal FileName As String) As Boolean
    If LenB(Dir$(FileName)) > 0 Then FileExist = True
End Function

' Retrieves a string from a text file
Public Function GetVar(file As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' Writes a string in a text file
Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, file)
End Sub

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub SetPlayerGraphics()

    frmMain.fraPlayerCreate.Visible = True

    With Player(MyIndex).Graphics
        .Body = 1
        .BodyDir = "norm"
        .Hair = 1
        .HairDir = "norm"
        .Legs = 1
        .LegsDir = "norm"
        .Skin = 1
    End With
    
End Sub

Public Sub SetupGUI()
Dim i As Long
Dim Dir As String

    Dir = App.Path & "\graphics\gui\main\"
    
    With frmMain
        .picInfo.Picture = Nothing
        .picInventory.Picture = Nothing
        .picOptions.Picture = Nothing
        .picCharacter.Picture = Nothing
        .picScreen.Picture = Nothing
        .picSkills.Picture = Nothing
        .Picture = Nothing
        .imgHp.Picture = Nothing
        .imgSp.Picture = Nothing
        .imgXp.Picture = Nothing
        For i = 1 To Tabs.Tabs_Count - 1
            .picTab(i).Picture = Nothing
            .picTab(i).Picture = LoadPicture(App.Path & "\graphics\gui\main\tabs\" & i & ".bmp")
        Next
        .picBank.Picture = LoadPicture(App.Path & "\graphics\gui\main\bank.bmp")
        .txtMyChat.BackColor = RGB(8, 8, 8)
        .txtChat.BackColor = RGB(8, 8, 8)
        .BackColor = RGB(31, 31, 31)

        Select Case Options.FullScreen
            Case True
                Form.Titlebar = CBool(False)
                Form.Sizeable = CBool(True)
                frmMain.WindowState = 2 ' fullscreen
                With .txtMyChat
                    .Height = 21
                    .Left = 965
                    .top = 741
                    .Width = 398
                End With
                With .txtChat
                    .Height = 120
                    .Left = 965
                    .top = 622
                    .Width = 398
                End With
                With .picScreen
                    .Height = 384 * 2
                    .Left = 10 / 5
                    .top = 10 / 5
                    .Width = 480 * 2
                End With
                With .fraPlayerCreate
                    .top = 232 * 2
                    .Left = 66 * 4.5
                End With
            Case False
                frmMain.WindowState = 0
                .Picture = LoadPicture(Dir & "main.bmp")
                .picInfo.Picture = LoadPicture(Dir & "info.bmp")
                .picCharacter.Picture = LoadPicture(Dir & "character.bmp")
                .picSkills.Picture = LoadPicture(Dir & "skills.bmp")
                .imgHp.Picture = LoadPicture(Dir & "\bars\health.jpg")
                .imgSp.Picture = LoadPicture(Dir & "\bars\spirit.jpg")
                .imgXp.Picture = LoadPicture(Dir & "\bars\experience.jpg")
                frmMain.Width = 12000
                frmMain.Height = 8800
                Form.Sizeable = CBool(True)
                Form.Titlebar = CBool(True)
                With .txtMyChat
                    .Height = 21
                    .Left = 12
                    .top = 521
                    .Width = 476
                End With
                With .txtChat
                    .Height = 120
                    .Left = 12
                    .top = 402
                    .Width = 476
                End With
                With .picScreen
                    .Height = 384
                    .Left = 10
                    .top = 10
                    .Width = 480
                End With
                With .fraPlayerCreate
                    .top = 232
                    .Left = 66
                End With
            End Select
    End With
    
End Sub

Public Sub SetupMenuGui()
Dim i As Long
    
With frmMenu
    .Picture = Nothing
    .picOptions.Picture = Nothing
    .picPlayer.Picture = Nothing
    For i = 1 To 4
        .imgMenuButton(i).Picture = Nothing
    Next
    
    .Width = 13575
    .Height = 7860
    For i = 1 To 4
        .imgMenuButton(i).Width = 223
        .imgMenuButton(i).Height = 40
        .imgMenuButton(i).Left = 339
        .imgMenuButton(i).Picture = LoadPicture(Norm(i))
    Next
    .imgMenuButton(1).top = 235
    .imgMenuButton(2).top = 281
    .imgMenuButton(3).top = 327
    .imgMenuButton(4).top = 373
    .Picture = LoadPicture(App.Path & "\graphics\gui\menu\background.bmp")
    With .picOptions
        .Left = 182
        .top = 185
        .Width = 541
        .Height = 279
        .Picture = LoadPicture(App.Path & "\graphics\gui\menu\Frame.bmp")
    End With
    With .picPlayer
        .Left = 182
        .top = 185
        .Width = 541
        .Height = 279
        .Picture = LoadPicture(App.Path & "\graphics\gui\menu\Frame.bmp")
    End With
End With

End Sub
