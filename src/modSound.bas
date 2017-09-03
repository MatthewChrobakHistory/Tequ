Attribute VB_Name = "modSound"
Option Explicit

Public SoundIndex As Long
Public MusicIndex As Long

Public CurMusic As String
Public CurSound As String

Public SoundMusicOn As Boolean

Public Const SOUND_PATH As String = "\audio\sounds\"
Public Const MUSIC_PATH As String = "\audio\music\"

Public Sub InitSound()
    ' change and set the current path, to prevent from VB not finding BASS.DLL
    Call ChDrive(App.Path)
    Call ChDir(App.Path)

    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of bass.dll was loaded.", vbCritical)
        End
    End If

    ' initialize BASS
    If (BASS_Init(-1, 44100, 0, frmMain.hwnd, 0) = 0) Then
        Call BassError("Failed to initialise the device.")
        End
    End If
    
    ' it worked!
    SoundMusicOn = True
End Sub

Public Sub CloseSound()
    If SoundMusicOn = False Then Exit Sub

    ' Stop everything
    StopMusic
    StopSound
    
    ' Free bass.dll
    Call BASS_Free
End Sub

Public Sub StopSound()
    If SoundMusicOn = False Then Exit Sub
    
    BASS_ChannelStop (SoundIndex)
    CurSound = vbNullString
End Sub

Public Sub StopMusic()
    If SoundMusicOn = False Then Exit Sub
    
    BASS_ChannelStop (MusicIndex)
    CurMusic = vbNullString
End Sub

Public Sub PlayMusic(ByVal FileName As String)
    If SoundMusicOn = False Then Exit Sub
    If Not FileExist(App.Path & MUSIC_PATH & FileName) Then Exit Sub
    If CurMusic = FileName Then Exit Sub
    
    If Options.Music = 0 Then Exit Sub
    
    ' Stop the current music
    StopMusic
    
    ' Set it to loop
    MusicIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & MUSIC_PATH & FileName), 0, 0, BASS_SAMPLE_LOOP)
    
    ' Set the volume
    Call SetVolume(MusicIndex, frmMain.scrlVolume.value)
    
    ' See if we can find it
    If MusicIndex = 0 Then
        Call BassError("Can't open stream.")
    Else
        Call BASS_ChannelPlay(MusicIndex, False)
    End If
    
    ' Set the new current music
    CurMusic = FileName
End Sub

Public Sub PlaySound(ByVal FileName As String)
    If SoundMusicOn = False Then Exit Sub
    If Not FileExist(App.Path & SOUND_PATH & FileName) Then Exit Sub
    
    If Options.Sound = 0 Then Exit Sub
    
    ' Create the new channel
    SoundIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & FileName), 0, 0, 0)
    
    If SoundIndex = 0 Then
        Call BassError("Can't open sound.")
        Exit Sub
    Else
        ' Set the volume
        Call SetVolume(SoundIndex, frmMain.scrlSFX.value)
        Call BASS_ChannelPlay(SoundIndex, False)
    End If
        
    CurSound = FileName
End Sub

Public Sub SetVolume(ByVal channel As Long, ByVal Volume As Double)
    Call BASS_ChannelSetAttribute(channel, BASS_ATTRIB_VOL, Volume)
End Sub

' Display error messages
Public Sub BassError(ByVal ES As String)
    Debug.Print ES & vbCrLf & vbCrLf & "Error Code: " & BASS_ErrorGetCode
End Sub


