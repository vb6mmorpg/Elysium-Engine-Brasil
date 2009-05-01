Attribute VB_Name = "modSound"
' Copyright (c) 2008 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

' DirectSound7
Public Ds As DirectSound
Public dsbuffer As DirectSoundBuffer
Public DsDesc As DSBUFFERDESC
Public DsWave As WAVEFORMATEX

' DirectMusic7
Public perf As DirectMusicPerformance
Public seg As DirectMusicSegment
Public segstate As DirectMusicSegmentState
Public loader As DirectMusicLoader

Public CurrentSong As String

Sub InitDirectSM()

    ' DirectSound7
    Set Ds = DX.DirectSoundCreate(vbNullString)
    Call Ds.SetCooperativeLevel(frmMirage.hWnd, DSSCL_NORMAL)
    
    ' DirectMusic7
    Set loader = DX.DirectMusicLoaderCreate()
    Set perf = DX.DirectMusicPerformanceCreate()
    Call perf.Init(Nothing, 0)
    perf.SetPort -1, 80
    Call perf.SetMasterAutoDownload(True)
    perf.SetMasterVolume (MusicVolume * 42)
    
End Sub

Public Sub PlayMidi(ByVal Song As String)
On Error GoTo ErrHandler

If Song = vbNullString Or Song = "Nenhuma" Then
    Call StopMidi
    Exit Sub
End If

If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Music")) = "1" Then

    If FileExist("Músicas\" & Song) = False Then
    Call AddText("Não foi possível tocar a música " & Song & ".", 2)
    Exit Sub
    End If

    If CurrentSong <> Song Then
        CurrentSong = Song
        If Not (seg Is Nothing) Then Set seg = Nothing
        Set seg = loader.LoadSegment("Músicas\" & Song)
        seg.SetStandardMidiFile
        Call perf.PlaySegment(seg, 0, 0)
    End If
Else
    Call StopMidi
End If

Exit Sub

ErrHandler:
    Call AddText("Houve uma falha ao tentar tocar a música.", 2)
    Exit Sub
End Sub

Public Sub StopMidi()
Dim I As Long

    CurrentSong = vbNullString
    If perf Is Nothing Then Exit Sub
    Call perf.Stop(seg, segstate, 0, 0)
End Sub

Public Sub MakeMidiLoop()
    If seg Is Nothing Then Exit Sub
    If perf Is Nothing Then Exit Sub
    If perf.IsPlaying(seg, segstate) = False And CurrentSong <> vbNullString Then
        Set segstate = perf.PlaySegment(seg, 0, 0)
    End If
End Sub

Public Sub PlaySound(ByVal Sound As String)
On Error GoTo ErrHandler

If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Sound")) = "1" Then
    
        If FileExist("SFX\" & Sound) = False Then
        Call AddText("Não foi possível tocar o som " & Sound, 2)
        Exit Sub
        End If
        
    If Not (dsbuffer Is Nothing) Then Set dsbuffer = Nothing
    Set dsbuffer = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\" & Sound, DsDesc, DsWave)
    dsbuffer.Play DSBPLAY_DEFAULT
End If
    
Exit Sub

ErrHandler:
    Call AddText("Houve uma falha ao tentar reproduzir um efeito sonoro.", 2)
    Exit Sub
End Sub

Public Sub StopSound()
    dsbuffer.Stop
    dsbuffer.SetCurrentPosition 0
End Sub
