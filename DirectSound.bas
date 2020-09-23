Attribute VB_Name = "modDirectSound"
Option Explicit

'***DIRECT Sound!!! YES!***

'Comes in both DX7 and DX8 versions!
'The old DX7 runs a bit slower. DX8 is pretty fast

'Const SoundsNum = 16 'How many buffers to reserve for playing sounds

'DirectX8 Declarations
'Private DX As New DirectX8
'Private DSound As DirectSound8
'Public Sounds() As DirectSoundSecondaryBuffer8

Private Priorities() As Byte


'DX7 Declarations
Private DX7 As New DirectX7
Private DSound As DirectSound
Public Sounds() As DirectSoundBuffer

Public Sub SetUpSounds(WhathWnd As Long)

ReDim Sounds(1 To BSettings(SetInd("BUFFERS", BSettings)).SettingData)
'ReDim Reserved(UBound(Sounds))
ReDim Priorities(UBound(Sounds))
Dim Desc As DSBUFFERDESC
Dim i As Integer

Dim WF As WAVEFORMATEX 'DX7 only

'Set DSound = DX.DirectSoundCreate("")
Set DSound = DX7.DirectSoundCreate("")

DSound.SetCooperativeLevel WhathWnd, DSSCL_NORMAL  'The hWnd is just the form that needs the sound
                                                        'Change as necessary
Const BlankSound = "sounds\blank.wav"
For i = LBound(Sounds) To UBound(Sounds)
    'All buffers are loaded with a default sound
    'Set Sounds(i) = DSound.CreateSoundBufferFromFile(BlankSound, Desc) 'DX8
    Set Sounds(i) = DSound.CreateSoundBufferFromFile(BlankSound, Desc, WF) 'DX7

Next i

End Sub

Public Sub StopSounds()
'Stops all sounds
Dim i As Byte

For i = LBound(Sounds) To UBound(Sounds)
    Sounds(i).Stop
Next i


End Sub

Public Sub PlaySound(File As String, Optional Priority As Byte = 1, Optional Looping As Boolean)
'Sub for playing all sounds using Direct-X
'Looks for a buffer that is not storing a sound
'that is playing. If it can't find one,
'then it just wont play the sound

'Makes sure file exists before playing it.
If FileExists(File) Then
    Dim CurSound As Integer
    Dim i As Byte
    Dim Desc As DSBUFFERDESC
    Dim HowPlay As Integer
    
    'Searches for a free buffer
    For i = LBound(Sounds) To UBound(Sounds)
        If Sounds(i).GetStatus <> DSBSTATUS_PLAYING And Sounds(i).GetStatus <> DSBSTATUS_LOOPING Then
        
            CurSound = i
            Exit For
            
        End If
        
        If i = UBound(Sounds) Then
            If Priority <= Priorities(i) Then
                'No free buffers
                CurSound = -1
            Else
                'Force play
                CurSound = i
            End If
        End If
    Next i
    
    'Wont play if there is no free buffer
    If CurSound <> -1 Then
    
        '======================================
        'DIRECT-X 8 VERSION
        
        'Set Sounds(CurSound) = DSound.CreateSoundBufferFromFile(File, Desc)
            
        '--------------------------------------
        'DX7 version
        
        Dim WF As WAVEFORMATEX
        '
        Set Sounds(CurSound) = DSound.CreateSoundBufferFromFile(File, Desc, WF)
        
        '=======================================
        If Looping = False Then
            HowPlay = DSBPLAY_DEFAULT
        Else
            HowPlay = DSBPLAY_LOOPING
        End If
        Sounds(CurSound).Play HowPlay
        Priorities(CurSound) = Priority
        
    End If
    
End If

End Sub
