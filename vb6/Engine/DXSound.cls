VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtSoundBuffer
    Alive As Boolean
    Buffer As DirectSoundBuffer
End Type

Private DirectX As New DirectX7
Private DirectSound As DirectSound

Private SBuffers() As udtSoundBuffer
'


' Init direct sound
''''''''''''''''''''
Public Sub Initialize(hWnd As Long)
    ReDim SBuffers(0)
    Set DirectSound = DirectX.DirectSoundCreate("")
    DirectSound.SetCooperativeLevel hWnd, DSSCL_EXCLUSIVE
End Sub


' Build a sound buffer from a file
'''''''''''''''''''''''''''''''''''
Public Function LoadSound(sPath As String, Optional iChannels As Integer = 2, Optional iBitsPerSample As Integer = 16) As Long
On Error GoTo errOut
Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX
Dim newID As Long

    'determine the ID of the new sound
    If SBuffers(UBound(SBuffers)).Alive Then ReDim Preserve SBuffers(UBound(SBuffers) + 1)
    newID = UBound(SBuffers)
    
    'set constant properties
    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DsWave.nFormatTag = WAVE_FORMAT_PCM
    DsWave.lSamplesPerSec = 22050
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
    
    'set variable properties
    DsWave.nChannels = iChannels
    DsWave.nBitsPerSample = iBitsPerSample

    'build sound buffer
    Set SBuffers(newID).Buffer = DirectSound.CreateSoundBufferFromFile(sPath, DsDesc, DsWave)
    SBuffers(newID).Alive = True
    
    LoadSound = newID
    
Exit Function
errOut:
SBuffers(newID).Alive = False
Set SBuffers(newID).Buffer = Nothing
LoadSound = -1
End Function



Public Sub SetFrequency(soundID As Long, ByVal lFrequency As Long)
    If lFrequency < 100 Then lFrequency = 100
    If lFrequency > 100000 Then lFrequency = 100000
    SBuffers(soundID).Buffer.SetFrequency lFrequency
End Sub
Public Property Get Frequency(soundID As Long) As Long
    Frequency = SBuffers(soundID).Buffer.GetFrequency
End Property


Public Sub SetVolume(soundID As Long, ByVal lVolume As Long)
    If lVolume < 0 Then lVolume = 0
    If lVolume > 10000 Then lVolume = 10000
    lVolume = lVolume - 10000
    SBuffers(soundID).Buffer.SetVolume lVolume
End Sub
Public Property Get Volume(soundID As Long) As Long
    Volume = SBuffers(soundID).Buffer.GetVolume
End Property


Public Sub SetPan(soundID As Long, ByVal lPan As Long)
    If lPan < -10000 Then lPan = -10000
    If lPan > 10000 Then lPan = 10000
    SBuffers(soundID).Buffer.SetPan lPan
End Sub
Public Property Get Pan(soundID As Long) As Long
    Pan = SBuffers(soundID).Buffer.GetPan
End Property


Public Property Get StillPlaying(soundID As Long) As Boolean
    If SBuffers(soundID).Buffer.GetStatus = DSBSTATUS_PLAYING _
    Or SBuffers(soundID).Buffer.GetStatus = DSBSTATUS_LOOPING Then
        StillPlaying = True
    Else
        StillPlaying = False
    End If
End Property


' Play a sound
'''''''''''''''
Public Sub PlaySound(soundID As Long, Optional bLoop As Boolean = False)
    If bLoop Then
        SBuffers(soundID).Buffer.Play DSBPLAY_LOOPING
    Else
        SBuffers(soundID).Buffer.Play DSBPLAY_DEFAULT
    End If
End Sub


' Stop or pause a sound
''''''''''''''''''''''''
Public Sub StopSound(soundID As Long, Optional ByVal bPause As Boolean = False)
    SBuffers(soundID).Buffer.Stop
    If Not bPause Then SBuffers(soundID).Buffer.SetCurrentPosition 0
End Sub


' Destroys everything automatically
''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
Dim i As Long
    For i = 0 To UBound(SBuffers)
        If SBuffers(i).Alive Then Set SBuffers(i).Buffer = Nothing
    Next i
    Set DirectSound = Nothing
    Set DirectX = Nothing
End Sub
