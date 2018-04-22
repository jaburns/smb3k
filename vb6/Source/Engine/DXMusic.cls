VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DirectX As New DirectX7

Private dmPerformance As DirectMusicPerformance
Private dmSegment As DirectMusicSegment
Private dmState As DirectMusicSegmentState
Private dmLoader As DirectMusicLoader

Private curVolume As Byte
'

Public Property Let Volume(RHS As Byte)
    curVolume = RHS
    dmPerformance.SetMasterVolume (RHS * 42 - 3000)
End Property
Public Property Get Volume() As Byte
    Volume = curVolume
End Property


Public Function LoadFile(sPath As String) As Boolean
On Error GoTo errout
LoadFile = False

    If Dir$(sPath) = "" Then GoTo errout
    Set dmSegment = dmLoader.LoadSegment(sPath)
    dmSegment.SetStandardMidiFile
    
LoadFile = True
errout:
End Function


Public Sub PlayMusic(Optional ByVal lRepeats As Long = -1)

    If lRepeats < 0 Then lRepeats = 2147483647

    dmSegment.SetLoopPoints 0, 0
    dmSegment.SetRepeats lRepeats

    Set dmState = dmPerformance.PlaySegment(dmSegment, 0, 0)

End Sub

Public Sub StopMusic()
    dmPerformance.Stop dmSegment, dmState, 0, 0
End Sub



Private Sub Class_Initialize()

    Set dmLoader = DirectX.DirectMusicLoaderCreate()
    Set dmPerformance = DirectX.DirectMusicPerformanceCreate()

    With dmPerformance
        .Init Nothing, 0
        .SetPort -1, 80
        .SetMasterAutoDownload True
    End With
    
    curVolume = 75
    dmPerformance.SetMasterVolume 150
    
End Sub
Private Sub Class_Terminate()
    If Not (dmPerformance Is Nothing) Then dmPerformance.CloseDown
    
    Set dmPerformance = Nothing
    Set dmSegment = Nothing
    Set dmState = Nothing
    Set dmLoader = Nothing
End Sub
