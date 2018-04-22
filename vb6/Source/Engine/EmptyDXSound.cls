VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXSounde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize(hWnd As Long)
End Sub

Public Function LoadSound(sPath As String, Optional iChannels As Integer = 2, Optional iBitsPerSample As Integer = 16) As Long
End Function

Public Sub SetFrequency(soundID As Long, ByVal lFrequency As Long)
End Sub

Public Property Get Frequency(soundID As Long) As Long
End Property

Public Sub SetVolume(soundID As Long, ByVal lVolume As Long)
End Sub

Public Property Get Volume(soundID As Long) As Long
End Property

Public Sub SetPan(soundID As Long, ByVal lPan As Long)
End Sub

Public Property Get Pan(soundID As Long) As Long
End Property

Public Property Get StillPlaying(soundID As Long) As Boolean
StillPlaying = False
End Property

Public Sub PlaySound(soundID As Long, Optional bLoop As Boolean = False)
End Sub

Public Sub StopSound(soundID As Long, Optional ByVal bPause As Boolean = False)
End Sub
