Attribute VB_Name = "mVideo"
Option Explicit

Private oV As DXMovie

Public Function VideoInUse() As Boolean
    VideoInUse = oV.DStillPlaying
End Function

Public Sub InitVideo()
    Set oV = New DXMovie
End Sub

Public Sub LoadMPEGMusic(ByVal sPath As String)
    oV.DStop
    oV.DOpenFile sPath, 0, 0, 0, 0, frmMain.hWnd
End Sub

Public Sub LoadMPEGVideo(ByVal sPath As String)
    oV.DStop
    oV.DOpenFile sPath, 0, 0, 640, 480, frmMain.hWnd
End Sub

Public Sub PauseMPEG()
    oV.DPause
End Sub
Public Sub PlayMPEG()
    oV.DPlay
End Sub
Public Sub StopMPEG()
    oV.DStop
End Sub

Public Sub KillVideo()
    Set oV = Nothing
End Sub
