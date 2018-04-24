VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objBasicAudio As IBasicAudio
Private objBasicVideo As IBasicVideo
Private objMediaEvent As IMediaEvent
Private objVideoWindow As IVideoWindow
Private objMediaControl As IMediaControl
Private objMediaPosition As IMediaPosition
Private Const WS_VISIBLE = &H10000000

' Clean up after DShow
'''''''''''''''''''''''
Private Sub DDestroy()
    On Local Error GoTo RemoveDShowError
    
    'If a MediaControl instance exists, then stop it from playing
    If ObjPtr(objMediaControl) > 0 Then
        objMediaControl.Stop
    End If
    'If a VideoWindow instance exists, then remove the link to the render target
    If ObjPtr(objVideoWindow) > 0 Then
        objVideoWindow.Owner = 0
    End If
    'Destroy all objects
    If ObjPtr(objBasicAudio) > 0 Then Set objBasicAudio = Nothing
    If ObjPtr(objBasicVideo) > 0 Then Set objBasicVideo = Nothing
    If ObjPtr(objMediaControl) > 0 Then Set objMediaControl = Nothing
    If ObjPtr(objVideoWindow) > 0 Then Set objVideoWindow = Nothing
    If ObjPtr(objMediaPosition) > 0 Then Set objMediaPosition = Nothing
Exit Sub

RemoveDShowError:
    Err.Clear
    Exit Sub
End Sub

' Open a file for viewing
'''''''''''''''''''''''''''
Public Sub DOpenFile(Filename As String, RectTop As Integer, RectLeft As Integer, rectWidth As Integer, rectHeight As Integer, hWndTarget As Long)
On Local Error GoTo OpenFileError
    
    'First clean up
    DDestroy
    
    Set objMediaControl = New FilgraphManager
    Call objMediaControl.RenderFile(Filename)
    
    ' Set the audio properties
    Set objBasicAudio = objMediaControl
    objBasicAudio.Volume = 0
    objBasicAudio.Balance = 0
    
    ' Set the video properties
    Set objVideoWindow = objMediaControl
    objVideoWindow.WindowStyle = WS_VISIBLE
    objVideoWindow.Top = RectTop
    objVideoWindow.Left = RectLeft
    objVideoWindow.width = rectWidth
    objVideoWindow.height = rectHeight
    objVideoWindow.Owner = hWndTarget
    
    Set objMediaEvent = objMediaControl
    Set objMediaPosition = objMediaControl
    objMediaPosition.Rate = 1
    
Exit Sub

OpenFileError:
    Err.Clear
    Resume Next
End Sub

' Play if DShow is initialized and a file is loaded
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DPlay()
On Error Resume Next
    If ObjPtr(objMediaPosition) > 0 Then
        'If at end, then restart
        If CLng(objMediaPosition.CurrentPosition) = CLng(objMediaPosition.Duration) Then
            objMediaPosition.CurrentPosition = 0
        End If
        objMediaControl.Run
    End If
End Sub

' Pause if DShow is initialized and a file is loaded
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DPause()
    If ObjPtr(objMediaControl) > 0 Then
        objMediaControl.Pause
    End If
End Sub

' Stop if DShow is initialized and a file is loaded
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DStop()
    If (ObjPtr(objMediaControl) > 0) And (ObjPtr(objMediaPosition) > 0) Then
        objMediaControl.Stop
        objMediaPosition.CurrentPosition = 0
    End If
End Sub

' Seek to a certain position in the video
''''''''''''''''''''''''''''''''''''''''''
Public Sub DSeek(Amount As Double)
    'Seek if DShow is initialized and a file is loaded
    If ObjPtr(objMediaPosition) > 0 Then
        'if position is out of bounds then correct
        If objMediaPosition.CurrentPosition + Amount < 0 Then
            objMediaPosition.CurrentPosition = 0
        ElseIf objMediaPosition.CurrentPosition + Amount > objMediaPosition.Duration Then
            objMediaPosition.CurrentPosition = objMediaPosition.Duration
        Else
            'else just set the new position
            objMediaPosition.CurrentPosition = objMediaPosition.CurrentPosition + Amount
        End If
    End If
End Sub

' Returns true if the movie is currently playing
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get DStillPlaying() As Boolean
    On Local Error Resume Next
    Dim EvCode As Long
    'Check only if DShow is initialized and a file is loaded
    If ObjPtr(objMediaEvent) > 0 Then
        objMediaEvent.WaitForCompletion 0, EvCode
        If EvCode = 0 Then
            DStillPlaying = True
        Else
            DStillPlaying = False
        End If
    End If
End Property


' Get and let the rate of the movie
''''''''''''''''''''''''''''''''''''
Public Property Get DRate() As Double
    If ObjPtr(objMediaPosition) > 0 Then DRate = objMediaPosition.Rate Else DRate = 0
End Property
Public Property Let DRate(RHS As Double)
    If RHS <= 0 Then RHS = 0.1
    If ObjPtr(objMediaPosition) > 0 Then objMediaPosition.Rate = RHS
End Property


' Get the duration, its read only
''''''''''''''''''''''''''''''''''
Public Property Get DDuration() As Double
    If ObjPtr(objMediaPosition) > 0 Then DDuration = objMediaPosition.Duration
End Property


' Get and let the position of the movie
''''''''''''''''''''''''''''''''''''''''
Public Property Get DPosition() As Double
    If ObjPtr(objMediaPosition) > 0 Then DPosition = objMediaPosition.CurrentPosition
End Property
Public Property Let DPosition(RHS As Double)
    If ObjPtr(objMediaPosition) > 0 Then
        If RHS < 0 Then RHS = 0
        If RHS > objMediaPosition.Duration Then RHS = objMediaPosition.Duration
        objMediaPosition.CurrentPosition = RHS
    End If
End Property


' Get and let the volume of the sound ( 0 to 100)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get DVolume() As Long
    If ObjPtr(objBasicAudio) > 0 Then DVolume = objBasicAudio.Volume
    DVolume = (DVolume + 10000) / 100
End Property
Public Property Let DVolume(RHS As Long)
    'Volume must be between 0 (loudest) and -10000 (disabled)
    If RHS > 100 Then RHS = 100
    If RHS < 0 Then RHS = 0
    
    If ObjPtr(objBasicAudio) > 0 Then objBasicAudio.Volume = (RHS * 100) - 10000
End Property


' Get and let the balance of the sound
' Between -100 left and 100 right
'''''''''''''''''''''''''''''''''''''''
Public Property Get DBalance() As Long
    If ObjPtr(objBasicAudio) > 0 Then DBalance = objBasicAudio.Balance
    DBalance = DBalance / 100
End Property
Public Property Let DBalance(RHS As Long)
    If RHS < -100 Then RHS = -100
    If RHS > 100 Then RHS = 100
    If ObjPtr(objBasicAudio) > 0 Then objBasicAudio.Balance = RHS * 100
End Property


Private Sub Class_Terminate()
    DDestroy
End Sub
