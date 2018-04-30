Attribute VB_Name = "mInput"
Option Explicit

Public Type udtControlList
    btnRun As Byte
    btnShoot As Byte
    btnJump As Byte
    btnRelease As Byte
    btnLeaveLevel As Byte
    btnPause As Byte
End Type
Public Type udtControlData
    btnKeyboard As udtControlList
    btnJoystick As udtControlList
    btnLeft As Byte
    btnRight As Byte
    btnDuck As Byte
    btnClimb As Byte
    btnQuit As Byte
    joyXSens As Integer
    joyYSens As Integer
End Type

Public oControlData As udtControlData

Private oKeyboard As DXKeyboard
Private oJoystick As DXJoystick

Public Enum GameKeys
    Left
    Right
    Up
    Down
    Climb
    Jump
    Run
    Shoot
    Quit
    Pause
    Release
    LeaveLevel
    DebugA
    DebugB
    JoystickTestKey
    LoadGameKey
    QuitFromPauseKey
End Enum

Public Function SaveSlotKeyDown(ByVal i As Long) As Boolean
SaveSlotKeyDown = False
    If i = 0 Then
        SaveSlotKeyDown = oKeyboard.IsDown(11)
    ElseIf i < 10 Then
        SaveSlotKeyDown = oKeyboard.IsDown(CByte(i + 1))
    End If
End Function

Public Sub InitInput(lhWnd As Long)
    LoadControlsFromFile App.Path & "\Controls.3kc"
    Set oKeyboard = New DXKeyboard
    Set oJoystick = New DXJoystick
    oKeyboard.Initialize lhWnd
    oJoystick.Initialize lhWnd
End Sub

Public Sub UpdateInput()
    oJoystick.Update
End Sub

Public Sub KillInput()
    Set oKeyboard = Nothing
    Set oJoystick = Nothing
End Sub


Public Function JoystickExists() As Boolean
    JoystickExists = oJoystick.JoystickPresent
End Function


Public Function GameKeyDown(lKey As GameKeys) As Boolean
    
    GameKeyDown = JoypadKeyDown(lKey)
    If GameKeyDown Then Exit Function
    
    Select Case lKey
        Case GameKeys.Left
            If (oKeyboard.IsDown(oControlData.btnLeft) And Not oKeyboard.IsDown(oControlData.btnRight)) Then GameKeyDown = True
        Case GameKeys.Right
            If (oKeyboard.IsDown(oControlData.btnRight) And Not oKeyboard.IsDown(oControlData.btnLeft)) Then GameKeyDown = True
        Case GameKeys.Climb
            If (oKeyboard.IsDown(oControlData.btnClimb) And Not oKeyboard.IsDown(oControlData.btnKeyboard.btnJump)) Then GameKeyDown = True
        Case GameKeys.Up
            If (oKeyboard.IsDown(oControlData.btnClimb)) Then GameKeyDown = True
        Case GameKeys.Down
            If (oKeyboard.IsDown(oControlData.btnDuck)) Then GameKeyDown = True
        Case GameKeys.Jump
            If (oKeyboard.IsDown(oControlData.btnKeyboard.btnJump)) Then GameKeyDown = True
        Case GameKeys.Run
            If (oKeyboard.IsDown(oControlData.btnKeyboard.btnRun)) Then GameKeyDown = True
        Case GameKeys.Shoot
            If (oKeyboard.IsDown(oControlData.btnKeyboard.btnShoot)) Then GameKeyDown = True
        Case GameKeys.Quit
            If (oKeyboard.IsDown(oControlData.btnQuit)) Then GameKeyDown = True
        Case GameKeys.Pause
            If (oKeyboard.IsDown(oControlData.btnKeyboard.btnPause)) Then GameKeyDown = True
        Case GameKeys.Release
            If (oKeyboard.IsDown(oControlData.btnKeyboard.btnRelease)) Then GameKeyDown = True
        Case GameKeys.LeaveLevel
            If (oKeyboard.IsDown(oControlData.btnKeyboard.btnLeaveLevel)) Then GameKeyDown = True
        Case GameKeys.JoystickTestKey
            If (oKeyboard.IsDown(oKeyboard.Key_J)) Then GameKeyDown = True
        Case GameKeys.DebugA
            If oKeyboard.IsDown(oKeyboard.Key_F9) Then GameKeyDown = True
        Case GameKeys.DebugB
            If oKeyboard.IsDown(oKeyboard.Key_F10) Then GameKeyDown = True
        Case GameKeys.LoadGameKey
            If oKeyboard.IsDown(oKeyboard.Key_L) Then GameKeyDown = True
        Case GameKeys.QuitFromPauseKey
            If oKeyboard.IsDown(oKeyboard.Key_Q) Then GameKeyDown = True
    End Select
    
End Function


Private Function JoypadKeyDown(lKey As GameKeys) As Boolean

    oJoystick.Update
    Select Case lKey
        Case GameKeys.Left
            If oJoystick.XAxis < -oControlData.joyXSens Then JoypadKeyDown = True
        Case GameKeys.Right
            If oJoystick.XAxis > oControlData.joyXSens Then JoypadKeyDown = True
        Case GameKeys.Climb
            If (oJoystick.YAxis < -oControlData.joyYSens And Not oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnJump))) Then JoypadKeyDown = True
        Case GameKeys.Up
            If oJoystick.YAxis < -oControlData.joyYSens Then JoypadKeyDown = True
        Case GameKeys.Down
            If oJoystick.YAxis > oControlData.joyYSens Then JoypadKeyDown = True
        Case GameKeys.Jump
            If oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnJump)) Then JoypadKeyDown = True
        Case GameKeys.Run
            If oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnRun)) Then JoypadKeyDown = True
        Case GameKeys.Shoot
            If oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnShoot)) Then JoypadKeyDown = True
        Case GameKeys.Quit
            If (oKeyboard.IsDown(oControlData.btnQuit)) Then JoypadKeyDown = True
        Case GameKeys.Pause
            If oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnPause)) Then JoypadKeyDown = True
        Case GameKeys.Release
            If oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnRelease)) Then JoypadKeyDown = True
        Case GameKeys.LeaveLevel
            If oJoystick.ButtonDown(CInt(oControlData.btnJoystick.btnLeaveLevel)) Then JoypadKeyDown = True
    End Select
    
End Function

Public Sub SetDefaultControls()
    With oControlData.btnKeyboard
        .btnJump = 44 'oKeyboard.Key_Z
        .btnRun = 42 'oKeyboard.Key_LSHIFT
        .btnShoot = 57 'oKeyboard.Key_SPACE
        .btnRelease = 45 'oKeyboard.Key_X
        .btnLeaveLevel = 14 'oKeyboard.Key_BACKSPACE
        .btnPause = 25 'oKeyboard.Key_P
    End With
    With oControlData.btnJoystick
        .btnJump = 1
        .btnRun = 2
        .btnShoot = 3
        .btnRelease = 4
        .btnLeaveLevel = 5
        .btnPause = 6
    End With
    With oControlData
        .btnLeft = 203 'oKeyboard.Key_LEFT
        .btnRight = 205 'oKeyboard.Key_RIGHT
        .btnDuck = 208 'oKeyboard.Key_DOWN
        .btnClimb = 200 'oKeyboard.Key_UP
        .btnQuit = 1 'oKeyboard.Key_ESCAPE
        .joyXSens = 2500
        .joyYSens = 3500
    End With
End Sub

Public Function DebugKeysPressed() As Boolean
With oKeyboard
DebugKeysPressed = .IsDown(.Key_PAGEUP) And .IsDown(.Key_PAGEDOWN) And .IsDown(.Key_TAB)
End With
End Function

Public Sub LoadControlsFromFile(ByVal sPath As String)
'On Error GoTo errOut
'Dim fFile As Long
'fFile = FreeFile
'Open sPath For Input As fFile
'Close fFile
'Open sPath For Binary Access Read Lock Write As fFile
'Get #fFile, 1, oControlData
'Close fFile
'GoTo endNoErr
'errOut:
'Close fFile
SetDefaultControls
'endNoErr:
End Sub

Public Sub SaveControlsToFile(ByVal sPath As String)
Dim i As Long
'Dim fFile As Long
'fFile = FreeFile
'Open sPath For Binary Access Write Lock Read Write As fFile
'Put #fFile, 1, oControlData
'Close fFile
End Sub

