VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DX As New DirectX7
Private DI As DirectInput
Private DIDev As DirectInputDevice

' All DX Keyboard consts as read only properties
Public Property Get Key_ESCAPE() As Byte
    Key_ESCAPE = 1
End Property
Public Property Get Key_1() As Byte
    Key_1 = 2
End Property
Public Property Get Key_2() As Byte
    Key_2 = 3
End Property
Public Property Get Key_3() As Byte
    Key_3 = 4
End Property
Public Property Get Key_4() As Byte
    Key_4 = 5
End Property
Public Property Get Key_5() As Byte
    Key_5 = 6
End Property
Public Property Get Key_6() As Byte
    Key_6 = 7
End Property
Public Property Get Key_7() As Byte
    Key_7 = 8
End Property
Public Property Get Key_8() As Byte
    Key_8 = 9
End Property
Public Property Get Key_9() As Byte
    Key_9 = 10
End Property
Public Property Get Key_0() As Byte
    Key_0 = 11
End Property
Public Property Get Key_MINUS() As Byte
    Key_MINUS = 12
End Property
Public Property Get Key_EQUALS() As Byte
    Key_EQUALS = 13
End Property
Public Property Get Key_BACKSPACE() As Byte
    Key_BACKSPACE = 14
End Property
Public Property Get Key_TAB() As Byte
    Key_TAB = 15
End Property
Public Property Get Key_Q() As Byte
    Key_Q = 16
End Property
Public Property Get Key_W() As Byte
    Key_W = 17
End Property
Public Property Get Key_E() As Byte
    Key_E = 18
End Property
Public Property Get Key_R() As Byte
    Key_R = 19
End Property
Public Property Get Key_T() As Byte
    Key_T = 20
End Property
Public Property Get Key_Y() As Byte
    Key_Y = 21
End Property
Public Property Get Key_U() As Byte
    Key_U = 22
End Property
Public Property Get Key_I() As Byte
    Key_I = 23
End Property
Public Property Get Key_O() As Byte
    Key_O = 24
End Property
Public Property Get Key_P() As Byte
    Key_P = 25
End Property
Public Property Get Key_LBRACKET() As Byte
    Key_LBRACKET = 26
End Property
Public Property Get Key_RBRACKET() As Byte
    Key_RBRACKET = 27
End Property
Public Property Get Key_RETURN() As Byte
    Key_RETURN = 28
End Property
Public Property Get Key_LCONTROL() As Byte
    Key_LCONTROL = 29
End Property
Public Property Get Key_A() As Byte
    Key_A = 30
End Property
Public Property Get Key_S() As Byte
    Key_S = 31
End Property
Public Property Get Key_D() As Byte
    Key_D = 32
End Property
Public Property Get Key_F() As Byte
    Key_F = 33
End Property
Public Property Get Key_G() As Byte
    Key_G = 34
End Property
Public Property Get Key_H() As Byte
    Key_H = 35
End Property
Public Property Get Key_J() As Byte
    Key_J = 36
End Property
Public Property Get Key_K() As Byte
    Key_K = 37
End Property
Public Property Get Key_L() As Byte
    Key_L = 38
End Property
Public Property Get Key_SEMICOLON() As Byte
    Key_SEMICOLON = 39
End Property
Public Property Get Key_APOSTROPHE() As Byte
    Key_APOSTROPHE = 40
End Property
Public Property Get Key_GRAVE() As Byte
    Key_GRAVE = 41
End Property
Public Property Get Key_LSHIFT() As Byte
    Key_LSHIFT = 42
End Property
Public Property Get Key_BACKSLASH() As Byte
    Key_BACKSLASH = 43
End Property
Public Property Get Key_Z() As Byte
    Key_Z = 44
End Property
Public Property Get Key_X() As Byte
    Key_X = 45
End Property
Public Property Get Key_C() As Byte
    Key_C = 46
End Property
Public Property Get Key_V() As Byte
    Key_V = 47
End Property
Public Property Get Key_B() As Byte
    Key_B = 48
End Property
Public Property Get Key_N() As Byte
    Key_N = 49
End Property
Public Property Get Key_M() As Byte
    Key_M = 50
End Property
Public Property Get Key_COMMA() As Byte
    Key_COMMA = 51
End Property
Public Property Get Key_PERIOD() As Byte
    Key_PERIOD = 52
End Property
Public Property Get Key_SLASH() As Byte
    Key_SLASH = 53
End Property
Public Property Get Key_RSHIFT() As Byte
    Key_RSHIFT = 54
End Property
Public Property Get Key_MULTIPLY() As Byte
    Key_MULTIPLY = 55
End Property
Public Property Get Key_LALT() As Byte
    Key_LALT = 56
End Property
Public Property Get Key_SPACE() As Byte
    Key_SPACE = 57
End Property
Public Property Get Key_CAPSLOCK() As Byte
    Key_CAPSLOCK = 58
End Property
Public Property Get Key_F1() As Byte
    Key_F1 = 59
End Property
Public Property Get Key_F2() As Byte
    Key_F2 = 60
End Property
Public Property Get Key_F3() As Byte
    Key_F3 = 61
End Property
Public Property Get Key_F4() As Byte
    Key_F4 = 62
End Property
Public Property Get Key_F5() As Byte
    Key_F5 = 63
End Property
Public Property Get Key_F6() As Byte
    Key_F6 = 64
End Property
Public Property Get Key_F7() As Byte
    Key_F7 = 65
End Property
Public Property Get Key_F8() As Byte
    Key_F8 = 66
End Property
Public Property Get Key_F9() As Byte
    Key_F9 = 67
End Property
Public Property Get Key_F10() As Byte
    Key_F10 = 68
End Property
Public Property Get Key_NUMLOCK() As Byte
    Key_NUMLOCK = 69
End Property
Public Property Get Key_SCROLL() As Byte
    Key_SCROLL = 70
End Property
Public Property Get Key_NUMPAD7() As Byte
    Key_NUMPAD7 = 71
End Property
Public Property Get Key_NUMPAD8() As Byte
    Key_NUMPAD8 = 72
End Property
Public Property Get Key_NUMPAD9() As Byte
    Key_NUMPAD9 = 73
End Property
Public Property Get Key_SUBTRACT() As Byte
    Key_SUBTRACT = 74
End Property
Public Property Get Key_NUMPAD4() As Byte
    Key_NUMPAD4 = 75
End Property
Public Property Get Key_NUMPAD5() As Byte
    Key_NUMPAD5 = 76
End Property
Public Property Get Key_NUMPAD6() As Byte
    Key_NUMPAD6 = 77
End Property
Public Property Get Key_ADD() As Byte
    Key_ADD = 78
End Property
Public Property Get Key_NUMPAD1() As Byte
    Key_NUMPAD1 = 79
End Property
Public Property Get Key_NUMPAD2() As Byte
    Key_NUMPAD2 = 80
End Property
Public Property Get Key_NUMPAD3() As Byte
    Key_NUMPAD3 = 81
End Property
Public Property Get Key_NUMPAD0() As Byte
    Key_NUMPAD0 = 82
End Property
Public Property Get Key_DECIMAL() As Byte
    Key_DECIMAL = 83
End Property
Public Property Get Key_F11() As Byte
    Key_F11 = 87
End Property
Public Property Get Key_F12() As Byte
    Key_F12 = 88
End Property
Public Property Get Key_NUMPADENTER() As Byte
    Key_NUMPADENTER = 156
End Property
Public Property Get Key_RCONTROL() As Byte
    Key_RCONTROL = 157
End Property
Public Property Get Key_DIVIDE() As Byte
    Key_DIVIDE = 181
End Property
Public Property Get Key_RALT() As Byte
    Key_RALT = 184
End Property
Public Property Get Key_PAUSE() As Byte
    Key_PAUSE = 197
End Property
Public Property Get Key_HOME() As Byte
    Key_HOME = 199
End Property
Public Property Get Key_UP() As Byte
    Key_UP = 200
End Property
Public Property Get Key_PAGEUP() As Byte
    Key_PAGEUP = 201
End Property
Public Property Get Key_LEFT() As Byte
    Key_LEFT = 203
End Property
Public Property Get Key_RIGHT() As Byte
    Key_RIGHT = 205
End Property
Public Property Get Key_END() As Byte
    Key_END = 207
End Property
Public Property Get Key_DOWN() As Byte
    Key_DOWN = 208
End Property
Public Property Get Key_PAGEDOWN() As Byte
    Key_PAGEDOWN = 209
End Property
Public Property Get Key_INSERT() As Byte
    Key_INSERT = 210
End Property
Public Property Get Key_DELETE() As Byte
    Key_DELETE = 211
End Property

'this function is used to read key values
Public Function GetDXKeyName(bKeycode As Byte) As String
    Select Case bKeycode
        Case 1
            GetDXKeyName = "ESCAPE"
        Case 2
            GetDXKeyName = "1"
        Case 3
            GetDXKeyName = "2"
        Case 4
            GetDXKeyName = "3"
        Case 5
            GetDXKeyName = "4"
        Case 6
            GetDXKeyName = "5"
        Case 7
            GetDXKeyName = "6"
        Case 8
            GetDXKeyName = "7"
        Case 9
            GetDXKeyName = "8"
        Case 10
            GetDXKeyName = "9"
        Case 11
            GetDXKeyName = "0"
        Case 12
            GetDXKeyName = "MINUS"
        Case 13
            GetDXKeyName = "EQUALS"
        Case 14
            GetDXKeyName = "BACKSPACE"
        Case 15
            GetDXKeyName = "TAB"
        Case 16
            GetDXKeyName = "Q"
        Case 17
            GetDXKeyName = "W"
        Case 18
            GetDXKeyName = "E"
        Case 19
            GetDXKeyName = "R"
        Case 20
            GetDXKeyName = "T"
        Case 21
            GetDXKeyName = "Y"
        Case 22
            GetDXKeyName = "U"
        Case 23
            GetDXKeyName = "I"
        Case 24
            GetDXKeyName = "O"
        Case 25
            GetDXKeyName = "P"
        Case 26
            GetDXKeyName = "LBRACKET"
        Case 27
            GetDXKeyName = "RBRACKET"
        Case 28
            GetDXKeyName = "RETURN"
        Case 29
            GetDXKeyName = "LCONTROL"
        Case 30
            GetDXKeyName = "A"
        Case 31
            GetDXKeyName = "S"
        Case 32
            GetDXKeyName = "D"
        Case 33
            GetDXKeyName = "F"
        Case 34
            GetDXKeyName = "G"
        Case 35
            GetDXKeyName = "H"
        Case 36
            GetDXKeyName = "J"
        Case 37
            GetDXKeyName = "K"
        Case 38
            GetDXKeyName = "L"
        Case 39
            GetDXKeyName = "SEMICOLON"
        Case 40
            GetDXKeyName = "APOSTROPHE"
        Case 41
            GetDXKeyName = "GRAVE"
        Case 42
            GetDXKeyName = "LSHIFT"
        Case 43
            GetDXKeyName = "BACKSLASH"
        Case 44
            GetDXKeyName = "Z"
        Case 45
            GetDXKeyName = "X"
        Case 46
            GetDXKeyName = "C"
        Case 47
            GetDXKeyName = "V"
        Case 48
            GetDXKeyName = "B"
        Case 49
            GetDXKeyName = "N"
        Case 50
            GetDXKeyName = "M"
        Case 51
            GetDXKeyName = "COMMA"
        Case 52
            GetDXKeyName = "PERIOD"
        Case 53
            GetDXKeyName = "SLASH"
        Case 54
            GetDXKeyName = "RSHIFT"
        Case 55
            GetDXKeyName = "MULTIPLY"
        Case 56
            GetDXKeyName = "LALT"
        Case 57
            GetDXKeyName = "SPACE"
        Case 58
            GetDXKeyName = "CAPSLOCK"
        Case 59
            GetDXKeyName = "F1"
        Case 60
            GetDXKeyName = "F2"
        Case 61
            GetDXKeyName = "F3"
        Case 62
            GetDXKeyName = "F4"
        Case 63
            GetDXKeyName = "F5"
        Case 64
            GetDXKeyName = "F6"
        Case 65
            GetDXKeyName = "F7"
        Case 66
            GetDXKeyName = "F8"
        Case 67
            GetDXKeyName = "F9"
        Case 68
            GetDXKeyName = "F10"
        Case 69
            GetDXKeyName = "NUMLOCK"
        Case 70
            GetDXKeyName = "SCROLL"
        Case 71
            GetDXKeyName = "NUMPAD7"
        Case 72
            GetDXKeyName = "NUMPAD8"
        Case 73
            GetDXKeyName = "NUMPAD9"
        Case 74
            GetDXKeyName = "SUBTRACT"
        Case 75
            GetDXKeyName = "NUMPAD4"
        Case 76
            GetDXKeyName = "NUMPAD5"
        Case 77
            GetDXKeyName = "NUMPAD6"
        Case 78
            GetDXKeyName = "ADD"
        Case 79
            GetDXKeyName = "NUMPAD1"
        Case 80
            GetDXKeyName = "NUMPAD2"
        Case 81
            GetDXKeyName = "NUMPAD3"
        Case 82
            GetDXKeyName = "NUMPAD0"
        Case 83
            GetDXKeyName = "DECIMAL"
        Case 87
            GetDXKeyName = "F11"
        Case 88
            GetDXKeyName = "F12"
        Case 156
            GetDXKeyName = "NUMPADENTER"
        Case 157
            GetDXKeyName = "RCONTROL"
        Case 181
            GetDXKeyName = "DIVIDE"
        Case 184
            GetDXKeyName = "RALT"
        Case 197
            GetDXKeyName = "PAUSE"
        Case 199
            GetDXKeyName = "HOME"
        Case 200
            GetDXKeyName = "UP"
        Case 201
            GetDXKeyName = "PAGEUP"
        Case 203
            GetDXKeyName = "LEFT"
        Case 205
            GetDXKeyName = "RIGHT"
        Case 207
            GetDXKeyName = "END"
        Case 208
            GetDXKeyName = "DOWN"
        Case 209
            GetDXKeyName = "PAGEDOWN"
        Case 210
            GetDXKeyName = "INSERT"
        Case 211
            GetDXKeyName = "DELETE"
    End Select
End Function



'
' initializes directinput
'
Public Sub Initialize(hWnd As Long)
    Set DI = DX.DirectInputCreate()
    Set DIDev = DI.CreateDevice("GUID_SysKeyboard")
    DIDev.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDev.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DIDev.Acquire
End Sub

'
' returns true if the key passed is being pressed
'
Public Function IsDown(keyCode As Byte) As Boolean
    Dim diState As DIKEYBOARDSTATE
    DIDev.GetDeviceStateKeyboard diState
    IsDown = (diState.Key(keyCode) <> 0)
End Function

'
' unaquires keyboard and destroys directInput
'
Private Sub Class_Terminate()
    'Unaquire keyboard
    DIDev.Unacquire
    'Destroy objects
    Set DIDev = Nothing
    Set DI = Nothing
    Set DX = Nothing
End Sub
