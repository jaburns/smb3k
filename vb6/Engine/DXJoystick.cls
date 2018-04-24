VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXJoystick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DirectX As New DirectX7
Private DInput As DirectInput
Private DIDev As DirectInputDevice
Private diDevEnum As DirectInputEnumDevices

Private bJoystickPresent As Boolean

Private joyCaps As DIDEVCAPS

Private iJoyX As Integer
Private iJoyY As Integer
Private iButtonCount As Integer
Private bButton(31) As Boolean
'

Public Property Get ButtonCount() As Integer
    ButtonCount = iButtonCount
End Property

Public Property Get JoystickPresent() As Boolean
    JoystickPresent = bJoystickPresent
End Property

Public Property Get ButtonDown(ByVal iButton As Integer) As Boolean

    ButtonDown = False
    If Not bJoystickPresent Then Exit Property
    If iButton > iButtonCount Or iButton < 1 Then Exit Property
    
    ButtonDown = bButton(iButton - 1)
    
End Property

Public Property Get XAxis() As Integer
    
    XAxis = 0
    If Not bJoystickPresent Then Exit Property
    
    XAxis = iJoyX
    
End Property

Public Property Get YAxis() As Integer
    
    YAxis = 0
    If Not bJoystickPresent Then Exit Property
    
    YAxis = iJoyY
    
End Property


Public Sub Initialize(hWnd As Long)
On Error GoTo errOut:
Dim DiProp_Dead As DIPROPLONG
Dim DiProp_Range As DIPROPRANGE
Dim DiProp_Saturation As DIPROPLONG

bJoystickPresent = False
       
    Set DInput = DirectX.DirectInputCreate()
    Set diDevEnum = DInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)

    If diDevEnum.GetCount = 0 Then GoTo errOut
    
    Set DIDev = DInput.CreateDevice(diDevEnum.GetItem(1).GetGuidInstance)
    DIDev.SetCommonDataFormat DIFORMAT_JOYSTICK
    DIDev.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DIDev.GetCapabilities joyCaps
    
    With DiProp_Dead
        .lData = 1000
        .lObj = DIJOFS_X
        .lSize = Len(DiProp_Dead)
        .lHow = DIPH_BYOFFSET
        .lObj = DIJOFS_X
        DIDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_Y
        DIDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End With
    
    With DiProp_Saturation
        .lData = 9500
        .lHow = DIPH_BYOFFSET
        .lSize = Len(DiProp_Saturation)
        .lObj = DIJOFS_X
         DIDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_Y
         DIDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
    End With
    
    With DiProp_Range
        .lHow = DIPH_DEVICE
        .lSize = Len(DiProp_Range)
        .lMin = 0
        .lMax = 10000
    End With
    DIDev.SetProperty "DIPROP_RANGE", DiProp_Range
    DIDev.Acquire
    
    iButtonCount = joyCaps.lButtons
    
bJoystickPresent = True

errOut:
End Sub


Public Sub Update()
On Error Resume Next
Dim Joystick As DIJOYSTATE
Dim i As Long

    If Not bJoystickPresent Then Exit Sub
    
    DIDev.Poll
    DIDev.GetDeviceStateJoystick Joystick
        
    iJoyX = Joystick.X - 5000
    iJoyY = Joystick.Y - 5000
    
    For i = 0 To iButtonCount - 1
        If Joystick.buttons(i) = 0 Then bButton(i) = False Else bButton(i) = True
    Next

End Sub


Private Sub Class_Terminate()
On Error Resume Next
    DIDev.Unacquire
    Set DirectX = Nothing
    Set DInput = Nothing
    Set DIDev = Nothing
    Set diDevEnum = Nothing
End Sub
