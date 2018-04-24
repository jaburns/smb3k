VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DXMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DX As New DirectX7
Private DI As DirectInput
Private DIDev As DirectInputDevice

Private mX As Single
Private mY As Single
Private mSpeed As Single
Private buttonState(3) As Boolean

Private mBound_L As Long
Private mBound_T As Long
Private mBound_R As Long
Private mBound_B As Long

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Public Property Get X() As Single
    X = mX
End Property
Public Property Get Y() As Single
    Y = mY
End Property
Public Property Get ButtonDown(lButton As Long) As Single
    If lButton < 0 Or lButton > 3 Then Exit Property
    ButtonDown = buttonState(lButton)
End Property

Public Sub MoveTo(newX As Single, newY As Single)
    mX = newX
    mY = newY
End Sub


Public Sub SetMouseVisibile(bVisible As Boolean)
Dim lRet As Long
    Do
        lRet = ShowCursor(bVisible)
    Loop Until IIf(bVisible, lRet >= 0, lRet <= -1)
End Sub


Public Function Initialize(hWnd As Long, Left As Long, Top As Long, Right As Long, Bottom As Long, Optional Sensitivity As Single = 1) As Boolean
On Error GoTo errOut

    Set DI = DX.DirectInputCreate()
    Set DIDev = DI.CreateDevice("GUID_SysMouse")
    
    With DIDev
        .SetCommonDataFormat DIFORMAT_MOUSE
        .SetCooperativeLevel hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
        .Acquire
    End With
    
    mSpeed = Sensitivity
    mBound_L = Left
    mBound_T = Top
    mBound_R = Right
    mBound_B = Bottom
    mX = CSng(Left)
    mY = CSng(Top)
    
Initialize = True
errOut:
End Function


Public Sub Destroy()
    DIDev.Unacquire
    Set DIDev = Nothing
    Set DI = Nothing
    Set DX = Nothing
End Sub


Public Sub Update()
On Error GoTo errOut
Dim mState As DIMOUSESTATE
Dim i As Long

    DIDev.GetDeviceStateMouse mState
       
    mX = mX + mState.X * mSpeed
    If mX < mBound_L Then mX = mBound_L
    If mX > mBound_R Then mX = mBound_R
    
    mY = mY + mState.Y * mSpeed
    If mY < mBound_T Then mY = mBound_T
    If mY > mBound_B Then mY = mBound_B
    
    For i = 0 To 3
        buttonState(i) = (mState.buttons(i) <> 0)
    Next i
        
errOut:
End Sub

Private Sub Class_Terminate()
Destroy
End Sub
