VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocBrickBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos(3) As Single
Private yPos(3) As Single
Private ySpeed As Single
Private bActive As Boolean
Private sAngle As Single

Public Sub CreateAt(x As Single, y As Single)
Dim i As Long
    bActive = True
    For i = 0 To 3
        xPos(i) = x
        yPos(i) = y
    Next i
    ySpeed = -15
End Sub

Public Sub HandleBrick()
Dim bTest As Boolean
Dim mX As Single
Dim mY As Single
Dim i As Long

    If Not bActive Then Exit Sub

    sAngle = sAngle + 10

    xPos(0) = xPos(0) - 1.75
    xPos(1) = xPos(1) + 1.75
    xPos(2) = xPos(2) - 1.75
    xPos(3) = xPos(3) + 1.75
    
    ySpeed = ySpeed + 1
    yPos(0) = yPos(0) + ySpeed
    yPos(1) = yPos(1) + ySpeed
    yPos(2) = yPos(2) + (ySpeed / IIf(ySpeed < 0, 2, 1))
    yPos(3) = yPos(3) + (ySpeed / IIf(ySpeed < 0, 2, 1))
    
    bTest = True
    For i = 0 To 3
        bTest = bTest And ((xPos(i) - screenLeft < 0) Or (xPos(i) - screenLeft > 640) Or (yPos(i) - screenTop < 0) Or (yPos(i) - screenTop > 480))
    Next i
    If bTest Then bActive = False

    For i = 0 To 3
        GFX.DrawSurface surfList.Objects, 0, 64, 16, 16, xPos(i) - screenLeft, yPos(i) - screenTop, , , sAngle
    Next i
    
End Sub
