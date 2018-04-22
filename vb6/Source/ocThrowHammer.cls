VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocThrowHammer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single
Private bActive As Boolean
Private bDying As Boolean
Private sRotation As Single

Public Property Get isActive() As Boolean
isActive = bActive
End Property
Public Property Get getXPos() As Single
getXPos = xPos
End Property
Public Property Get getYPos() As Single
getYPos = yPos
End Property
Public Sub DestroyMe()
bActive = False
End Sub


Public Sub CreateAt(X As Single, Y As Single, sXSpeed As Single)
    xPos = X
    yPos = Y
    xSpeed = sXSpeed * 1.2
    If xSpeed > -2 And xSpeed < 0 Then xSpeed = -2
    If xSpeed < 2 And xSpeed > 0 Then xSpeed = 2
    If xSpeed = 0 Then xSpeed = IIf(Mario.isFacingRight, 2, -2)
    ySpeed = -20
    bActive = True
End Sub

Public Sub HandleBall()
Dim tTile As udeLTileType
Dim xSrc As Long

    If Not bActive Then Exit Sub
    
    xPos = xPos + xSpeed
    ySpeed = ySpeed + 1
    If ySpeed > 9 Then ySpeed = 9
    yPos = yPos + ySpeed
    
    If xPos < screenLeft Or xPos > screenLeft + 640 Or yPos > screenTop + 480 Then bActive = False
    
    sRotation = sRotation + (xSpeed * 4)
    If sRotation >= 360 Then sRotation = sRotation - 360
    If sRotation < 0 Then sRotation = sRotation + 360
    
    GFX.DrawSurface surfList.Objects, 0, 122, 16, 32, (xPos - screenLeft) - 8, (yPos - screenTop) - 16, , , sRotation
   ' If xSpeed <= -8.5 Or xSpeed > 8.5 Then HitSpecialBlock CLng(xPos), CLng(yPos), xSpeed > 0
    
End Sub


