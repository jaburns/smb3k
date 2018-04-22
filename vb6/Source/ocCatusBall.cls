VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocCactusBall"
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

Public Property Get isActive() As Boolean
isActive = bActive
End Property

Public Sub CreateAt(x As Single, y As Single, xVelocity As Single)
    bActive = True
    xPos = x
    yPos = y
    xSpeed = xVelocity
    ySpeed = -7
End Sub

Public Sub HandleBall()

    If Not bActive Then Exit Sub
    If yPos > screenTop + 512 Or yPos < screenTop _
    Or xPos < screenLeft - 16 Or xPos > screenLeft + 656 Then bActive = False: Exit Sub

    ySpeed = ySpeed + 0.2
    If ySpeed > 5 Then ySpeed = 5
    xPos = xPos + xSpeed
    yPos = yPos + ySpeed
    
    GFX.DrawSurface surfList.EnemyList.MisterPokey, 64, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32

End Sub
