VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocPowerStar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single
Private growVal As Long
Private bActive As Boolean


Public Sub CreateAt(X As Single, Y As Single, bFaceRight As Boolean)
    xPos = X
    yPos = Y
    If bFaceRight Then xSpeed = 2 Else xSpeed = -2
    ySpeed = 0.01
    growVal = 1
    bActive = True
    PlaySound Sounds.Sprout
End Sub

Public Sub HandleStar()

    If Not bActive Then Exit Sub
    
    If growVal < 32 Then
        growVal = growVal + 1
        GFX.DrawSurface surfList.Sprites, 64, 0, 32, growVal, (xPos - screenLeft) - 16, ((yPos - screenTop) - growVal)
        Exit Sub
    End If
    
    If GetLevelSideWarp Then
        If xPos <= -1 Then xPos = GetLevelWidth()
        If xPos >= GetLevelWidth() + 1 Then xPos = 0
    Else
        If xPos <= 16 Then 
            xPos = 16
            xSpeed = 2
        End If
        If xPos >= GetLevelWidth() - 16 Then 
            xPos = GetLevelWidth() - 16
            xSpeed = -2
        End If
    End If

    xPos = xPos + xSpeed
    
    If isTileSolid(xPos + 16, yPos - 16) Then xSpeed = -2
    If isTileSolid(xPos - 16, yPos - 16) Then xSpeed = 2
    
    ySpeed = ySpeed + 0.2
    If ySpeed > 5 Then ySpeed = 5
    yPos = yPos + ySpeed
    
    If ySpeed > 0 Then
        If isTileSolid(CLng(xPos), CLng(yPos), True) Then
            yPos = (__intDiv(yPos , 32) * 32)
            ySpeed = -7
        End If
    Else
        If isTileSolid(CLng(xPos), CLng(yPos) - 32) Then
            yPos = (__intDiv(yPos , 32) * 32) + 32
            If ySpeed < 0 Then ySpeed = -ySpeed
        End If
    End If
    
    If Abs(xPos - Mario.xPos) < 32 And Abs(yPos - Mario.yPos) < 32 And Mario.bAlive Then
        bActive = False
        Mario.addShield 25, True
        Exit Sub
    End If
    
    GFX.DrawSurface surfList.Sprites, 64, 0, 32, 32, (xPos - screenLeft) - 16, (yPos - screenTop) - 32

End Sub


