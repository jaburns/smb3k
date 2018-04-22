VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocFireball"
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
Private lDeadCount As Single

Public Property Get isActive() As Boolean
isActive = bActive And Not bDying
End Property
Public Property Get getXPos() As Single
getXPos = xPos
End Property
Public Property Get getYPos() As Single
getYPos = yPos
End Property
Public Sub DestroyMe()
bDying = True
lDeadCount = 1
End Sub

Public Sub CreateAt(X As Single, Y As Single, bFaceRight As Boolean)
    xPos = X
    yPos = Y
    If bFaceRight Then xSpeed = 7 Else xSpeed = -7
    ySpeed = 1
    bActive = True
    bDying = False
End Sub

Public Sub HandleBall()
    
    If Not bActive Then Exit Sub
    If bDying Then
        lDeadCount = lDeadCount + (lDeadCount)
        If lDeadCount > 20 Then bActive = False
        GFX.DrawSurface surfList.Objects, 20, 167, 16, 16, (xPos - screenLeft) - 8 - lDeadCount, (yPos - screenTop) - 8 - lDeadCount, 16 + (2 * lDeadCount), 16 + (2 * lDeadCount), sRotation, 100 * ((20 - lDeadCount) / 20)
        Exit Sub
    End If
    
    If isTileSolid(CLng(xPos) + 8, CLng(yPos)) Then
        TestToDestroyBlock xPos + 8, yPos
        DestroyMe
    ElseIf isTileSolid(CLng(xPos) - 8, CLng(yPos)) Then
        TestToDestroyBlock xPos - 8, yPos
        DestroyMe
    End If
            
    If ySpeed > 0 Then
        If isTileSolid(CLng(xPos), CLng(yPos) + 8, True) Then
            If TestToDestroyBlock(xPos, yPos + 8) Then DestroyMe
            yPos = ((yPos \ 32) * 32) + 24
            ySpeed = -12
        End If
    Else
        If isTileSolid(CLng(xPos), CLng(yPos) - 8) Then
            If TestToDestroyBlock(xPos, yPos - 8) Then DestroyMe
            yPos = ((yPos \ 32) * 32) + 8
            If ySpeed < 0 Then ySpeed = -ySpeed
        End If
    End If
    
    xPos = xPos + xSpeed
    ySpeed = ySpeed + 1
    If ySpeed > 7 Then ySpeed = 7
    yPos = yPos + ySpeed
    
    If xPos < screenLeft Or xPos > screenLeft + 640 Or yPos > screenTop + 480 Then
        bActive = False
    End If
    
    sRotation = sRotation + (xSpeed * 4)
    If sRotation >= 360 Then sRotation = sRotation - 360
    If sRotation < 0 Then sRotation = sRotation + 360
    
    GFX.DrawSurface surfList.Objects, 0, 93, 16, 16, (xPos - screenLeft) - 8, (yPos - screenTop) - 8, , , sRotation

End Sub


Private Function TestToDestroyBlock(sX As Single, sY As Single) As Boolean
    TestToDestroyBlock = False
    If GetTileEnemy(sX \ 32, sY \ 32) = OBJFIRETILE Then
        KillTile CLng(sX), CLng(sY)
        BreakBrick (sX \ 32) * 32 + 16, (sY \ 32) * 32 + 16
        PlaySound Sounds.BreakBrick
        TestToDestroyBlock = True
    End If
End Function


