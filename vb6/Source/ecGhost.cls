VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecGhost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const GHOST_ACCEL As Single = 0.05
Private Const GHOST_SPEED As Single = 1.5


Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single
Private bScary As Boolean


Public Sub CreateAt(x As Long, y As Long)
    xPos = CSng(x)
    yPos = CSng(y)
End Sub



Public Sub HandleMe()
Dim xSrc As Single
Dim i As Long

    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 _
    Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then Exit Sub
    
    
    bScary = (Mario.isFacingRight And xPos < Mario.xPos) Or _
             (Mario.isFacingRight = False And xPos > Mario.xPos)
    
    If bScary Then
    
        If xPos < Mario.xPos Then xSpeed = xSpeed + GHOST_ACCEL
        If xSpeed > GHOST_SPEED Then xSpeed = GHOST_SPEED
        If xPos > Mario.xPos Then xSpeed = xSpeed - GHOST_ACCEL
        If xSpeed < -GHOST_SPEED Then xSpeed = -GHOST_SPEED
        If yPos < Mario.yPos Then ySpeed = ySpeed + GHOST_ACCEL
        If ySpeed > GHOST_SPEED Then ySpeed = GHOST_SPEED
        If yPos > Mario.yPos Then ySpeed = ySpeed - GHOST_ACCEL
        If ySpeed < -GHOST_SPEED Then ySpeed = -GHOST_SPEED
        
    Else
    
        If xSpeed > 0.2 Then
        xSpeed = xSpeed - GHOST_ACCEL
        ElseIf xSpeed < -0.2 Then
        xSpeed = xSpeed + GHOST_ACCEL
        Else
        xSpeed = 0
        End If
        If ySpeed > 0.2 Then
        ySpeed = ySpeed - GHOST_ACCEL
        ElseIf ySpeed < -0.2 Then
        ySpeed = ySpeed + GHOST_ACCEL
        Else
        ySpeed = 0
        End If
        
    End If
    
    xPos = xPos + xSpeed
    yPos = yPos + ySpeed

    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 14 And .yPos < yPos + IIf(Not .isTall, 32, 52) + 14 Then
            If Not .bInvinsible Then .Injure
        End If
    End If
    End With
        
    xSrc = 0
    If xPos > Mario.xPos Then xSrc = 64
    If bScary Then xSrc = xSrc + 32
    GFX.DrawSurface surfList.EnemyList.LittleBoo, xSrc, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 16
   
End Sub


