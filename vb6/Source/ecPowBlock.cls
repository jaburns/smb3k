VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecPowBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private ySpeed As Single
Private xSpeed As Single


Private bAlive As Boolean
Private lDeadTime As Long
Private bCarrying As Boolean




Public Sub CreateAt(X As Long, Y As Long)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    xSpeed = 0
    ySpeed = 0
    lDeadTime = 0
    bAlive = True
    bCarrying = False
End Sub



Public Sub HandleMe()
Dim tTile As udeLTileType

    If Not bAlive Then Exit Sub
    
    If lDeadTime Then
        lDeadTime = lDeadTime + 1
        If lDeadTime > 20 Then bAlive = False
        GFX.DrawSurface surfList.EnemyList.PowButton, 32, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        Exit Sub
    End If
    
    If bCarrying Then
        bCarrying = GameKeyDown(Run)
        xPos = Mario.xPos + IIf(Mario.isFacingRight, 28, -28)
        yPos = Mario.yPos - 2
        If Not bCarrying Then
            Mario.bHandsFull = False
            xSpeed = Mario.xSpeed
            ySpeed = IIf(GameKeyDown(Up), -10, 0)
            PlaySound Sounds.Kick
            If isTileSolid(CLng(xPos), yPos - 16) Then
                xPos = xPos + IIf(Mario.isFacingRight, -64, 64)
                xSpeed = -xSpeed
            End If
        End If
        GFX.DrawSurface surfList.EnemyList.PowButton, 0, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        Exit Sub
    End If
    
    MoveMe
    
    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 32 And .yPos < yPos - 20 Then
            PlaySound Sounds.Bump
            .InitiatePow
            .makeJump 0
            lDeadTime = 1
        ElseIf .yPos >= yPos - 20 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
            If GameKeyDown(Run) And Not .bHandsFull Then
                bCarrying = True
                .bHandsFull = True
            Else
                If .xPos < xPos Then .xPos = xPos - 32 Else .xPos = xPos + 32
                .makeRun 0
            End If
        End If
    End If
    End With
    
    GFX.DrawSurface surfList.EnemyList.PowButton, 0, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    
End Sub





Private Sub MoveMe()

    xPos = xPos + xSpeed
    
    If GetLevelSideWarp Then
        If xPos <= -1 Then xPos = GetLevelWidth
        If xPos >= GetLevelWidth + 1 Then xPos = 0
    Else
        If xPos <= 16 Then xPos = 16: xSpeed = 0
        If xPos >= GetLevelWidth - 16 Then xPos = GetLevelWidth - 16: xSpeed = 0
    End If
    
    If isTileSolid(xPos + 15, yPos - 16) Then
        xSpeed = 0
        xPos = ((xPos \ 32) * 32) + 16
    ElseIf isTileSolid(xPos - 15, yPos - 16) Then
        xSpeed = 0
        xPos = ((xPos \ 32) * 32) + 16
    End If
    
    ySpeed = ySpeed + 0.4
    If ySpeed > 6 Then ySpeed = 6
    yPos = yPos + ySpeed
    
    If isTileSolid(CLng(xPos), CLng(yPos), True) Then
        yPos = ((yPos \ 32) * 32)
        ySpeed = 0
        xSpeed = 0
    End If
    
    If ySpeed < 0 And isTileSolid(CLng(xPos), CLng(yPos) - 32) Then
        yPos = ((yPos \ 32) * 32) + 32
        ySpeed = Abs(ySpeed)
    End If
        
End Sub

