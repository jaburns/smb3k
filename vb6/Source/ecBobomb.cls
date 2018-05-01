VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecBobomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TICKING_TIME As Long = 300
Private Const EXPLODE_TIME As Long = 10

Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single

Private lFrame As Long
Private lMaxFrame As Long
Private bFrameOn As Boolean
Private bFaceRight As Boolean

Private bAlive As Boolean
Private bDying As Boolean
Private bStomped As Boolean
Private lStompTime As Long
Private bBlowingUp As Boolean

Private bWatchForEdges As Boolean


Public Sub CreateAt(X As Long, Y As Long, Optional ByVal bSmartFall As Boolean = False)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    xSpeed = -1
    ySpeed = 0
    bFrameOn = False
    lFrame = 0
    bAlive = True
    bDying = False
    bStomped = False
    bBlowingUp = False
    lStompTime = 0
    bWatchForEdges = bSmartFall
End Sub



Public Sub HandleMe()
Dim tTile As udeLTileType
Dim i As Long
Dim X As Long
Dim Y As Long
    
    If Not bAlive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CSng(xInit), CSng(yInit), bWatchForEdges
        Exit Sub
    End If
    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
        If xPos <> xInit And yPos <> yInit Then bAlive = False
        Exit Sub
    End If
    
    If bBlowingUp Then
        If lStompTime = 0 Then
            PlaySound Sounds.BobombExplode
            If findDist(Mario.xPos, Mario.yPos, xPos, yPos) <= 64 Or findDist(Mario.xPos, Mario.yPos - 32, xPos, yPos) <= 64 Then Mario.Injure
            For X = -2 To 2
                For Y = -2 To 2
                    If Not (Abs(X) = 2 And Abs(Y) = 2) Then
                        HitSpecialBlock xPos - (32 * X), (yPos - 16) + (32 * Y), IIf(X <= 0, True, False)
                    End If
                Next Y
            Next X
        End If
        lStompTime = lStompTime + 1
        If lStompTime > EXPLODE_TIME Then
            bAlive = False
            bBlowingUp = False
            Exit Sub
        End If
        If lStompTime Mod 2 Then GFX.DrawSurface surfList.EnemyList.Bobomb, 128, 0, 64, 64, xPos - screenLeft - 64, yPos - screenTop - 80, 128, 128
        Exit Sub
    End If
    
    lFrame = lFrame + 1
    lMaxFrame = 5
    If bStomped Then
        If lStompTime > (TICKING_TIME * 0.8) Then
            lMaxFrame = 1
        ElseIf lStompTime > (TICKING_TIME * 0.5) Then
            lMaxFrame = 3
        Else
            lMaxFrame = 6
        End If
    End If
    If lFrame >= lMaxFrame Then
        lFrame = 0
        bFrameOn = Not bFrameOn
    End If
    If Not bDying Then bFaceRight = (xSpeed > 0)

    If bDying Then
        If bStomped Then
            lStompTime = lStompTime + 1
            If lStompTime > TICKING_TIME Then
                lStompTime = 0
                bDying = False
                bBlowingUp = True
                Exit Sub
            End If
            With Mario
                If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
                    If .yPos >= yPos - 32 And .yPos < yPos - 20 Then
                        PlaySound Sounds.BumpOff
                        .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
                    ElseIf .yPos >= yPos - 20 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
                        xSpeed = .xSpeed 'If .xSpeed > 0 Then xSpeed = 5 Else xSpeed = -5
                        ySpeed = -4
                        .makeRun 0
                    End If
                End If
            End With
            For i = 1 To GetShellCount()
                If GetShellActive(i) Then
                    If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos) < 20 Then KillMe
                End If
            Next i
            MoveMe
            GFX.DrawSurface surfList.EnemyList.Bobomb, IIf(bFaceRight, 0, 64) + IIf(bFrameOn, 32, 0), 32, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        Else
            ySpeed = ySpeed + 0.3
            If ySpeed > 7 Then ySpeed = 7
            yPos = yPos + ySpeed
            If yPos > screenTop + 512 Then
                bAlive = False
                bDying = False
                Exit Sub
            End If
            GFX.DrawSurface surfList.EnemyList.Bobomb, IIf(bFaceRight, 0, 64), 32, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        End If
        Exit Sub
    End If
    
    MoveMe
           
    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 32 And .yPos < yPos - 20 Then
            If .bInvinsible Then
                KillMe
            Else
                .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
                KillMe True
            End If
        ElseIf .yPos >= yPos - 20 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
            If .bInvinsible Then KillMe Else .Injure
        End If
    End If
    End With
    
    For i = 0 To 1
        If getMarioHammerActive(i) Then
            If getMarioHammerX(i) > xPos - 16 And getMarioHammerX(i) < xPos + 16 And getMarioHammerY(i) > yPos - 32 And getMarioHammerY(i) < yPos Then KillMe
        End If
        If getMarioFireActive(i) Then
            If getMarioFireX(i) > xPos - 16 And getMarioFireX(i) < xPos + 16 And getMarioFireY(i) > yPos - 32 And getMarioFireY(i) < yPos Then
                killMarioFire i
            End If
        End If
    Next i
    
    For i = 1 To GetShellCount()
        If GetShellActive(i) Then
            If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos) < 32 Then
                KillMe
                DestroyShellIfCarrying (i)
            End If
        End If
    Next i

    GFX.DrawSurface surfList.EnemyList.Bobomb, IIf(bFaceRight, 0, 64) + IIf(bFrameOn, 32, 0), 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    
End Sub



Private Sub KillMe(Optional ByVal bIsStomped As Boolean = False)
    PlaySound Sounds.BumpOff
    bStomped = bIsStomped
    bDying = True
    ySpeed = IIf(bStomped, 0, -4)
    xSpeed = 0
End Sub



Private Sub MoveMe()
Dim tTile As udeLTileType

    xPos = xPos + xSpeed
    If bDying Then
        xSpeed = xSpeed * 0.965
        If (xSpeed < 0.3 And xSpeed > 0) Or (xSpeed > -0.3 And xSpeed < 0) Then xSpeed = 0
    End If
    
    If GetLevelSideWarp() Then
        If xPos <= -1 Then xPos = GetLevelWidth()
        If xPos >= GetLevelWidth() + 1 Then xPos = 0
    Else
        If xPos <= 16 Then 
            xPos = 16
            xSpeed = 1
        End If
        If xPos >= GetLevelWidth() - 16 Then 
            xPos = GetLevelWidth() - 16
            xSpeed = -1
        End If
    End If
    
    If isTileSolid(xPos + 16, yPos - 16) Then
        tTile = GetTileAtPoint(xPos + 16, yPos - 16)
        If GetTileAtPoint(xPos + 16, yPos - 16) = DEADLY Then KillMe
        xSpeed = -1
    ElseIf isTileSolid(xPos - 16, yPos - 16) Then
        tTile = GetTileAtPoint(xPos - 16, yPos - 16)
        If GetTileAtPoint(xPos - 16, yPos - 16) = DEADLY Then KillMe
        xSpeed = 1
    End If
    
Dim sTestX As Single
Dim sTestY As Single
Dim i As Long
    For i = 1 To GetShellCount()
        If GetShellisShelled(i) Then
            sTestX = GetShellX(i) - xPos
            sTestY = Abs(GetShellY(i) - yPos)
            If sTestY <= 32 Then
                If sTestX > 0 And sTestX <= 32 Then xSpeed = -1
                If sTestX < 0 And sTestX >= -32 Then xSpeed = 1
            End If
        End If
    Next i

Dim maxYSpeed As Long
    ySpeed = ySpeed + 0.3
    If GetTileAtPoint(CLng(xPos), CLng(yPos) - 8) = WATER Then maxYSpeed = 3 Else maxYSpeed = 5
    If ySpeed > maxYSpeed Then ySpeed = maxYSpeed
    yPos = yPos + ySpeed
    
    If isTileSolid(CLng(xPos), CLng(yPos), True) Then
        yPos = (__intDiv(yPos , 32) * 32)
        ySpeed = 0
        If GetTileAtPoint(CLng(xPos), CLng(yPos)) = DEADLY Then KillMe
    End If
    
    If bWatchForEdges Then
        If ySpeed = 0 And xSpeed <> 0 Then
            If Not isTileSolid(CLng(xPos) + IIf(xSpeed > 0, 16, -16), CLng(yPos) + 2, True) Then
                If xSpeed > 0 Then xSpeed = -1 Else xSpeed = 1
            End If
        End If
    End If
    
End Sub


Public Function isInBlastRadius(sX As Single, sY As Single) As Boolean
    isInBlastRadius = (findDist(sX, sY, xPos, yPos) <= 64) And bBlowingUp
End Function

