VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecKoopa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const SHELL_SPEED As Long = 6
Private Const WALK_SPEED As Long = 1

Private lShellSpeed As Long

Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single
Private curSpeed As Single

Private fileSrc As Long
Private SrcHeight As Long

Private lFrame As Long
Private bFrameOn As Boolean
Private lShellFrame As Long

Private bAlive As Boolean
Private bDying As Boolean
Private bShell As Boolean
Private bCarrying As Boolean
Private kStyle As udeKoopaStyle
Private bAlwaysShell As Boolean
Private bDefaultBouncy As Boolean
Private bBouncy As Boolean
Private lBounceFrame As Long

Private bWatchForEdges As Boolean

'Private myHammer As ocThrowHammer '----------------------------------------------------

Public Property Get xLoc() As Single
xLoc = xPos
End Property
Public Property Get yLoc() As Single
yLoc = yPos
End Property
Public Property Get isShreddin() As Boolean
If bDying Then Exit Property
isShreddin = bCarrying Or (bShell And ((xSpeed <> 0) Or (ySpeed <> 0)) And bAlive)
End Property
Public Property Get isShelled() As Boolean
If bDying Then Exit Property
isShelled = (Not bCarrying) And (bShell And (xSpeed = 0) And bAlive)
End Property
Public Property Get isAlive() As Boolean
    isAlive = bAlive
End Property
Public Sub killShellIfCarrying()
    If bCarrying Then
        KillMe
        Mario.bHandsFull = False
    End If
End Sub

Public Sub setShellSpeed(lNewShellSpeed As Long)
    lShellSpeed = lNewShellSpeed
End Sub



Public Sub createShell(X As Long, Y As Long, inXSpeed As Single, inYSpeed As Single)
    CreateAt X, Y, KOOPA_GREEN, True, False
    xSpeed = inXSpeed
    ySpeed = inYSpeed
End Sub


Public Sub CreateAt(X As Long, Y As Long, koopaStyle As udeKoopaStyle, bDefaultShell As Boolean, Optional ByVal isBouncy As Boolean = False, Optional ByVal bSmartFall As Boolean = False)

    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    
    If lShellSpeed = 0 Then lShellSpeed = SHELL_SPEED
    bAlwaysShell = bDefaultShell
    bShell = bAlwaysShell
    kStyle = koopaStyle
    Select Case kStyle
        Case KOOPA_GREEN:  SrcHeight = 64: fileSrc = surfList.EnemyList.DumbKoopa
        Case KOOPA_RED:  SrcHeight = 64: fileSrc = surfList.EnemyList.SmartKoopa
        Case KOOPA_BEETLE: SrcHeight = 32: fileSrc = surfList.EnemyList.BuzzyBeetle
        Case KOOPA_BOUNCER: SrcHeight = 32: fileSrc = surfList.EnemyList.BouncyKoopa
    End Select
    
    curSpeed = IIf(bShell, lShellSpeed, WALK_SPEED)
    xSpeed = IIf(bShell, 0, -WALK_SPEED)
    ySpeed = 0
    lBounceFrame = 0
    
    bDefaultBouncy = isBouncy
    bBouncy = bDefaultBouncy
    bFrameOn = False
    lFrame = 0
    bAlive = True
    bDying = False
    bCarrying = False
    bWatchForEdges = bSmartFall

End Sub



Public Sub HandleMe()
Dim tTile As udeLTileType
Dim sXSrc As Single
Dim sYSrc As Single
Dim i As Long
    
    If Not bAlive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 _
        Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CLng(xInit), CLng(yInit), kStyle, bAlwaysShell, bDefaultBouncy, bWatchForEdges
        Exit Sub
    End If
    If Not bCarrying Then
        If xPos < screenLeft - 128 Or xPos > screenLeft + 768 _
        Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
            If xPos <> xInit Or yPos <> yInit Then bAlive = bCarrying
            Exit Sub
        End If
    End If
    
    If bDying Then
        ySpeed = ySpeed + 0.3
        If ySpeed > 7 Then ySpeed = 7
        yPos = yPos + ySpeed
        If yPos > screenTop + 512 Then
            bAlive = False
            bDying = False
            Exit Sub
        End If
        GFX.DrawSurface fileSrc, IIf(kStyle = KOOPA_BOUNCER, 192, 128), 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        Exit Sub
    End If
    
    If bCarrying Then
        bCarrying = GameKeyDown(Run)
        xPos = Mario.xPos + IIf(Mario.isFacingRight, 28, -28)
        yPos = Mario.yPos - 2
        If Not bCarrying Then
            Mario.bHandsFull = False
            PlaySound Sounds.Kick
            If GameKeyDown(Up) Then
                xSpeed = Mario.xSpeed
                ySpeed = -10
                yPos = yPos - 12
            Else
                xSpeed = IIf(Mario.isFacingRight, lShellSpeed, -lShellSpeed)
                ySpeed = 0
            End If
            xPos = xPos + IIf(Mario.isFacingRight, 8, -8)
            If isTileSolid(CLng(xPos), yPos - 16) Then
                xPos = xPos + IIf(Mario.isFacingRight, -64, 64)
                xSpeed = -xSpeed
            End If
            xPos = xPos + xSpeed * 2
            yPos = yPos + ySpeed * 2
        End If
        If kStyle = KOOPA_BOUNCER Then
            GFX.DrawSurface fileSrc, 128, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        Else
            GFX.DrawSurface fileSrc, 0, CSng(SrcHeight), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        End If
        Exit Sub
    End If
    
    For i = 1 To GetShellCount
        If GetShellActive(i) And Not (GetShellX(i) = xPos And GetShellY(i) = yPos) Then
            If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos) < 20 Then
                KillMe
                DestroyShellIfCarrying (i)
            End If
        End If
    Next i
    If kStyle <> KOOPA_BEETLE Then If inBombBlast(xPos, yPos - 16) Then KillMe
        
    If bShell Then HandleShell Else HandleWalker
    
End Sub



Private Sub HandleWalker()
Dim i As Long

'-HAMMER KOOPA MAYHEM------------------------------------------------------------------
'    With myHammer
'    If Not .isActive Then
'        .CreateAt xPos, yPos, xSpeed * 4
'    Else
'        If Abs(.getXPos - Mario.xPos) < 24 Then _
'        If .getYPos < Mario.yPos And .getYPos > Mario.yPos - IIf(Mario.isTall, 52, 32) Then Mario.Injure
'    End If
'    .HandleBall
'    End With
'--------------------------------------------------------------------------------------
    
    lFrame = lFrame + 1
    If lFrame >= 5 Then
        lFrame = 0
        bFrameOn = Not bFrameOn
    End If
    
    MoveMe
    
    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 48 And .yPos < yPos - 30 Then
            If .bInvinsible Then
                KillMe
            Else
                .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
                PlaySound Sounds.HardShell
                ySpeed = 0
                If bBouncy Then
                    bBouncy = False
                Else
                    xSpeed = 0
                    bShell = True
                    curSpeed = lShellSpeed
                End If
            End If
        ElseIf .yPos >= yPos - 30 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
            If .bInvinsible Then KillMe Else .Injure
        End If
    End If
    End With

    For i = 0 To 1
        If getMarioHammerActive(i) Then
            If getMarioHammerX(i) > xPos - 16 And getMarioHammerX(i) < xPos + 16 _
            And getMarioHammerY(i) > yPos - 32 And getMarioHammerY(i) < yPos Then KillMe
        End If
        If getMarioFireActive(i) Then
            If getMarioFireX(i) > xPos - 16 And getMarioFireX(i) < xPos + 16 _
            And getMarioFireY(i) > yPos - 32 And getMarioFireY(i) < yPos Then
                killMarioFire i
                If kStyle <> KOOPA_BEETLE Then KillMe
            End If
        End If
    Next i
    
    GFX.DrawSurface fileSrc, IIf(xSpeed > 0, 0, 64) + IIf(bFrameOn, 32, 0), 0, 32, CSng(SrcHeight), xPos - screenLeft - 16, yPos - screenTop - IIf(kStyle = KOOPA_BEETLE Or kStyle = KOOPA_BOUNCER, 32, 64)
    If bBouncy Then
        If xSpeed > 0 Then
            GFX.DrawSurface surfList.MarioWings, IIf(bFrameOn, 32, 0), 0, 32, 32, xPos - screenLeft - 32, yPos - screenTop - 48
        Else
            GFX.DrawSurface surfList.MarioWings, IIf(bFrameOn, 32, 0), 32, 32, 32, xPos - screenLeft, yPos - screenTop - 48
        End If
    End If

End Sub








Private Sub HandleShell()
Dim i As Long

    If kStyle = KOOPA_BOUNCER And lBounceFrame > 0 Then lBounceFrame = lBounceFrame - 1

    lFrame = lFrame + 1
    If lFrame >= 2 Then
        lFrame = 0
        lShellFrame = lShellFrame + IIf(xSpeed > 0, -1, 1)
        If lShellFrame >= 4 Then lShellFrame = 0
        If lShellFrame <= -1 Then lShellFrame = 3
    End If
    
    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 32 And .yPos < yPos - 20 Then
            PlaySound Sounds.HardShell
            .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
            If xSpeed Then
                xSpeed = 0
            ElseIf .xPos < xPos - 8 Then
                xSpeed = curSpeed
            ElseIf .xPos > xPos + 8 Then
                xSpeed = -curSpeed
            Else
                xSpeed = 0
            End If
        ElseIf .yPos >= yPos - 20 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
            If xSpeed = 0 Then
                If .bInvinsible Then
                    KillMe
                ElseIf GameKeyDown(Run) And Not Mario.bHandsFull Then
                    bCarrying = True
                    Mario.bHandsFull = True
                ElseIf ySpeed = 0 Then
                    PlaySound Sounds.Kick
                    If xPos > .xPos Then xSpeed = curSpeed Else xSpeed = -curSpeed
                    .makeRun 0
                End If
            Else
                If .bInvinsible Then KillMe Else .Injure
            End If
        End If
    End If
    End With
    
    MoveMe

    For i = 0 To 1
        If getMarioHammerActive(i) Then
            If getMarioHammerX(i) > xPos - 16 And getMarioHammerX(i) < xPos + 16 _
            And getMarioHammerY(i) > yPos - 32 And getMarioHammerY(i) < yPos Then KillMe
        End If
        If getMarioFireActive(i) Then
            If getMarioFireX(i) > xPos - 16 And getMarioFireX(i) < xPos + 16 _
            And getMarioFireY(i) > yPos - 32 And getMarioFireY(i) < yPos Then
                killMarioFire i
                If kStyle <> KOOPA_BEETLE Then KillMe
            End If
        End If
    Next i
    
    If kStyle = KOOPA_BOUNCER Then
        GFX.DrawSurface fileSrc, IIf(lBounceFrame = 0, 128, 160), 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    Else
        GFX.DrawSurface fileSrc, IIf(xSpeed, lShellFrame * 32, 0), CSng(SrcHeight), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    End If

End Sub


Public Sub KillMe()
    If bCarrying Then Mario.bHandsFull = False
    PlaySound Sounds.BumpOff
    bDying = True
    ySpeed = -3
End Sub

Private Sub MoveMe()
Dim tTile As udeLTileType

    xPos = xPos + xSpeed
    
    If GetLevelSideWarp Then
        If xPos <= -1 Then xPos = GetLevelWidth
        If xPos >= GetLevelWidth + 1 Then xPos = 0
    Else
        If xPos <= 16 Then xPos = 16: xSpeed = curSpeed
        If xPos >= GetLevelWidth - 16 Then xPos = GetLevelWidth - 16: xSpeed = -curSpeed
    End If
    
    If isTileSolid(xPos + 15, yPos - 16) Then
        If GetTileAtPoint(xPos + 15, yPos - 16) = DEADLY Then KillMe
        If bShell Then HitSpecialBlock xPos + 16, yPos - 16, True
        If kStyle = KOOPA_BOUNCER Then xSpeed = -Abs(xSpeed) Else xSpeed = -curSpeed
    ElseIf isTileSolid(xPos - 15, yPos - 16) Then
        If GetTileAtPoint(xPos - 15, yPos - 16) = DEADLY Then KillMe
        If bShell Then HitSpecialBlock xPos - 16, yPos - 16, False
        If kStyle = KOOPA_BOUNCER Then xSpeed = Abs(xSpeed) Else xSpeed = curSpeed
    End If
    
Dim sTestX As Single
Dim sTestY As Single
Dim i As Long
    If Not bShell Then
        For i = 1 To GetShellCount
            If GetShellisShelled(i) Then
                sTestX = GetShellX(i) - xPos
                sTestY = Abs(GetShellY(i) - yPos)
                If sTestY <= 32 Then
                    If sTestX > 0 And sTestX <= 32 Then xSpeed = -1
                    If sTestX < 0 And sTestX >= -32 Then xSpeed = 1
                End If
            End If
        Next i
    End If
    
Dim maxYSpeed As Long
    ySpeed = ySpeed + 0.4
    If GetTileAtPoint(CLng(xPos), CLng(yPos) - 8) = WATER Then maxYSpeed = 3 Else maxYSpeed = 5
    If ySpeed > maxYSpeed Then ySpeed = maxYSpeed
    yPos = yPos + ySpeed
    If ySpeed > 0 And isTileSolid(CLng(xPos), CLng(yPos), True) Then
        If Not (kStyle = KOOPA_BOUNCER And bShell) Then
            If xSpeed > 0 Then xSpeed = curSpeed
            If xSpeed < 0 Then xSpeed = -curSpeed
        End If
        yPos = ((yPos \ 32) * 32)
        If kStyle = KOOPA_BOUNCER And bShell Then
            ySpeed = -10
            lBounceFrame = 2
        Else
            ySpeed = IIf(bBouncy And Not bShell, IIf(kStyle = KOOPA_RED, -13, -10), 0)
        End If
        If GetTileAtPoint(CLng(xPos), CLng(yPos)) = DEADLY Then KillMe
    End If
    
    'turn around at a cliff
    If bWatchForEdges And (Not bShell) Then
        If ySpeed = 0 And xSpeed <> 0 Then
            If Not isTileSolid(CLng(xPos) + IIf(xSpeed > 0, 16, -16), CLng(yPos) + 2, True) Then
                If xSpeed > 0 Then xSpeed = -curSpeed Else xSpeed = curSpeed
            End If
        End If
    End If
    
    'test above jumpers
    If bShell Or bBouncy Then
        If bShell Then If hitHiddenBlock(CLng(xPos), yPos - 32) Then PlaySound Sounds.Bump
        If isTileSolid(CLng(xPos), CLng(yPos) - IIf(bShell Or kStyle = KOOPA_BEETLE Or kStyle = KOOPA_BOUNCER, 32, 54)) Then
            If bShell Then HitSpecialBlock CLng(xPos), yPos - 32, xSpeed >= 0
            yPos = ((yPos \ 32) * 32) + 32
            ySpeed = Abs(ySpeed)
            If GetTileAtPoint(CLng(xPos), CLng(yPos) - IIf(bShell Or kStyle = KOOPA_BEETLE Or kStyle = KOOPA_BOUNCER, 32, 54)) = DEADLY Then KillMe
        End If
    End If
    
End Sub

'Private Sub Class_Initialize()
'Set myHammer = New ocThrowHammer
'End Sub

'Private Sub Class_Terminate()
'Set myHammer = Nothing
'End Sub
