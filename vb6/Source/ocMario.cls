VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocMario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FIRING_FREQ As Long = 25

Public xPos As Single
Public yPos As Single

Public xSrc As Long
Public ySrc As Long

Public SrcHeight As Long
Public SrcWidth As Long

Public bAlive As Boolean
Public mStatus As udeMarioStatus
Public bStarFrameOn As Boolean
Public bHasWon As Boolean

Public bHandsFull As Boolean

Public xOnScreen As Long
Public yOnScreen As Long

Private lPowTime As Long
Private lStarTime As Long
Private bStarOffence As Boolean
Private lBlueStarTime As Long
Private lWingFrameTime As Long
Private lEarFlapTime As Long
Private bEarFlapFrame As Boolean
Private sStarAngle As Single

Private lPipeStat As udeMarioPipeStat
Private lPipeValue As Long
Private lPipeValueMax As Long
Private lPipeUsed As Long

Private lLastFire As Long

Private frameCount As Long
Private curRunFrame As Long
Private curClimbFrame As Long
Private climbFrameCount As Long

Private sGravity As Single
Private sJump    As Single
Private sRun     As Single

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private lEquipStat As Long
'Private sEquipAngle As Single
'Private sEquipRising As Boolean
'Private lEquipFireStat As Long
'Private sEquipX As Single
'Private sEquipY As Single
'Private sEquipXSpeed As Single
'Private sEquipYSpeed As Single
'Private bShootRight As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private bFloating As Boolean
Private bSwimming As Boolean
Private bClimbing As Boolean
Private bTall As Boolean
Private bMarioFaceRight As Boolean
Private bHoldJump As Boolean
Private bForceStand As Boolean

Private bAlwaysSwimming As Boolean
Private lAlwaysSwimmingTop As Long


Public Property Get isFacingRight() As Boolean
    isFacingRight = bMarioFaceRight
End Property

Public Property Get xSpeed() As Single
    xSpeed = sRun
End Property

Public Property Get ySpeed() As Single
    ySpeed = sGravity
End Property

Public Property Get isPowerful() As Boolean
    isPowerful = lBlueStarTime > 0
End Property

Public Property Get isTall() As Boolean
    isTall = bTall
End Property

Public Property Get isSwimming() As Boolean
    isSwimming = bSwimming
End Property

Public Property Get bInvinsible() As Boolean
    bInvinsible = (lStarTime > 0) And bStarOffence
End Property

Public Property Get bPowSwitchOn() As Boolean
    bPowSwitchOn = lPowTime > 0
End Property

Public Property Get mPipeStatus() As udeMarioPipeStat
    mPipeStatus = lPipeStat
End Property

Public Property Get LastPipeUsed() As Long
    LastPipeUsed = lPipeUsed
End Property

Public Sub ResetXPos(newXPos As Single, newXSpeed As Single)
    sRun = newXSpeed
    xPos = newXPos
End Sub
Public Sub ResetYPos(newYPos As Single, newYSpeed As Single)
    sGravity = newYSpeed
    sJump = 0
    yPos = newYPos
End Sub

Public Sub SetAlwaysSwimming(bIsAlwaysSwimming As Boolean, lWaterTop As Long)
    bAlwaysSwimming = bIsAlwaysSwimming
    lAlwaysSwimmingTop = lWaterTop
End Sub

Public Sub makeExitPipe(lDirection As udeMarioPipeStat)
If lDirection < 4 Then Exit Sub
    lPipeStat = lDirection
    lPipeValue = 0
    lPipeValueMax = 32
    If (lDirection = mpExitDown Or lDirection = mpExitUp) And mStatus <> MarioSmall Then lPipeValueMax = 64
    If lDirection = mpExitDoor Then
        SetTile CLng(xPos), CLng(yPos) - 16, 1, 0, True
        SetTile CLng(xPos), CLng(yPos) - 48, 1, 0, True
    Else
        PlaySound Sounds.Pipe
    End If
End Sub

Public Sub addShield(lSeconds As Long, bAgressive As Boolean)
    If lStarTime And Not bAgressive Then Exit Sub
    lStarTime = lSeconds * 25
    bStarOffence = bAgressive
    If bStarOffence Then
        MusicStopMusic
        MusicLoadFile App.Path & "\Data\Music\Star.mid"
        MusicPlayMusic
    End If
End Sub

Public Sub makeJump(sVal As Single)
    sJump = sVal
    sGravity = 0
End Sub

Public Sub makeStand(sXPos As Single, sYPos As Single)
    bForceStand = True
    xPos = sXPos
    yPos = sYPos
    sGravity = 0
    sJump = 0
End Sub

Public Sub makeRun(sNewRun As Single)
    sRun = sNewRun
End Sub

Public Sub makeDie()
    bAlive = False
    xSrc = 256
End Sub

Public Sub addPowerTime(lSeconds As Long)
    lBlueStarTime = lSeconds * 25
End Sub

Public Sub InitiatePow()
    PlaySound Sounds.PowLoop ', True
    LevelBlockSwitch
    lPowTime = 4
End Sub

Public Sub endPowerTime()
    lBlueStarTime = 0
End Sub


Public Sub Injure()
    If lStarTime Then Exit Sub
    If lPipeStat Then Exit Sub
    
    If lBlueStarTime Then
        lBlueStarTime = 0
        addShield 4, False
        PlaySound Sounds.Powerdown
    ElseIf mStatus = MarioSmall Then
        bAlive = False
        xSrc = 256
    Else
        mStatus = MarioSmall
        addShield 4, False
        PlaySound Sounds.Powerdown
        ReleaseReserveItem
    End If
End Sub





Async Public Sub UpdatePosition()

    If yPos >= GetLevelHeight + 64 Then
        bAlive = False
        xSrc = 256
        yPos = yPos - 32
        Exit Sub
    End If

    HandlePiping
    If lPipeStat Then Exit Sub
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    If lEquipFireStat = 1 Then
'        Do
'        'magnificent mirages
'            'remain beautiful illusions
'            'you can't capture the horizon
'            'it is she who captures you
'            sEquipX = sEquipX + sEquipXSpeed
'            sEquipY = sEquipY + sEquipYSpeed
'            If isTileSolid(CLng(sEquipX), CLng(sEquipY)) Then
'                lEquipFireStat = 2
'                sEquipXSpeed = 0
'                sEquipYSpeed = 0
'                Exit Do
'            ElseIf sEquipX < screenLeft Or sEquipX > screenLeft + 640 Or sEquipY < screenTop Or sEquipY > screenTop + 480 Then
'                lEquipFireStat = 0
'                sEquipXSpeed = 0
'                sEquipYSpeed = 0
'                Exit Do
'            End If
'        Loop
'    ElseIf lEquipFireStat = 2 Then
'        sEquipXSpeed = sEquipXSpeed + 0.005 * (sEquipX - xPos)
'        sEquipXSpeed = sEquipXSpeed * 0.98
'        sEquipYSpeed = sEquipYSpeed + 0.005 * (sEquipY - yPos + 64)
'        sEquipYSpeed = sEquipYSpeed * 0.98
'        If isTileSolid(CLng(xPos), CLng(yPos) - IIf(bTall, 52, 32)) Then sEquipYSpeed = 0.5 - sGravity
'        'If sEquipYSpeed < -10 Then sEquipYSpeed = -10
'        xPos = xPos + sEquipXSpeed
'        yPos = yPos + sEquipYSpeed
'        If GameKeyDown(Jump) Then
'            sRun = sRun + sEquipXSpeed
'            sJump = 5
'            lEquipFireStat = 0
'        End If
'    End If
'    Select Case lEquipStat
'        Case 0
'            If GameKeyDown(Shoot) And lEquipFireStat = 0 Then lEquipStat = 1
'        Case 1
'            'If Not GameKeyDown(Shoot) Then
'                PlaySound Sounds.Sprout
'                lEquipStat = 2
'                bShootRight = bMarioFaceRight
'                lEquipFireStat = 0
'                sEquipAngle = 0
'            'End If
'        Case 2
'            If Not GameKeyDown(Shoot) Then
'                PlaySound Sounds.Fireball
'                sEquipX = xPos
'                sEquipY = yPos - 32
'                sEquipXSpeed = 4 * Cos((sEquipAngle - 90) * PI / 180) * IIf(bShootRight, 1, -1)
'                sEquipYSpeed = 4 * Sin((sEquipAngle - 90) * PI / 180)
'                lEquipStat = 3
'                lEquipFireStat = 1
'            End If
'            If bShootRight <> bMarioFaceRight Then
'                lEquipStat = 3
'            End If
''                xPos = xPos + (IIf(GameKeyDown(Run), 8, 4) * Cos((sEquipAngle - 90) * PI / 180) * IIf(bMarioFaceRight, 1, -1))
' '               yPos = yPos + (IIf(GameKeyDown(Run), 8, 4) * Sin((sEquipAngle - 90) * PI / 180))
'        Case 3
'            'If Not GameKeyDown(Shoot) Then
'            lEquipStat = 0
'    End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    bSwimming = GetTileAtPoint(CLng(xPos), CLng(yPos) - 4) = WATER Or GetTileAtPoint(CLng(xPos), CLng(yPos) - 28) = WATER
    If Not bSwimming And bTall Then bSwimming = GetTileAtPoint(CLng(xPos), CLng(yPos) - 52) = WATER
    If bAlwaysSwimming And (yPos > lAlwaysSwimmingTop) Then bSwimming = True

    If GetTileEnemy(__intDiv(xPos , 32), __intDiv(yPos , 32)) = OBJENDFLYING Then lBlueStarTime = 0
    
    If (GetTileAtPoint(CLng(xPos), CLng(yPos) - 4) = VINE Or GetTileAtPoint(CLng(xPos), CLng(yPos) - IIf(mStatus = MarioSmall, 30, 52)) = VINE) And GameKeyDown(Climb) Then
        bClimbing = Not bHandsFull
        sRun = Sgn(sRun)
    End If

    If Not GameKeyDown(Jump) And sJump = 0 Then bHoldJump = False

    If Not bClimbing Then
        MovePlayerY
        MovePlayerY
        HitTestY
        MovePlayerX
        HitTestX
    Else
        HandleClimbing
    End If
    
    If lWingFrameTime > 0 Then lWingFrameTime = lWingFrameTime - 1
            
    TestForCoins
    HandlePowerups
    DetermineFrame
    sStarAngle = sStarAngle + 5
    If sStarAngle >= 360 Then sStarAngle = sStarAngle - 360
    bStarFrameOn = Not bStarFrameOn
    
    If lPowTime Then
        If Not SoundPlaying(Sounds.PowLoop) Then
            lPowTime = lPowTime - 1
            PlaySound Sounds.PowLoop
        End If
        If lPowTime = 0 Then
            LevelBlockSwitch
            EndSound Sounds.PowLoop
        End If
    End If
    
    If GameKeyDown(Jump) Then bHoldJump = True
    bForceStand = False
    
    If Not bAlive Then xSrc = 256
    
End Sub




Async Private Sub HandlePiping()
Dim tTypeA As udeLTileType
Dim tTypeB As udeLTileType
Dim tTag As udeLTileTag
    
    If lPipeStat = mpNothing Then
    
        tTypeA = GetTileAtPoint(CLng(xPos) - 18, CLng(yPos) + 2)
        tTypeB = GetTileAtPoint(CLng(xPos) - 14, CLng(yPos) + 2)
        tTag = GetTileTagAtPoint(CLng(xPos) - 16, CLng(yPos) + 2)
        If GameKeyDown(Down) And sGravity = 0 And tTypeA = SOLID And tTypeB = SOLID And tTag >= 8 And tTag <= 23 And lPipeStat = mpNothing Then
            If oCurWorldData.LevelData(curLevel).PipeDest(tTag - 7).destLevel >= 0 Then
                lPipeStat = mpEnterDown
                lPipeValue = 0
                lPipeValueMax = IIf(mStatus = MarioSmall, 32, 64)
                lPipeUsed = tTag - 7
                PlaySound Sounds.Pipe
            End If
        End If
        
        tTypeA = GetTileAtPoint(CLng(xPos) - 18, CLng(yPos) - IIf(mStatus = MarioSmall, 38, 64))
        tTypeB = GetTileAtPoint(CLng(xPos) - 14, CLng(yPos) - IIf(mStatus = MarioSmall, 38, 64))
        tTag = GetTileTagAtPoint(CLng(xPos) - 16, CLng(yPos) - IIf(mStatus = MarioSmall, 38, 64))
        If GameKeyDown(Up) And tTypeA = SOLID And tTypeB = SOLID And tTag >= 8 And tTag <= 23 And lPipeStat = mpNothing Then
            If oCurWorldData.LevelData(curLevel).PipeDest(tTag - 7).destLevel >= 0 Then
                lPipeStat = mpEnterUp
                lPipeValue = 0
                lPipeValueMax = IIf(mStatus = MarioSmall, 32, 64)
                lPipeUsed = tTag - 7
                PlaySound Sounds.Pipe
                yPos = IIf(mStatus = MarioSmall, 64, 90) + (__intDiv((CLng(yPos) - IIf(mStatus = MarioSmall, 38, 64)) , 32) * 32)
            End If
        End If
        
        tTypeA = GetTileAtPoint(CLng(xPos) + 18, CLng(yPos) - 2)
        tTypeB = GetTileAtPoint(CLng(xPos) + 18, CLng(yPos) - 30)
        tTag = GetTileTagAtPoint(CLng(xPos) + 18, CLng(yPos) - 16)
        If GameKeyDown(Right) And tTypeA = SOLID And sGravity = 0 And tTypeB = SOLID And tTag >= 8 And tTag <= 23 And lPipeStat = mpNothing Then
            If oCurWorldData.LevelData(curLevel).PipeDest(tTag - 7).destLevel >= 0 Then
                lPipeStat = mpEnterRight
                lPipeValue = 0
                lPipeValueMax = 32
                lPipeUsed = tTag - 7
                PlaySound Sounds.Pipe
            End If
        End If
    
        tTypeA = GetTileAtPoint(CLng(xPos) - 18, CLng(yPos) - 2)
        tTypeB = GetTileAtPoint(CLng(xPos) - 18, CLng(yPos) - 30)
        tTag = GetTileTagAtPoint(CLng(xPos) - 18, CLng(yPos) - 16)
        If GameKeyDown(Left) And tTypeA = SOLID And sGravity = 0 And tTypeB = SOLID And tTag >= 8 And tTag <= 23 And lPipeStat = mpNothing Then
            If oCurWorldData.LevelData(curLevel).PipeDest(tTag - 7).destLevel >= 0 Then
                lPipeStat = mpEnterLeft
                lPipeValue = 0
                lPipeValueMax = 32
                lPipeUsed = tTag - 7
                PlaySound Sounds.Pipe
            End If
        End If
        
        tTag = GetTileTagAtPoint(CLng(xPos), CLng(yPos) - 16)
        If GameKeyDown(Up) And sGravity = 0 And tTag >= 8 And tTag <= 23 And lPipeStat = mpNothing Then
            If TileExistsAtPoint(CLng(xPos), CLng(yPos) - 16) Then
            If oCurWorldData.LevelData(curLevel).PipeDest(tTag - 7).destLevel >= 0 Then
                lPipeStat = mpEnterDoor
                lPipeValue = 0
                lPipeValueMax = 16
                lPipeUsed = tTag - 7
                SetTile CLng(xPos), CLng(yPos) - 16, 1, 0, True
                SetTile CLng(xPos), CLng(yPos) - 48, 1, 0, True
                PlaySound Sounds.Door
            End If
            End If
        End If
        If lPipeStat = mpNothing Then Exit Sub
        
    End If
        
    lPipeValue = lPipeValue + 2
    If lPipeValue >= lPipeValueMax Then
        If lPipeStat < 6 Then
            Do While SoundPlaying(Sounds.Pipe)
                DoEvents
            Loop
            bAlive = False
        Else
            If lPipeStat = mpExitDoor Then
                SetTile CLng(xPos), CLng(yPos) - 16, -1, 0, True
                SetTile CLng(xPos), CLng(yPos) - 48, -1, 0, True
            End If
            lPipeStat = mpNothing
        End If
    End If
    
    Select Case mStatus
        Case MarioSmall
            xSrc = 0
        Case MarioBig
            xSrc = 32
        Case MarioMoonboot
            xSrc = 64
        Case MarioFlower
            xSrc = 96
        Case MarioHammer
            xSrc = 128
    End Select
    
    If lPipeStat = mpEnterDown Or lPipeStat = mpEnterUp Then
        ySrc = IIf(mStatus = MarioSmall, 32, 0)
        SrcHeight = IIf(mStatus = MarioSmall, 32, 64) - lPipeValue
        SrcWidth = 32
    ElseIf lPipeStat = mpExitDown Or lPipeStat = mpExitUp Then
        ySrc = IIf(mStatus = MarioSmall, 32, 0)
        SrcHeight = lPipeValue
        SrcWidth = 32
    ElseIf lPipeStat = mpEnterRight Or lPipeStat = mpEnterLeft Then
        ySrc = IIf(mStatus = MarioSmall, 96, 64)
        If lPipeStat = mpEnterLeft Then ySrc = ySrc + 64
        SrcHeight = IIf(mStatus = MarioSmall, 32, 64)
        SrcWidth = lPipeValue
    ElseIf lPipeStat = mpExitRight Or lPipeStat = mpExitLeft Then
        ySrc = IIf(mStatus = MarioSmall, 96, 64)
        If lPipeStat = mpExitLeft Then ySrc = ySrc + 64
        SrcHeight = IIf(mStatus = MarioSmall, 32, 64)
        SrcWidth = 32 - lPipeValue
    ElseIf lPipeStat = mpExitDoor Or lPipeStat = mpEnterDoor Then
        If lPipeStat = mpEnterDoor Then ySrc = IIf(mStatus = MarioSmall, 224, 192) Else ySrc = IIf(mStatus = MarioSmall, 32, 0)
        SrcHeight = IIf(mStatus = MarioSmall, 32, 64)
        SrcWidth = 32
    End If
    
errOut:
End Sub






Private Sub HandleClimbing()
Dim sngTop As Single
Dim bTTest As Boolean
Dim oldCount As Long
   
    If mStatus = MarioSmall Then sngTop = 30 Else sngTop = 52

    If GameKeyDown(Jump) Then
        sRun = 0
        sJump = 5
        sGravity = 1
        bClimbing = False
        PlaySound Sounds.Jump
        Exit Sub
    End If
    
    If (GetTileAtPoint(CLng(xPos), CLng(yPos) - 2) <> VINE) And (GetTileAtPoint(CLng(xPos), CLng(yPos) - sngTop) <> VINE) Then
        sRun = 0
        sJump = 0
        sGravity = 1
        bClimbing = False
        Exit Sub
    End If
    
    oldCount = climbFrameCount
    
    If GameKeyDown(Climb) And (Not isTileSolid(CLng(xPos) - 14, CLng(yPos) - sngTop - 2)) And (Not isTileSolid(CLng(xPos) + 14, CLng(yPos) - sngTop - 2)) Then
        yPos = yPos - 2
        If climbFrameCount = oldCount Then climbFrameCount = climbFrameCount + 1
    End If
        
    If GameKeyDown(Down) And (Not isTileSolid(CLng(xPos) - 14, CLng(yPos) + 2, True)) And (Not isTileSolid(CLng(xPos) + 14, CLng(yPos) + 2, True)) Then
        yPos = yPos + 2
        If climbFrameCount = oldCount Then climbFrameCount = climbFrameCount + 1
    End If

    If GameKeyDown(Left) Then
        sRun = -1
        bTTest = (Not isTileSolid(CLng(xPos) - 16, CLng(yPos))) And (Not isTileSolid(CLng(xPos) - 16, CLng(yPos) - 30))
        If mStatus <> MarioSmall Then bTTest = bTTest And (Not isTileSolid(CLng(xPos) - 16, CLng(yPos) - 52))
        If bTTest Then
            xPos = xPos - 2
            If climbFrameCount = oldCount Then climbFrameCount = climbFrameCount + 1
        End If
    End If
        
    If GameKeyDown(Right) Then
        sRun = 1
        bTTest = (Not isTileSolid(CLng(xPos) + 16, CLng(yPos))) And (Not isTileSolid(CLng(xPos) + 16, CLng(yPos) - 30))
        If mStatus <> MarioSmall Then bTTest = bTTest And (Not isTileSolid(CLng(xPos) + 16, CLng(yPos) - 52))
        If bTTest Then
            xPos = xPos + 2
            If climbFrameCount = oldCount Then climbFrameCount = climbFrameCount + 1
        End If
    End If
    
    If climbFrameCount >= 5 Then
        climbFrameCount = 0
        If curClimbFrame = 0 Then curClimbFrame = 32 Else curClimbFrame = 0
    End If
    
End Sub




Private Sub HandlePowerups()

    If lStarTime > 0 Then
        lStarTime = lStarTime - 1
        If lStarTime = 0 Then
            If bStarOffence Then
                bStarOffence = False
                MusicStopMusic
                MusicLoadFile sCurMusic
                MusicPlayMusic
                addShield 2, False
            End If
            bStarFrameOn = False
        End If
    End If
    
    lEarFlapTime = lEarFlapTime + 1
    If lEarFlapTime > 4 Then
        bEarFlapFrame = Not bEarFlapFrame
        lEarFlapTime = 0
    End If
    
    If lBlueStarTime > 0 Then
        lBlueStarTime = lBlueStarTime - 1
    End If

    If lLastFire > 0 Then lLastFire = lLastFire - 1
    bFloating = False
    If GameKeyDown(Shoot) Then
        Select Case mStatus
            Case MarioMoonboot
                If sJump = 0 Then bFloating = True
                If Not SoundPlaying(Sounds.Floating) And (sGravity - sJump) > 0 Then PlaySound Sounds.Floating
            Case MarioFlower
                If lLastFire = 0 Then
                    If MarioThrowFire(xPos + (-32 * bMarioFaceRight) - 16, yPos - 32, bMarioFaceRight) Then
                        lLastFire = FIRING_FREQ
                        PlaySound Sounds.Fireball
                    End If
                End If
            Case MarioHammer
                If lLastFire = 0 Then
                    If MarioThrowHammer(xPos + (-32 * bMarioFaceRight) - 16, yPos - 32, sRun) Then
                        lLastFire = FIRING_FREQ
                        PlaySound Sounds.Fireball
                    End If
                End If
        End Select
    End If
    
End Sub




Private Sub DetermineFrame()
Dim maxFrame As Long

    If lPipeStat Then Exit Sub
    
    If bForceStand Then sGravity = 0

    frameCount = frameCount + 1
    If GameKeyDown(Run) Then maxFrame = 5 Else maxFrame = 10
    If frameCount >= maxFrame Then
        frameCount = 0
        curRunFrame = IIf(curRunFrame = 1, 0, 1)
    End If

    If ySrc <> 0 And ySrc <> 64 Then ySrc = 0
    If GameKeyDown(Right) Then bMarioFaceRight = True
    If GameKeyDown(Left) Then bMarioFaceRight = False
    If bMarioFaceRight Then ySrc = 0 Else ySrc = 64

    bTall = True
    If Not bAlive Then
        xSrc = 256
    ElseIf sJump Then
        xSrc = 96
    ElseIf sGravity Then
        xSrc = 64
    ElseIf GameKeyDown(Down) Then
        xSrc = 160
        bTall = False
    ElseIf sRun > 0 Then
        If GameKeyDown(Left) Then
            xSrc = 128
            ySrc = 64
            bMarioFaceRight = False
        Else
            xSrc = 32 + (32 * curRunFrame)
            ySrc = 0
            bMarioFaceRight = True
        End If
    ElseIf sRun < 0 Then
        If GameKeyDown(Right) Then
            xSrc = 128
            ySrc = 0
            bMarioFaceRight = True
        Else
            xSrc = 32 + (32 * curRunFrame)
            ySrc = 64
            bMarioFaceRight = False
        End If
    Else
        xSrc = 0
    End If
    
    If bClimbing Then ySrc = 0
    
    SrcHeight = 64
    If mStatus = MarioSmall Then
        bTall = False
        ySrc = ySrc / 2
        SrcHeight = 32
        ySrc = ySrc + 128
    ElseIf mStatus = MarioMoonboot Then
        ySrc = ySrc + 192
    ElseIf mStatus = MarioFlower Then
        ySrc = ySrc + 320
    ElseIf mStatus = MarioHammer Then
        ySrc = ySrc + 448
    End If
    
    If bClimbing Then
        xSrc = 288
        If curClimbFrame Then ySrc = ySrc + SrcHeight
    End If
    If bSwimming And (sGravity <> 0) Then xSrc = xSrc + 128
    
    If (bFloating And xSrc = 64 And bEarFlapFrame) Or (lLastFire > (FIRING_FREQ - 8)) Then xSrc = IIf(bHandsFull, 256, 320)
    
End Sub



Private Sub MovePlayerY()
Dim sJumpVal As Single

    'make the player jump/swim
    If GameKeyDown(Jump) And Not bHoldJump And sJump = 0 Then
        If bSwimming Then
            If bAlwaysSwimming And yPos > (lAlwaysSwimmingTop + 32) Then
                sJump = 5
            Else
                If (GetTileAtPoint(CLng(xPos), CLng(yPos) - 32) = WATER) Then sJump = 5 Else sJump = 8
            End If
            PlaySound Sounds.Swim
        ElseIf lBlueStarTime Then
            sJump = 6
            lWingFrameTime = 10
            PlaySound Sounds.Fly
        ElseIf sGravity = 0 Then
            If mStatus = MarioMoonboot Or sRun > 6.5 Or sRun < -6.5 Then sJump = 12 Else sJump = 10
            PlaySound Sounds.Jump
        End If
    End If
    
    'move the player
   ' If lEquipFireStat <> 2 Then
        If GameKeyDown(Jump) Then sJumpVal = 0.956 Else sJumpVal = 0.89
        sJump = sJump * sJumpVal
        If sJump < 1 Then sJump = 0
        If sGravity < 0.5 Then sGravity = 0.5
        sGravity = (sGravity * 1.08) - sJump
        If bSwimming Then
            If sGravity > 1.2 Then sGravity = 1.2
        ElseIf bFloating Then
            If sGravity > 0.8 Then sGravity = 0.8
        Else
            If sGravity > 5 Then sGravity = 5
        End If
    'End If
    yPos = yPos + sGravity
          
End Sub


Private Sub MovePlayerX()
Dim sDimRun As Single
Dim sMaxRun As Single
Dim sAcceleration As Single  'movement while moving in direction pressed
Dim sFriction As Single      'movement while turning around key down
Dim tType As udeLTileType


    'set the maximum speed
    If GameKeyDown(Run) And Not bSwimming Then sMaxRun = 6.7 Else sMaxRun = 3.5
    
    'make ice slippery
    If GetTileAtPoint(CLng(xPos), CLng(yPos) + 1) = ICE Then
        sAcceleration = 0.1
        sFriction = 0.1
    ElseIf GetTileAtPoint(CLng(xPos), CLng(yPos) + 1) = Background Then
        sAcceleration = 0.25
        sFriction = 0.3
    Else
        sAcceleration = 0.38
        sFriction = 0.4
    End If
    
    'apply speed
    If GameKeyDown(Right) And Not GameKeyDown(Down) And sRun < sMaxRun Then
        If sRun > 0 Then
            sRun = sRun + sAcceleration
        Else
            sRun = sRun + sFriction
        End If
    ElseIf GameKeyDown(Left) And Not GameKeyDown(Down) And sRun > -sMaxRun Then
        If sRun < 0 Then
            sRun = sRun - sAcceleration
        Else
            sRun = sRun - sFriction
        End If
    Else
        If GameKeyDown(Down) Then sDimRun = 0.2 Else sDimRun = 0.15
        If GetTileAtPoint(CLng(xPos), CLng(yPos) + 1) = ICE Then sDimRun = sDimRun / 2
        If sRun > 0.8 Then
            sRun = sRun - sDimRun
        ElseIf sRun < -0.8 Then
            sRun = sRun + sDimRun
        Else
            sRun = 0
        End If
    End If
    
    'move player
    xPos = xPos + sRun

End Sub


Private Sub HitTestY()
Dim tType As udeLTileType
Dim tTypeB As udeLTileType
Dim tTypeC As udeLTileType
Dim tSolidA As Boolean
Dim tSolidB As Boolean
Dim tSolidC As Boolean
Dim sOldGravity As Single
Dim lHeadTop As Long

    'determine where the top of mario is
    If bTall Then lHeadTop = 48 Else lHeadTop = 30

    'test the tiles INSIDE mario
    If isTileSolid(CLng(xPos), CLng(yPos - (lHeadTop / 2))) And sGravity < 0 Then yPos = yPos + (2 * Abs(sGravity))

    'test the tiles above mario
    If sGravity < 0 Then
        If isTileSolid(CLng(xPos), CLng(yPos) - lHeadTop) Or hitHiddenBlock(CLng(xPos), CLng(yPos) - lHeadTop) Then
        
            PlaySound Sounds.Bump
            sOldGravity = sGravity
            sJump = 0
            
            HitSpecialBlock CLng(xPos), CLng(yPos) - lHeadTop, xPos Mod 32 < 16, Mario.mStatus = MarioSmall
    
            'handle blocks
            tType = GetTileAtPoint(CLng(xPos), CLng(yPos) - lHeadTop)
            If tType = BOUNCY Then
                sGravity = 4
                PlaySound Sounds.Boing
            ElseIf tType = INJURETILE Then
                Me.Injure
            ElseIf tType = DEADLY Then
                Me.makeDie
            End If
            
            yPos = yPos + 10
            
        End If
    End If
    
    'test the tiles below mario
    If sJump = 0 Then
        tSolidA = isTileSolid(CLng(xPos) - 12, CLng(yPos), True)
        tSolidB = isTileSolid(CLng(xPos) + 12, CLng(yPos), True)
        tSolidC = isTileSolid(CLng(xPos), CLng(yPos), True)
        If tSolidA Or tSolidB Then
        
            sGravity = 0
            yPos = __intDiv(yPos , 32) * 32
            
            tType = GetTileAtPoint(CLng(xPos) - 12, CLng(yPos))
            tTypeB = GetTileAtPoint(CLng(xPos) + 12, CLng(yPos))
            tTypeC = GetTileAtPoint(CLng(xPos), CLng(yPos))
            If (tSolidA And tType = BOUNCY) Or (tSolidB And tTypeB = BOUNCY) Then
                If (Not tSolidC) Or (tTypeC = BOUNCY) Then
                    sJump = 16
                    PlaySound Sounds.Boing
                End If
            ElseIf (tSolidA And tType = INJURETILE) Or (tSolidB And tTypeB = INJURETILE) Then
                If (Not tSolidC) Or (tTypeC = INJURETILE) Then Me.Injure
            ElseIf (tSolidA And tType = DEADLY) Or (tSolidB And tTypeB = DEADLY) Then
                If (Not tSolidC) Or (tTypeC = DEADLY) Then Me.makeDie
            End If
            
        End If
    End If

End Sub

Private Sub HitTestX()
Dim tTypeA As udeLTileType
Dim tTypeB As udeLTileType
Dim bColl As Boolean
Dim yTest As Long

    'keep the player on the screen or warp him accross
    If GetLevelSideWarp Then
        If xPos <= -9 Then xPos = GetLevelWidth + 8
        If xPos >= GetLevelWidth + 9 Then xPos = -8
    Else
        If xPos <= 16 Then 
            xPos = 16
            sRun = 0
        End If
        If xPos >= GetLevelWidth - 16 Then 
            xPos = GetLevelWidth - 16
            sRun = 0
        End If
    End If
    
    'horizontal bottom collision detection
    If yPos > 8 Then yTest = CLng(yPos) - 8 Else yTest = 8
    If isTileSolid(CLng(xPos) + 15, yTest) Or isTileSolid(CLng(xPos) - 15, yTest) Then
        
        tTypeA = GetTileAtPoint(CLng(xPos) + 15, yTest)
        tTypeB = GetTileAtPoint(CLng(xPos) - 15, yTest)
        If tTypeA = BOUNCY Or tTypeB = BOUNCY Then
            If sRun > 0 Then sRun = -5 Else sRun = 5
            PlaySound Sounds.Boing
        ElseIf tTypeA = INJURETILE Or tTypeB = INJURETILE Then
            Me.Injure
            sRun = 0
        ElseIf tTypeA = DEADLY Or tTypeB = DEADLY Then
            Me.makeDie
            sRun = 0
        Else
            sRun = 0
        End If

        xPos = (__intDiv(CLng(xPos) , 32) * 32) + 16
        
    End If
    
    'test if marios face is hitting a block
    If bTall Then yTest = 40 Else yTest = 24
    If isTileSolid(CLng(xPos) + 15, CLng(yPos) - yTest) Or isTileSolid(CLng(xPos) - 15, CLng(yPos) - yTest) Then
        
        tTypeA = GetTileAtPoint(CLng(xPos) + 15, CLng(yPos) - yTest)
        tTypeB = GetTileAtPoint(CLng(xPos) - 15, CLng(yPos) - yTest)
        If tTypeA = BOUNCY Or tTypeB = BOUNCY Then
            If sRun > 0 Then sRun = -5 Else sRun = 5
            PlaySound Sounds.Boing
        ElseIf tTypeA = INJURETILE Or tTypeB = INJURETILE Then
            Me.Injure
            sRun = 0
        ElseIf tTypeA = DEADLY Or tTypeB = DEADLY Then
            Me.makeDie
            sRun = 0
        Else
            sRun = 0
        End If
        
        If bTall Then
            If isTileSolid(CLng(xPos), CLng(yPos) - 40) Then xPos = xPos - 2 Else xPos = (__intDiv(CLng(xPos) , 32) * 32) + 16
        Else
            xPos = (__intDiv(CLng(xPos) , 32) * 32) + 16
        End If
        
    End If
    
End Sub


Private Sub TestForCoins()
Dim tType(5) As udeLTileType

    'get all tiles mario can possibly touch
    tType(0) = GetTileAtPoint(CLng(xPos) - 14, CLng(yPos) - 2)
    tType(1) = GetTileAtPoint(CLng(xPos) + 14, CLng(yPos) - 2)
    tType(2) = GetTileAtPoint(CLng(xPos) - 14, CLng(yPos) - 28)
    tType(3) = GetTileAtPoint(CLng(xPos) + 14, CLng(yPos) - 28)
    tType(4) = GetTileAtPoint(CLng(xPos) - 14, CLng(yPos) - 50)
    tType(5) = GetTileAtPoint(CLng(xPos) + 14, CLng(yPos) - 50)

    
    'test for death tiles
Dim bool As Boolean
Dim i As Long
    For i = 0 To 5
    bool = bool Or tType(i) = DEADLY
    Next i
    If bool Then
        bAlive = False
        xSrc = 256
    End If
    
    'test for coins
    If tType(0) = Coin Then
        KillTile CLng(xPos) - 14, CLng(yPos) - 2
        gameCoins = gameCoins + 1
        PlaySound Sounds.Coin
    ElseIf tType(1) = Coin Then
        KillTile CLng(xPos) + 14, CLng(yPos) - 2
        gameCoins = gameCoins + 1
        PlaySound Sounds.Coin
    ElseIf tType(2) = Coin Then
        KillTile CLng(xPos) - 14, CLng(yPos) - 28
        gameCoins = gameCoins + 1
        PlaySound Sounds.Coin
    ElseIf tType(3) = Coin Then
        KillTile CLng(xPos) + 14, CLng(yPos) - 28
        gameCoins = gameCoins + 1
        PlaySound Sounds.Coin
    ElseIf tType(4) = Coin And bTall Then
        KillTile CLng(xPos) - 14, CLng(yPos) - 50
        gameCoins = gameCoins + 1
        PlaySound Sounds.Coin
    ElseIf tType(5) = Coin And bTall Then
        KillTile CLng(xPos) + 14, CLng(yPos) - 50
        gameCoins = gameCoins + 1
        PlaySound Sounds.Coin
    End If
    
End Sub




Public Sub DrawOnScreen()
    xOnScreen = xPos - screenLeft
    yOnScreen = yPos - screenTop
    If lPipeStat Then
        If lPipeStat = mpEnterDown Or lPipeStat = mpExitUp Then
            GFX.DrawSurface surfList.MarioPipe, CSng(xSrc), CSng(ySrc), SrcWidth, SrcHeight, CSng(xOnScreen) - 16, CSng(yOnScreen) - SrcHeight
        ElseIf lPipeStat = mpEnterUp Or lPipeStat = mpExitDown Then
            GFX.DrawSurface surfList.MarioPipe, CSng(xSrc), CSng(ySrc) + IIf(mStatus = MarioSmall, 32, 64) - SrcHeight, SrcWidth, SrcHeight, CSng(xOnScreen) - 16, CSng(yOnScreen) - IIf(mStatus = MarioSmall, 32, 58)
        ElseIf lPipeStat = mpEnterRight Or lPipeStat = mpExitLeft Then
            GFX.DrawSurface surfList.MarioPipe, CSng(xSrc), CSng(ySrc), 32 - SrcWidth, SrcHeight, CSng(xOnScreen) - 16 + SrcWidth, CSng(yOnScreen) - SrcHeight
        ElseIf lPipeStat = mpEnterLeft Or lPipeStat = mpExitRight Then
            GFX.DrawSurface surfList.MarioPipe, CSng(xSrc) + SrcWidth, CSng(ySrc), 32 - SrcWidth, SrcHeight, CSng(xOnScreen) - 16, CSng(yOnScreen) - SrcHeight
        ElseIf lPipeStat = mpEnterDoor Or lPipeStat = mpExitDoor Then
            GFX.DrawSurface surfList.MarioPipe, CSng(xSrc), CSng(ySrc), SrcWidth, SrcHeight, CSng(xOnScreen) - 16, CSng(yOnScreen) - SrcHeight
        End If
    Else
        If Not (bStarFrameOn And lStarTime > 0) Then
            If (lBlueStarTime > 0) And Not (lBlueStarTime < 100 And bStarFrameOn) Then
                If bMarioFaceRight Then
                    GFX.DrawSurface surfList.MarioWings, IIf(lWingFrameTime > 0, 32, 0), 32, 32, 32, CSng(xOnScreen), CSng(yOnScreen) - SrcHeight - IIf(bTall, 8, 20)
                Else
                    GFX.DrawSurface surfList.MarioWings, IIf(lWingFrameTime > 0, 32, 0), 0, 32, 32, CSng(xOnScreen) - 32, CSng(yOnScreen) - SrcHeight - IIf(bTall, 8, 20)
                End If
            End If
            GFX.DrawSurface IIf(bHandsFull, surfList.MarioCarry, surfList.Mario), CSng(xSrc), CSng(ySrc), 32, SrcHeight, CSng(xOnScreen) - 16, CSng(yOnScreen) - SrcHeight
            If (lBlueStarTime > 0) And Not (lBlueStarTime < 100 And bStarFrameOn) Then
                If bMarioFaceRight Then
                    GFX.DrawSurface surfList.MarioWings, IIf(lWingFrameTime > 0, 32, 0), 0, 32, 32, CSng(xOnScreen) - 32, CSng(yOnScreen) - SrcHeight - IIf(bTall, 8, 20)
                Else
                    GFX.DrawSurface surfList.MarioWings, IIf(lWingFrameTime > 0, 32, 0), 32, 32, 32, CSng(xOnScreen), CSng(yOnScreen) - SrcHeight - IIf(bTall, 8, 20)
                End If
            End If
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If lEquipStat = 2 Then
'            GFX.DrawSurface surfList.Objects, 0, 192, 16, 16, CSng(xOnScreen) + (96 * Cos((sEquipAngle - 90) * PI / 180) * IIf(bMarioFaceRight, 1, -1)) - 16, CSng(yOnScreen) + (96 * Sin((sEquipAngle - 90) * PI / 180)) - 32
'        End If
'        If lEquipFireStat > 0 Then
'            GFX.DrawSurface surfList.Objects, 0, 192, 16, 16, sEquipX - screenLeft, sEquipY - screenTop
'        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If ((lStarTime > 0) And bStarOffence) Then GFX.DrawSurface surfList.Objects, 22, 64, 48, 45, CSng(xOnScreen) - 24, CSng(yOnScreen) - (SrcHeight / 2) - 22, , , sStarAngle, 50
    End If
End Sub







Private Sub Class_Initialize()
bAlive = True
bTall = MarioSmall
bHasWon = False
bMarioFaceRight = True
bFloating = False
End Sub
