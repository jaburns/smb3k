Attribute VB_Name = "mMainGame"
Option Explicit

Public Enum udeLevelReturnValue
    lvrQuit
    lvrPass
    lvrPipe
    lvrDie
    lvrPipeOut
    lvrAbort
End Enum

Public Enum udeScrollStyle
    lScr_Free
    lScr_Right
    lScr_Left
    lScr_Up
    lScr_Down
End Enum

Private Const BAR_FORECOLOR As Long = vbGreen
Private Const BAR_BACKCOLOR As Long = vbBlue

'this object has all the info on mario
Public Mario As New ocMario

Public lTime As Long
Public lMaxTime As Long

Public levelReturnValue As udeLevelReturnValue
Public levelReturnPipe As Long

Public marioReserveItem As udeMarioReserveItem
Public powerStarX As Single
Public powerStarY As Single
Public powerStarZ As Single

Public levScrollStyle As udeScrollStyle
Public levScrollSpeed As Single

Private drawLevelX As Long
Private drawLevelY As Long

Private lTickTimeCount As Long
Private bAlreadyStartedLevel As Boolean
Public bTimingTime As Boolean
Public bDisableTime As Boolean

Private timeFireAngle As Single
Private layer3y As Single           'height of rising/falling lava/water
Private layer3speed As Single       'speed of   '
Private layer3frm As Long
Private layer3frmcount As Long
Public layer3enabled As Boolean
Public layer3sinusoid As Boolean
Public layer3maximum As Long


Async Public Sub MainGameLoop(sMusic As String, lX As Long, lY As Long, lInTime As Long, startPipe As udeMarioPipeStat)

    Set Mario = New ocMario
    Mario.bAlive = True
    Mario.xPos = CSng(lX)
    Mario.yPos = CSng(lY)
    Mario.mStatus = gameMarioStat
    powerStarX = 0
    powerStarY = 0
    powerStarZ = 0
    layer3enabled = False
    layer3sinusoid = False
    
    levScrollStyle = oCurWorldData.LevelData(curLevel).iScrollStyle
    levScrollSpeed = oCurWorldData.LevelData(curLevel).iScrollSpeed
    
    levelReturnValue = lvrQuit
    
    GFX.SetFont "Courier New", 8, True, False, False, False

    InitObjects
    FindAndCreateEnemies
           
    bDisableTime = (oCurWorldData.LevelData(curLevel).TimeGiven = 0)
    If Not bDisableTime Then
        If lInTime > 0 Then SetLevelTime lTime Else SetLevelTime CLng(oCurWorldData.LevelData(curLevel).TimeGiven), True
    End If
    
    If startPipe <> mpNothing Then Mario.makeExitPipe startPipe

    lTime = DecrementTime
    
    CenterLevelMarioX
    CenterLevelMarioY
    HandleScrolling
    showFadeIn False
    
    If fileExist(App.Path & "\Data\Music\" & sMusic & ".mid") Then
        sCurMusic = App.Path & "\Data\Music\" & sMusic & ".mid"
    Else
        sCurMusic = App.Path & "\Mods\" & sWorldSetName & "\Music\" & sMusic & ".mid"
    End If
    If fileExist(sCurMusic) Then
        MusicLoadFile sCurMusic
        MusicPlayMusic
    End If
    
    If Not bAlreadyStartedLevel Then lTickTimeCount = 0
    
    layer3frm = 0
    layer3frmcount = 0
    layer3speed = 0.1
    layer3y = GetLevelHeight
    Do Until Not Mario.bAlive Or Mario.bHasWon
    
        UpdateInput
        If GameKeyDown(Pause) Or GameKeyDown(Quit) Then
            PauseGame
            If GameKeyDown(QuitFromPauseKey) Then Exit Do
        End If
        If GameKeyDown(LeaveLevel) And oWorldPassData.bWorldPassed(curWorld).bLevelPassed(curLevel) Then
            levelReturnValue = lvrAbort
            Exit Do
        End If
        
        If GameKeyDown(DebugA) And DebugModeEnabled Then
            Mario.mStatus = Mario.mStatus + 1
            If Mario.mStatus > 4 Then Mario.mStatus = 0
        End If
        
        If GameKeyDown(DebugB) And DebugModeEnabled Then
            levelReturnValue = lvrAbort
            bAlreadyStartedLevel = False
            gameMarioStat = Mario.mStatus
            KillEnemies
            KillObjects
            KillLevel
            Exit Sub
        End If
        
        GFX.BeginScene 25
            
            If Not bDisableTime Then
                lTickTimeCount = lTickTimeCount + 1
                lTime = DecrementTime
                If lTime <= 0 Then
                    Mario.bAlive = False
                    Mario.xSrc = 256      'deATH frame
                    GFX.EndScene
                    Exit Do
                End If
            End If
                        
            timeFireAngle = timeFireAngle + 5
            If timeFireAngle > 360 Then timeFireAngle = timeFireAngle - 360
            
            If gameCoins >= 100 Then
                gameCoins = 0
                gameLives = gameLives + 1
                PlaySound Sounds.OneUp
            End If
                        
            If GameKeyDown(Release) Then ReleaseReserveItem
            
            Mario.UpdatePosition
            GFX.DrawRect 0, 0, 640, 480, 0
            HandleScrolling
            DrawLevel drawLevelX, drawLevelY
            HandleEnemies
            HandleObjects
            Mario.DrawOnScreen
            'HandleSavePoint
            If layer3enabled Then HandleLayerThree
            DrawBar
            If bTimingTime Then GFX.DrawText CStr(CLng((lTickTimeCount / 30) * 100) / 100), 60, 60, 0
            'HandleAndDrawPowerStar
                        
        GFX.EndScene
    
    Loop

    MusicStopMusic
    
    If levelReturnValue = lvrAbort Then
        bAlreadyStartedLevel = False
        gameMarioStat = Mario.mStatus
        ShowMarioWin True
    Else
        gameMarioStat = Mario.mStatus
        If Not Mario.bAlive Then
            If Mario.LastPipeUsed = 0 Then
                ShowMarioDie
                levelReturnValue = lvrDie
                gameMarioStat = MarioSmall
                bAlreadyStartedLevel = False
            Else
                bAlreadyStartedLevel = True
                levelReturnPipe = Mario.LastPipeUsed
                levelReturnValue = lvrPipe
            End If
        ElseIf Mario.bHasWon Then
            bAlreadyStartedLevel = False
            ShowMarioWin
            levelReturnValue = lvrPass
        Else
            bAlreadyStartedLevel = False
            levelReturnValue = lvrQuit
        End If
    End If
    
    showFadeOut False
    KillEnemies
    KillObjects
    KillLevel

End Sub


Private Sub HandleScrolling()

    If levScrollStyle = lScr_Right Then
    
        If drawLevelX < GetLevelWidth - 640 Then drawLevelX = drawLevelX + levScrollSpeed
        If drawLevelX > GetLevelWidth - 640 Then drawLevelX = GetLevelWidth - 640
        CenterLevelMarioY
        If Mario.xPos < drawLevelX + 16 Then
            Mario.ResetXPos drawLevelX + 16, levScrollSpeed
            If isTileSolid(Mario.xPos + 18, Mario.yPos - 16) Then Mario.makeDie
        ElseIf Mario.xPos > drawLevelX + 624 Then
            Mario.ResetXPos drawLevelX + 624, 0
        End If
        
    ElseIf levScrollStyle = lScr_Left Then
    
        If drawLevelX > 0 Then drawLevelX = drawLevelX - levScrollSpeed
        If drawLevelX < 0 Then drawLevelX = 0
        CenterLevelMarioY
        If Mario.xPos > drawLevelX + 624 Then
            Mario.ResetXPos drawLevelX + 624, -levScrollSpeed
            If isTileSolid(Mario.xPos - 18, Mario.yPos - 16) Then Mario.makeDie
        ElseIf Mario.xPos < drawLevelX + 16 Then
            Mario.ResetXPos drawLevelX + 16, 0
        End If
        
    ElseIf levScrollStyle = lScr_Up Then
    
        If drawLevelY > 0 Then drawLevelY = drawLevelY - levScrollSpeed
        If drawLevelY < 0 Then drawLevelY = 0
        CenterLevelMarioX
        If Mario.yPos > drawLevelY + 544 Then Mario.makeDie
        If Mario.yPos < drawLevelY + 32 Then Mario.makeStand Mario.xPos, Mario.yPos
        
    ElseIf levScrollStyle = lScr_Down Then
    
        If drawLevelY < GetLevelHeight - 480 Then drawLevelY = drawLevelY + levScrollSpeed
        If drawLevelY > GetLevelHeight - 480 Then drawLevelY = GetLevelHeight - 480
        CenterLevelMarioX
        If Mario.yPos < drawLevelY + IIf(Mario.isTall, 52, 32) Then
            Mario.ResetYPos Mario.yPos + (4 * levScrollSpeed), 1
            If isTileSolid(Mario.xPos, Mario.yPos + 2, True) Then Mario.makeDie
        ElseIf Mario.yPos > drawLevelY + 544 Then
            Mario.makeDie
        End If
        
    Else
        CenterLevelMarioX
        CenterLevelMarioY
    End If
    
End Sub

Private Sub CenterLevelMarioX()
    If Mario.xPos <= 320 Then
        drawLevelX = 0
    ElseIf Mario.xPos >= GetLevelWidth - 320 Then
        drawLevelX = GetLevelWidth - 640
    Else
        drawLevelX = Mario.xPos - 320
    End If
End Sub

Private Sub CenterLevelMarioY()
    If Mario.yPos <= 240 Then
        drawLevelY = 0
    ElseIf Mario.yPos >= GetLevelHeight - 240 Then
        drawLevelY = GetLevelHeight - 480
    Else
        drawLevelY = Mario.yPos - 240
    End If
End Sub


'Private Sub HandleAndDrawPowerStar()
'    powerStarZ = powerStarZ + 1
'    If powerStarZ = 2 Then
'        powerStarZ = 0
'        powerStarX = powerStarX + 48
'        If powerStarX = 528 Then
'            powerStarX = 0
'            powerStarY = powerStarY + 48
'            If powerStarY = 96 Then powerStarY = 0
'        End If
'    End If
'    GFX.DrawSurface surfList.PowerStar, powerStarX, powerStarY, 48, 48, 50, 50
'End Sub

Async Private Sub ShowMarioDie()
Dim offsetY As Single
Dim ySpeed As Single

    ySpeed = -10
    
    MusicStopMusic
    PlaySound Sounds.Death
    Do
    GFX.BeginScene 25
    
        offsetY = offsetY + ySpeed
        ySpeed = ySpeed + 0.4
        
        DrawLevel drawLevelX, drawLevelY
        HandleObjects
        HandleEnemies
        If layer3enabled Then DrawLayerThree
        GFX.DrawSurface surfList.Mario, Mario.xSrc, Mario.ySrc, 32, Mario.SrcHeight, CSng(Mario.xOnScreen) - 16, (CSng(Mario.yOnScreen) - Mario.SrcHeight) + offsetY
        If lTime <= 0 And Not bDisableTime Then
            GFX.SetFont "Impact", 24, True, True, False, False
            GFX.DrawText "TIME UP!", 240, 180, &HFF
        End If
        DrawBar
        
    GFX.EndScene
    If ((CSng(Mario.yOnScreen) - 64) + offsetY > 480) And Not SoundPlaying(Sounds.Death) Then Exit Do
    Loop

End Sub



Async Private Sub ShowMarioWin(Optional ByVal bQuickExit As Boolean = False)
Dim lGrowth As Long
Dim bDone As Boolean
Dim sRiseSpeed As Single
Dim sOffsetY As Single
Dim iColor As Integer
Dim lSound As Long

    MusicStopMusic
    lSound = IIf(bQuickExit, Sounds.QuickExitNoise, Sounds.Win)
    PlaySound lSound
    
    Do While SoundPlaying(lSound)
    GFX.BeginScene 25
    
        If sEndWheelSpeed >= 0.1 Then sEndWheelSpeed = sEndWheelSpeed * 0.97 Else sEndWheelSpeed = 0
        If lGrowth < 32 Then 
            lGrowth = lGrowth + 3 
        Else 
            If lGrowth > 32 Then lGrowth = 32
        End If

        If sOffsetY < 480 Then
            sRiseSpeed = sRiseSpeed + 0.2
            sOffsetY = sOffsetY + sRiseSpeed
        End If

        'determine the color the wheel is on
        iColor = 1
        If sEndWheelAngle >= 315 Then
            iColor = 2
        ElseIf sEndWheelAngle >= 270 Then
            iColor = 1
        ElseIf sEndWheelAngle >= 225 Then
            iColor = 2
        ElseIf sEndWheelAngle >= 180 Then
            iColor = 1
        ElseIf sEndWheelAngle >= 135 Then
            iColor = 2
        ElseIf sEndWheelAngle >= 90 Then
            iColor = 1
        ElseIf sEndWheelAngle >= 45 Then
            iColor = 2
        End If
        
        DrawLevel drawLevelX, drawLevelY
        HandleObjects
        HandleEnemies
        If layer3enabled Then DrawLayerThree
        If Not bQuickExit Then
            GFX.DrawSurface surfList.Sprites, 96, 0, 32, lGrowth, CSng(lEndBlockX), CSng(lEndBlockY) - lGrowth
        End If
        If bTimingTime Then GFX.DrawText CStr(CLng((lTickTimeCount / 30) * 100) / 100), 60, 60, 0
        GFX.DrawSurface surfList.Mario, Mario.xSrc, Mario.ySrc, 32, Mario.SrcHeight, CSng(Mario.xOnScreen) - 16, (CSng(Mario.yOnScreen) - Mario.SrcHeight) - sOffsetY
        DrawBar
        
    GFX.EndScene
    Loop
    
    If bQuickExit Then Exit Sub
    
    If iColor = 2 Then gameGreens = gameGreens + 1 Else gameGreens = gameGreens - 1
    If gameGreens <= 0 Or gameGreens >= 6 Then
        gameGreens = 3
        gameLives = gameLives + 1
        PlaySound Sounds.OneUp
    End If

End Sub



Private Sub HandleLayerThree()
    If layer3sinusoid Then
        layer3speed = layer3speed + 0.7  'speed is actually a time increment here
        If layer3speed > 360 Then layer3speed = layer3speed - 360
        layer3y = GetLevelHeight - (Sin((layer3speed - 90) * PI / 180) * (layer3maximum / 2)) - (layer3maximum / 2) - 16
    Else
        If layer3speed < 2 Then layer3speed = layer3speed * 1.08
        layer3y = layer3y - layer3speed
    End If
    DrawLayerThree
    If Mario.yPos > (layer3y + 8) Then
        Mario.bAlive = False
        Mario.xSrc = 256
    End If
End Sub
Private Sub DrawLayerThree()
Dim lOffX As Long
Dim i As Long
    If layer3y > screenTop + 480 Then Exit Sub
    layer3frmcount = layer3frmcount + 1
    If layer3frmcount = 4 Then
        layer3frmcount = 0
        layer3frm = layer3frm + 1
        If layer3frm = 4 Then layer3frm = 0
    End If
    lOffX = (screenLeft Mod 32)
    For i = 0 To IIf(lOffX, 20, 19)
        GFX.DrawSurface surfList.EnemyList.RisingLava, layer3frm * 32, 0, 32, 32, (i * 32) - lOffX, layer3y - screenTop, 32, 32
    Next i
    GFX.DrawSurface surfList.EnemyList.RisingLava, 128, 0, 32, 32, 0, layer3y - screenTop + 32, 640, 480 - (layer3y - screenTop + 32)
End Sub




Private Sub DrawBar()
Dim sBonus As String
Dim timeWidth As Single
Dim i As Long
    
    If Not bDisableTime Then
        timeWidth = ((lTime - (lTimeIncCount / 30)) / lMaxTime) * 640
        For i = 0 To __intDiv(timeWidth , 64)
            If i = __intDiv(timeWidth , 64) Then
                GFX.DrawSurface surfList.BarIcons, 64, 32, timeWidth Mod 64, 16, i * 64, 0
            Else
                GFX.DrawSurface surfList.BarIcons, 64, 32, 64, 16, i * 64, 0
            End If
        Next i
        GFX.DrawSurface surfList.Objects, 0, 192, 16, 16, ((i - 1) * 64) + (timeWidth Mod 64) - 8, 0, , , timeFireAngle, 75
    End If
    DrawBonusMeter
    
End Sub


Public Sub ReserveMariosItem()
    Select Case Mario.mStatus
        Case udeMarioStatus.MarioBig
            marioReserveItem = rsvMushroom
        Case udeMarioStatus.MarioFlower
            marioReserveItem = rsvFlower
        Case udeMarioStatus.MarioHammer
            marioReserveItem = rsvHammer
        Case udeMarioStatus.MarioMoonboot
            marioReserveItem = rsvCarrot
    End Select
End Sub


Public Sub ReleaseReserveItem()
    Select Case marioReserveItem
        Case udeMarioReserveItem.rsvMushroom
            MakeMushroom screenLeft + 320, screenTop + 56, True, 0, 0, True
        Case udeMarioReserveItem.rsvFlower
            MakeFlower screenLeft + 320, screenTop + 56, True
        Case udeMarioReserveItem.rsvHammer
            MakeHammerPickup screenLeft + 320, screenTop + 56, True
        Case udeMarioReserveItem.rsvCarrot
            MakeMoonboot screenLeft + 320, screenTop + 56, True, True
    End Select
    marioReserveItem = rsvEmpty
End Sub



Async Private Sub showFadeOut(Optional ByVal insideScene As Boolean = True)
Dim i As Long
    If insideScene Then GFX.EndScene
    For i = 0 To 100 Step 4
    GFX.BeginScene 25
        DrawLevel drawLevelX, drawLevelY
        HandleObjects
        HandleEnemies
        If layer3enabled Then DrawLayerThree
        If levelReturnValue = lvrQuit Then Mario.DrawOnScreen
        DrawBar
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
    GFX.EndScene
    Next i
    If insideScene Then GFX.BeginScene 25
End Sub

Async Private Sub showFadeIn(Optional ByVal insideScene As Boolean = True)
Dim i As Long
    If insideScene Then GFX.EndScene
    Mario.UpdatePosition
    For i = 100 To 0 Step -4
    GFX.BeginScene 25
        DrawLevel drawLevelX, drawLevelY
        HandleObjects
        HandleEnemies
        Mario.DrawOnScreen
        DrawBar
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
    GFX.EndScene
    Next i
    If insideScene Then GFX.BeginScene 25
End Sub



Async Private Sub PauseGame()
Dim sAngle As Single
Dim sFactor As Single
Dim lEscapeCount As Long

    Do While GameKeyDown(Pause)
        DoEvents
    Loop
    Do While GameKeyDown(Quit)
        DoEvents
    Loop
    
    GFX.SetFont "Courier New", 8, True, False, False, False
    
    MusicStopMusic
    Do Until GameKeyDown(Pause) Or GameKeyDown(Quit) Or GameKeyDown(QuitFromPauseKey)
    GFX.BeginScene 25
    
        sAngle = sAngle + 2
        sFactor = (Cos(sAngle * 3.141592654 / 180) * 20)
        DrawLevel drawLevelX, drawLevelY, True
        GFX.DrawSurface surfList.Objects, 0, 0, 256, 63, 192, 208, , , sFactor, 75
        GFX.DrawText " Press [Q] to quit game ", 0, 0, vbWhite, 0
    
    GFX.EndScene
    Loop
    MusicPlayMusic
    
    Do While GameKeyDown(Pause) Or GameKeyDown(Quit)
        DoEvents
    Loop
    
End Sub

