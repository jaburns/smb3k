Attribute VB_Name = "main"
Option Explicit



Private curPlayLevel As Long
Private mpExitPipe As udeMarioPipeStat
Private nextStartX As Long
Private nextStartY As Long
Private nextStartPipe As Long
Private nextMapNode As Long
Private lSaveTime As Long
Private bGameOver As Boolean

Private sWorldSet() As String




'
' this is where the game starts and ends
'
Private Sub Form_Load()

    KillInput
    InitInput hWnd

    'init everything
    GFX.Initialize frmMain.hWnd, 640, 480, 16
    Set oMouse = New DXMouse
    oMouse.Initialize hWnd, 0, 0, 640, 480, 1
    InitSurfaces
    LoadSounds
    InitObjects
    Music.Volume = 80
'   bMusicEnabled = IIf(InStr(Command$, "-nomusic"), False, True)
    bSoundEnabled = True 'IIf(InStr(Command$, "-nosound"), False, True)
    
    'hide the mouse
    Mouse_Hide
    
    'start the game going
    MainLoop
    
    'show the mouse
    Mouse_Show
    
    'clean up everything and exit
    KillSounds
    KillInput
    Set Music = Nothing
    Set Mario = Nothing
    Set GFX = Nothing
    Set oMouse = Nothing
    frmFrontEnd.ReShow
    
errOut:
End Sub



'
' this is the main loop of the game
'
Private Sub MainLoop()
Dim bHasChangedWorlds As Boolean
Dim lEscapeCount As Long
Dim i As Long

    mapExitToLoadGame = False

    LoadSavedGame
    bGameOver = False
    lSaveTime = 0
    LoadWorldList
    LoadBlankWorldPassInfo
    
returnToVeryTop:
    
    If Not ShowLoadGameScreen(True) Then Exit Sub
    
returnToTop:
    
    Do
        If levelReturnValue <> lvrPipe Then
            Do
                bHasChangedWorlds = False
                mapExitToLoadGame = False
                curPlayLevel = ShowWorldMap(curWorld, IIf(levelReturnValue = lvrPipeOut, nextMapNode, -1)) '-1 means same one
                If mapExitToLoadGame Then Exit Do
                If curPlayLevel = -1 Then
                    If mapExitToNewWorld > 0 Then
                        levelReturnValue = lvrPipeOut
                        nextMapNode = -10 - mapExitWarpID  'if next map node is between -11 and -20 the map starts at warp number abs(node)-10
                        curWorld = mapExitToNewWorld - 1
                        cwdLoadWorldData App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(curWorld) & "\Worldata.3kw"
                        bHasChangedWorlds = True
                        ShowBlackTextScreen sWorldNameList(curWorld)
                    Else
                        Exit Do
                    End If
                Else
                    curLevel = curPlayLevel
                End If
            Loop While bHasChangedWorlds
        End If
        If curPlayLevel = -1 And mapExitToNewWorld <= 0 Then Exit Do
        If mapExitToLoadGame Then Exit Do
        
        If mapLevelEntryMode > -1 And levelReturnValue <> lvrPipe Then
            Select Case mapLevelEntryMode
                Case 0
                    mpExitPipe = mpExitUp
                Case 1
                    mpExitPipe = mpExitDown
                Case 2
                    mpExitPipe = mpExitLeft
                Case 3
                    mpExitPipe = mpExitRight
                Case 4
                    mpExitPipe = mpExitDoor
            End Select
            nextStartPipe = mapLevelEntryTag
        End If
        
        If surfList.Backdrop <> 0 Then GFX.DestroySurface surfList.Backdrop
        surfList.Backdrop = GFX.CreateSurface(App.Path & "\Data\Backdrops\" & oCurWorldData.LevelData(curLevel).Background, 0, False, True)
        If surfList.Tileset <> 0 Then GFX.DestroySurface surfList.Tileset
        surfList.Tileset = GFX.CreateSurface(App.Path & "\Data\Tilesets\" & oCurWorldData.LevelData(curLevel).TileFile & ".bmp", 0, False)
        LoadLevel App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(curWorld) & "\" & oCurWorldData.LevelData(curLevel).Filename & ".3kl", App.Path & "\Data\Tilesets\" & oCurWorldData.LevelData(curLevel).TileFile & ".3kt"
        LoadEnemySkinFile App.Path & "\Data\Enemies\" & oCurWorldData.LevelData(curLevel).EnemySkin & ".3ke"
        LoadEnemySurfaces
        
        HandleLevelEntry
        MainGameLoop oCurWorldData.LevelData(curLevel).MusicFile, nextStartX, nextStartY, lSaveTime, mpExitPipe
        HandleLevelExit
        
        If levelReturnValue = lvrQuit Or bGameOver Then Exit Do
        
    Loop
    
    If bGameOver Then
        With GFX
            PlaySound Sounds.GameOverSound
            For i = 100 To 0 Step -1
            .BeginScene 25
                .DrawSurface surfList.GameOverScreen, 0, 0, 640, 480, 0, 0
                .DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
            .EndScene
            Next i
            For i = 0 To 100
            .BeginScene 25
                .DrawSurface surfList.GameOverScreen, 0, 0, 640, 480, 0, 0
            .EndScene
            Next i
            For i = 0 To 100
            .BeginScene 25
                .DrawSurface surfList.GameOverScreen, 0, 0, 640, 480, 0, 0
                .DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
            .EndScene
            Next i
        End With
        bGameOver = False
    End If
    
    If mapExitToLoadGame Then If ShowLoadGameScreen(False) Then GoTo returnToTop Else Exit Sub
    GoTo returnToVeryTop
    
End Sub



Private Sub HandleLevelEntry()
    If mpExitPipe = mpNothing Then
        nextStartX = GetLevelStartX
        nextStartY = GetLevelStartY
    ElseIf mpExitPipe = mpExitUp Then
        nextStartX = GetLevelPipeX(nextStartPipe)
        nextStartY = GetLevelPipeY(nextStartPipe)
    ElseIf mpExitPipe = mpExitDown Then
        nextStartX = GetLevelPipeX(nextStartPipe)
        nextStartY = GetLevelPipeY(nextStartPipe) + IIf(gameMarioStat = MarioSmall, 64, 90)
    ElseIf mpExitPipe = mpExitLeft Then
        nextStartX = GetLevelPipeX(nextStartPipe) - 48
        nextStartY = GetLevelPipeY(nextStartPipe) + 32
    ElseIf mpExitPipe = mpExitRight Then
        nextStartX = GetLevelPipeX(nextStartPipe) + 16
        nextStartY = GetLevelPipeY(nextStartPipe) + 32
    ElseIf mpExitPipe = mpExitDoor Then
        nextStartX = GetLevelPipeX(nextStartPipe) - 16
        nextStartY = GetLevelPipeY(nextStartPipe) + 32
    End If
End Sub

Private Sub HandleLevelExit()
Dim oldCurLevel As Long

    nextStartPipe = 0
    Select Case levelReturnValue
        Case lvrAbort
            curLevel = curPlayLevel
            lSaveTime = 0
            mpExitPipe = mpNothing
        Case lvrPass
            curLevel = curPlayLevel
            lSaveTime = 0
            mpExitPipe = mpNothing
            oWorldPassData.bWorldPassed(curWorld).bLevelPassed(curLevel) = True
        Case lvrDie
            curLevel = curPlayLevel
            lSaveTime = 0
            mpExitPipe = mpNothing
            gameLives = gameLives - 1
            If gameLives < 0 Then bGameOver = True
        Case lvrPipe
            Select Case oCurWorldData.LevelData(curLevel).PipeDest(levelReturnPipe).destDir
                Case 0
                    mpExitPipe = mpExitLeft
                Case 1
                    mpExitPipe = mpExitRight
                Case 2
                    mpExitPipe = mpExitUp
                Case 3
                    mpExitPipe = mpExitDown
                Case 4
                    mpExitPipe = mpExitDoor
            End Select
            nextStartPipe = oCurWorldData.LevelData(curLevel).PipeDest(levelReturnPipe).destTag
            oldCurLevel = curLevel
            curLevel = oCurWorldData.LevelData(curLevel).PipeDest(levelReturnPipe).destLevel - 1
            lSaveTime = lTime
            If curLevel = -1 Then
                curLevel = curPlayLevel
                lSaveTime = 0
                mpExitPipe = mpNothing
                levelReturnValue = lvrPipeOut
                nextMapNode = oCurWorldData.LevelData(oldCurLevel).PipeDest(levelReturnPipe).destTag
            End If
    End Select
    
End Sub





'is true if returaning to game
Private Function ShowLoadGameScreen(bFirstRun As Boolean) As Boolean
Dim i As Long
Dim u As Long
Dim curSel As Long
Dim selMax As Long
Dim lKeyDown As Long
    
    Music.LoadFile App.Path & "\Mods\" & sWorldSetName & "\Menu.mid"
    Music.PlayMusic
    GFX.SetFont "Courier New", 8, True, False, False, False
    ShowLoadGameScreen = True
    lKeyDown = 0
    curSel = IIf(bFirstRun, 0, 6)
    selMax = IIf(bFirstRun, 6, 7)
    
    Do
        
        If GameKeyDown(Down) And lKeyDown <> GameKeys.Down Then
            curSel = curSel + 1
            If curSel > selMax Then curSel = 0
        End If
        If GameKeyDown(Up) And lKeyDown <> GameKeys.Up Then
            curSel = curSel - 1
            If curSel < 0 Then curSel = selMax
        End If
        If GameKeyDown(Down) Then
            lKeyDown = GameKeys.Down
        ElseIf GameKeyDown(Up) Then
            lKeyDown = GameKeys.Up
        Else
            lKeyDown = GameKeys.Right
        End If
    
        GFX.BeginScene 25
            GFX.DrawSurface surfList.TitleScreen, 0, 0, 640, 480, 0, 0
            GFX.DrawSurface surfList.MenuObjects, 0, 0, 256, 160, 192, 288
            GFX.DrawSurface surfList.MenuObjects, 256, curSel * 16, 16, 16, 204, 296 + (18 * curSel)
            GFX.DrawText "Begin a New Game", 224, 296, vbWhite
            For i = 1 To selMax
                If i <= 5 Then
                    GFX.DrawText oSaveGameData.gameSlot(i).sSaveString, 224, 296 + (18 * i), vbWhite
                ElseIf i = 6 Then
                    If bFirstRun Then
                        GFX.DrawText "Quit", 224, 296 + (18 * i), vbWhite
                    Else
                        GFX.DrawText "Return to Game", 224, 296 + (18 * i), vbWhite
                    End If
                ElseIf i = 7 Then
                    GFX.DrawText "Quit", 224, 296 + (18 * i), vbWhite
                End If
            Next i
        GFX.EndScene
        
        If GameKeyDown(Jump) Then
            If curSel < 6 Then
                If curSel = 0 Then
                    LoadBlankWorldPassInfo
                    curWorld = 0
                    localWorld = 0
                    nextMapNode = -2
                    gameMarioStat = MarioSmall
                    marioReserveItem = rsvEmpty
                    gameCoins = 0
                    gameLives = 3
                    gameGreens = 3
                    blBossesPassed = BP_NONE
                ElseIf (oSaveGameData.gameSlot(curSel).sSaveString = "[EMPTY]" Or oSaveGameData.gameSlot(curSel).lastWorldSaved <= 9) Then
                    LoadBlankWorldPassInfo
                    curWorld = 0
                    localWorld = 0
                    nextMapNode = -2
                    gameMarioStat = MarioSmall
                    marioReserveItem = rsvEmpty
                    gameCoins = 0
                    gameLives = 3
                    gameGreens = 3
                    blBossesPassed = BP_NONE
                Else
                    oWorldPassData = oSaveGameData.gameSlot(curSel).worldPassedData
                    curWorld = oSaveGameData.gameSlot(curSel).lastWorldSaved - 10
                    localWorld = curWorld
                    nextMapNode = oSaveGameData.gameSlot(curSel).lastMapNode
                    gameMarioStat = oSaveGameData.gameSlot(curSel).lastMarioStatus
                    marioReserveItem = oSaveGameData.gameSlot(curSel).lastMarioReserve
                    gameCoins = oSaveGameData.gameSlot(curSel).lastMarioCoins
                    gameLives = oSaveGameData.gameSlot(curSel).lastMarioLives
                    gameGreens = oSaveGameData.gameSlot(curSel).lastMarioGreens
                    blBossesPassed = oSaveGameData.gameSlot(curSel).lastBossesPassed
                End If
                levelReturnValue = lvrPipeOut
                cwdLoadWorldData App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(curWorld) & "\Worldata.3kw"
                Music.StopMusic
                ShowBlackTextScreen sWorldNameList(curWorld)
                Exit Do
            ElseIf curSel = 6 Then
                curSel = -1
                Exit Do
            ElseIf curSel = 7 Then
                ShowLoadGameScreen = False
                Exit Do
            End If
        ElseIf GameKeyDown(Quit) Then
            curSel = -1
            Exit Do
        End If
        
    Loop
    If bFirstRun Then ShowLoadGameScreen = (curSel <> -1)
    
    Music.StopMusic

End Function

