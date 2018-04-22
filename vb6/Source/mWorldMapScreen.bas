Attribute VB_Name = "mWorldMapScreen"
Option Explicit

Private Enum udeDirection
    dUp
    dDown
    dLeft
    dRight
End Enum

Private Const ANIM_SPEED As Long = 10  'lower the quicker
Private Const MarioMapSpeed As Single = 5

Public localWorld As Long

Private nodeMap As oWorldMap
Private mapWidth As Long
Private mapHeight As Long
Private lastNode As Long
Private curNode As Long
Private nextNode As Long

Private xOffset As Long
Private yOffset As Long
Private xMario As Single
Private yMario As Single
Private lMarioFrame As Long
Private bMarioFrame As Boolean

Private xSpeed As Single
Private ySpeed As Single
Private bGoing As Boolean
Private dirMario As udeDirection
Private levelBlip As Long
Private lImageFrame As Long
Private lImageFrameCount As Long

Public mapLevelEntryMode As Long
Public mapLevelEntryTag As Long
Public mapExitToNewWorld As Long
Public mapExitWarpID As Long
Public mapExitToLoadGame As Boolean



Public Function ShowWorldMap(lWorld As Long, Optional ByVal startNode As Long = -1) As Long
Dim i As Long, u As Long
Dim testTag As Byte

    mapExitToNewWorld = 0
    mapExitWarpID = 0
    
    GFX.SetFont "Courier New", 12, True, False, False, False
    bTimingTime = False
    
    'position mario according to function input
    If startNode > -1 And curNode <> startNode Then
        curNode = startNode
        lastNode = -1
    ElseIf startNode = -2 Then
        curNode = 0
        lastNode = -1
    End If

    'load map
    If localWorld <> lWorld Or curNode = 0 Then curNode = 1
    localWorld = lWorld
    Set nodeMap = New oWorldMap
    nodeMap.loadMap App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(localWorld) & "\Worldmap.3km"
    If surfList.WorldMap <> 0 Then GFX.DestroySurface surfList.WorldMap
    surfList.WorldMap = GFX.CreateSurface(App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(lWorld) & "\Worldmap.gif", 1, False, True)
    If surfList.WorldImages <> 0 Then GFX.DestroySurface surfList.WorldImages
    surfList.WorldImages = GFX.CreateSurface(App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(lWorld) & "\MapObjects.bmp", 0, False, False)
    frmMain.picLoadSize.Picture = LoadPicture(App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(lWorld) & "\Worldmap.gif")
    mapWidth = frmMain.picLoadSize.width
    mapHeight = frmMain.picLoadSize.height
    
    'position mario if hes just entering the world
    If startNode < -10 And startNode >= -20 Then
        lastNode = -1
        For i = 1 To nodeMap.nodeCount
            If nodeMap.NodeTag(i) = Abs(startNode) + 10 Then curNode = i
        Next i
    End If

    'init vars
    xMario = nodeMap.xPos(curNode)
    yMario = nodeMap.yPos(curNode)
    lMarioFrame = 0
    levelBlip = 0
    bGoing = False
    
    MusicLoadFile App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(localWorld) & "\Worldmap.mid"
    MusicPlayMusic
    
    showFadeIn False
    
    Do
    GFX.BeginScene 25
    
        UpdateInput
        If GameKeyDown(QuitFromPauseKey) Then DebugModeEnabled = False
        If DebugKeysPressed Then DebugModeEnabled = True
        If GameKeyDown(DebugA) Then showSaveScreen True
        'If GameKeyDown(DebugA) And DebugModeEnabled Then If gameMarioStat < 4 Then gameMarioStat = gameMarioStat + 1 Else gameMarioStat = 0
    
        levelBlip = levelBlip + 10
        If levelBlip > 360 Then levelBlip = levelBlip - 360
        
        lImageFrameCount = lImageFrameCount + 1
        If lImageFrameCount >= ANIM_SPEED Then
            lImageFrameCount = 0
            lImageFrame = lImageFrame + 1
            If lImageFrame >= 4 Then lImageFrame = 0
        End If
    
        If bGoing Then
        
            lMarioFrame = lMarioFrame + 2
            xMario = xMario - xSpeed
            yMario = yMario - ySpeed
            If findDist(xMario, yMario, nodeMap.xPos(nextNode), nodeMap.yPos(nextNode)) < MarioMapSpeed Then
                lastNode = curNode
                curNode = nextNode
                xMario = nodeMap.xPos(curNode)
                yMario = nodeMap.yPos(curNode)
                If nodeMap.PassThru(curNode) Then
                    Select Case dirMario
                        Case dUp
                            nextNode = nodeMap.UpNode(curNode)
                        Case dDown
                            nextNode = nodeMap.DownNode(curNode)
                        Case dLeft
                            nextNode = nodeMap.LeftNode(curNode)
                        Case dRight
                            nextNode = nodeMap.RightNode(curNode)
                    End Select
                    getSpeeds
                Else
                    xSpeed = 0
                    ySpeed = 0
                    bGoing = False
                End If
            End If
            
        Else

            lMarioFrame = lMarioFrame + 1
            If nodeMap.NodeTag(curNode) > 40 And nodeMap.NodeTag(curNode) <= 110 And GameKeyDown(DebugB) And DebugModeEnabled Then
                oWorldPassData.bWorldPassed(curWorld).bLevelPassed(nodeMap.NodeTag(curNode) - 41) = True
            End If
            
            For u = 0 To 0
            If GameKeyDown(Left) And nodeMap.LeftNode(curNode) > 0 Then
                If nodeMap.NodeTag(curNode) > 40 And nodeMap.NodeTag(curNode) <= 110 Then _
                    If Not oWorldPassData.bWorldPassed(curWorld).bLevelPassed(nodeMap.NodeTag(curNode) - 41) Then _
                      If lastNode <> -1 And nodeMap.LeftNode(curNode) <> lastNode Then Exit For
                nextNode = nodeMap.LeftNode(curNode)
                bGoing = True
                dirMario = dLeft
                getSpeeds
            ElseIf GameKeyDown(Right) And nodeMap.RightNode(curNode) > 0 Then
                If nodeMap.NodeTag(curNode) > 40 And nodeMap.NodeTag(curNode) <= 110 Then _
                    If Not oWorldPassData.bWorldPassed(curWorld).bLevelPassed(nodeMap.NodeTag(curNode) - 41) Then _
                        If lastNode <> -1 And nodeMap.RightNode(curNode) <> lastNode Then Exit For
                nextNode = nodeMap.RightNode(curNode)
                bGoing = True
                dirMario = dRight
                getSpeeds
            ElseIf GameKeyDown(Up) And nodeMap.UpNode(curNode) > 0 Then
                If nodeMap.NodeTag(curNode) > 40 And nodeMap.NodeTag(curNode) <= 110 Then _
                    If Not oWorldPassData.bWorldPassed(curWorld).bLevelPassed(nodeMap.NodeTag(curNode) - 41) Then _
                        If lastNode <> -1 And nodeMap.UpNode(curNode) <> lastNode Then Exit For
                nextNode = nodeMap.UpNode(curNode)
                bGoing = True
                dirMario = dUp
                getSpeeds
            ElseIf GameKeyDown(Down) And nodeMap.DownNode(curNode) > 0 Then
                If nodeMap.NodeTag(curNode) > 40 And nodeMap.NodeTag(curNode) <= 110 Then _
                    If Not oWorldPassData.bWorldPassed(curWorld).bLevelPassed(nodeMap.NodeTag(curNode) - 41) Then _
                        If lastNode <> -1 And nodeMap.DownNode(curNode) <> lastNode Then Exit For
                nextNode = nodeMap.DownNode(curNode)
                bGoing = True
                dirMario = dDown
                getSpeeds
            ElseIf GameKeyDown(Jump) Then
                testTag = nodeMap.NodeTag(curNode)
                If testTag > 0 And testTag < 21 Then
                    For i = 1 To nodeMap.nodeCount
                        If nodeMap.NodeTag(i) = testTag And i <> curNode Then
                            PlaySound Sounds.Pipe
                            showFadeOut
                            xMario = nodeMap.xPos(i)
                            yMario = nodeMap.yPos(i)
                            curNode = i
                            showFadeIn
                            Exit For
                        End If
                    Next i
                ElseIf testTag > 20 And testTag <= 40 Then
                    mapExitToNewWorld = nodeMap.warpExitWorld(curNode)
                    mapExitWarpID = testTag - 20
                    ShowWorldMap = 0
                    PlaySound Sounds.Pipe
                    drawMap
                    DrawBonusMeter
                    GFX.EndScene
                    Exit Do
                ElseIf testTag > 40 And testTag <= 110 Then
                    ShowWorldMap = testTag - 40
                    bTimingTime = GameKeyDown(Shoot)
                    PlaySound IIf(bTimingTime, Sounds.PowLoop, Sounds.EnterLevel)
                    drawMap
                    DrawBonusMeter
                    GFX.EndScene
                    Exit Do
                ElseIf testTag = 111 Then
                    showSaveScreen
                End If
            ElseIf GameKeyDown(Quit) Or GameKeyDown(Pause) Then
                mapExitToLoadGame = True
                levelReturnValue = lvrAbort
                GFX.EndScene
                Exit Do
            End If
            Next u
            
        End If
        
        If lMarioFrame > 10 Then
            bMarioFrame = Not bMarioFrame
            lMarioFrame = 0
        End If
        drawMap
        DrawBonusMeter
        
    GFX.EndScene
    Loop
    
    showFadeOut False
    
    mapLevelEntryMode = nodeMap.EntryMode(curNode)
    mapLevelEntryTag = nodeMap.EntryTag(curNode)
    
    MusicStopMusic
    Set nodeMap = Nothing
    ShowWorldMap = ShowWorldMap - 1
    
    If ShowWorldMap >= 0 Then ShowBlackTextScreen oCurWorldData.LevelData(ShowWorldMap).LevelName
    
End Function





Private Sub drawMap()
Dim xSrc As Single
Dim tst As Single
Dim i As Long

    If xMario > 320 And xMario < mapWidth - 320 Then
        xOffset = xMario - 320
    ElseIf xMario <= 320 Then
        xOffset = 0
    ElseIf xMario >= mapWidth - 320 Then
        xOffset = mapWidth - 640
    End If
    
    If yMario > 240 And yMario < mapHeight - 240 Then
        yOffset = yMario - 240
    ElseIf yMario <= 240 Then
        yOffset = 0
    ElseIf yMario >= mapHeight - 240 Then
        yOffset = mapHeight - 480
    End If
    
    Select Case gameMarioStat
        Case MarioSmall: xSrc = 0
        Case MarioBig: xSrc = 1
        Case MarioFlower: xSrc = 2
        Case MarioMoonboot: xSrc = 3
        Case MarioHammer: xSrc = 4
    End Select
        
    GFX.DrawSurface surfList.WorldMap, CSng(xOffset), CSng(yOffset), 640, 480, 0, 0
    For i = 0 To nodeMap.nodeCount
        If nodeMap.NodeImage(i) > 0 Then
            tst = (nodeMap.NodeImage(i) - 1) * 32
            If nodeMap.NodeTag(i) > 40 And nodeMap.NodeTag(i) <= 110 Then If oWorldPassData.bWorldPassed(curWorld).bLevelPassed(nodeMap.NodeTag(i) - 41) Then tst = tst + 32
            GFX.DrawSurface surfList.WorldImages, tst, lImageFrame * 32, 32, 32, nodeMap.xPos(i) - xOffset - 16, nodeMap.yPos(i) - yOffset - 16
        End If
    Next i
    GFX.DrawSurface surfList.MarioMap, IIf(bMarioFrame, 32, 0) + (xSrc * 64), 0, 32, 48, xMario - 16 - xOffset, yMario - 32 - yOffset
    If DebugModeEnabled Then
        GFX.SetFont "Courier New", 8, True, False, False, False
        GFX.DrawText "  DEBUG MODE ENABLED  ", 0, 0, vbWhite, 0
    End If
        
End Sub



Private Sub showFadeOut(Optional ByVal insideScene As Boolean = True)
Dim i As Long
    If insideScene Then GFX.EndScene
    If bTimingTime Then
        Do
        GFX.BeginScene 25
            drawMap
            DrawBonusMeter
            GFX.SetFont "Courier New", 12, True, False, False, False
            GFX.DrawText "Timing this level", 200, 100, 0
            GFX.DrawText "Get Ready!", 230, 120, 0
        GFX.EndScene
        Loop While SoundPlaying(Sounds.PowLoop)
        PlaySound Sounds.EnterLevel
    End If
    For i = 0 To 100 Step 4
    GFX.BeginScene 25
        drawMap
        DrawBonusMeter
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
        If bTimingTime Then
            GFX.SetFont "Courier New", 12, True, False, False, False
            GFX.DrawText "Timing this level", 200, 100, 0
            GFX.DrawText "Get Ready!", 230, 120, 0
        End If
    GFX.EndScene
    Next i
    If insideScene Then GFX.BeginScene 25
End Sub

Private Sub showFadeIn(Optional ByVal insideScene As Boolean = True)
Dim i As Long
    If insideScene Then GFX.EndScene
    For i = 100 To 0 Step -4
        GFX.BeginScene 25
        drawMap
        DrawBonusMeter
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
        GFX.EndScene
    Next i
    If insideScene Then GFX.BeginScene 25
End Sub




Private Sub getSpeeds()
Dim sV As Single
    sV = findDist(nodeMap.xPos(curNode), nodeMap.yPos(curNode), nodeMap.xPos(nextNode), nodeMap.yPos(nextNode)) / MarioMapSpeed
    xSpeed = (nodeMap.xPos(curNode) - nodeMap.xPos(nextNode)) / sV
    ySpeed = (nodeMap.yPos(curNode) - nodeMap.yPos(nextNode)) / sV
End Sub

Private Function findDist(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Single
findDist = Math.Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
End Function



'Private Sub DrawMapLevelName()
'Dim sLevName As String
'    If nodeMap.NodeTag(curNode) = 101 Then
'        sLevName = "- Save Game -"
'    Else
'        sLevName = oCurWorldData.LevelData(nodeMap.NodeTag(curNode) - 41).LevelName
'    End If
'    If Mid$(sLevName, 1, 3) = "---" Then Exit Sub
'    GFX.SetFont "Courier New", 8, True, False, False, False
'    GFX.DrawText " " & sLevName & " ", 0, 64, vbWhite, 0
'End Sub




Private Sub showSaveScreen(Optional ByVal insideScene As Boolean = True)
Dim i As Long, u As Long, curSel As Long, lKeyDown As Long
Dim bSavedGameLoaded As Boolean

    If insideScene Then GFX.EndScene
    For i = 0 To 100 Step 4
    GFX.BeginScene 25
        drawMap
        DrawBonusMeter
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
    GFX.EndScene
    Next i
    
    GFX.SetFont "Courier New", 8, True, False, False, False
    lKeyDown = 0
    curSel = 1

    DetermineLevelPassedCount
    
    Do
    
        If GameKeyDown(Down) And lKeyDown <> GameKeys.Down Then
            curSel = curSel + 1
            If curSel > 6 Then curSel = 1
        End If
        If GameKeyDown(Up) And lKeyDown <> GameKeys.Up Then
            curSel = curSel - 1
            If curSel < 1 Then curSel = 6
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
            GFX.DrawText "Select a Save Slot", 224, 296, vbRed
            For i = 1 To 6
                If i <= 5 Then
                    GFX.DrawText oSaveGameData.gameSlot(i).sSaveString, 224, 296 + (18 * i), vbWhite
                ElseIf i = 6 Then
                    GFX.DrawText "Return to Game", 224, 296 + (18 * i), vbWhite
                End If
            Next i
        GFX.EndScene
            
        If GameKeyDown(Jump) Then
            If curSel < 6 Then
                With oSaveGameData.gameSlot(curSel)
                    .worldPassedData = oWorldPassData
                    .lastWorldSaved = curWorld + 10
                    .lastMapNode = curNode
                    .sSaveString = CStr(Date) & " - Progress:" & CStr(lLevelPassedCount) 'sWorldNameList(oSaveGameData.gameSlot(i).lastWorldSaved - 10)
                    .lastMarioStatus = gameMarioStat
                    .lastMarioReserve = marioReserveItem
                    .lastMarioCoins = gameCoins
                    .lastMarioLives = gameLives
                    .lastMarioGreens = gameGreens
                    .lastBossesPassed = blBossesPassed
                End With
                SaveGameData
                For u = 0 To 100
                    GFX.BeginScene 25
                        GFX.DrawSurface surfList.TitleScreen, 0, 0, 640, 480, 0, 0
                        GFX.DrawSurface surfList.MenuObjects, 0, 0, 256, 160, 192, 288
                        GFX.DrawText "Select a Save Slot", 224, 296, vbRed
                        For i = 1 To 6
                            If i <= 5 Then
                                GFX.DrawText oSaveGameData.gameSlot(i).sSaveString, 224, 296 + (18 * i), vbWhite
                            ElseIf i = 6 Then
                                GFX.DrawText "Game Saved!", 224, 296 + (18 * i), vbRed
                            End If
                        Next i
                    GFX.EndScene
                Next u
            End If
            Exit Do
        End If
        
        If GameKeyDown(Quit) Then Exit Do
        
    Loop
    
    
    For i = 100 To 0 Step -4
    GFX.BeginScene 25
        drawMap
        DrawBonusMeter
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, , CSng(i), 0, 0, 0
    GFX.EndScene
    Next i
    If insideScene Then GFX.BeginScene 25

End Sub


