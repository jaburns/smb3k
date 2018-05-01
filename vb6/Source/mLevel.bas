Attribute VB_Name = "mLevel"
Option Explicit

Private curOLevel As New oLevel
Private curTiles As udtLTileSetData

Private Const ANIM_SPEED As Long = 10  'lower the quicker
Private curTileFrameCount As Long
Public curTileFrame As Long

Public sEndWheelAngle As Single
Public sEndWheelSpeed As Single
Public lEndBlockX As Long
Public lEndBlockY As Long
Public lTimeIncCount As Long

Private bgX1 As Long
Private bgX2 As Long

Private lPipeX(1 To 16) As Long
Private lPipeY(1 To 16) As Long

Public screenLeft As Long
Public screenTop As Long

Public Function DecrementTime() As Long

    lTimeIncCount = lTimeIncCount + 1
    If lTimeIncCount = 30 Then
        lTimeIncCount = 0
        curOLevel.SetTime curOLevel.Time - 1
    End If
    DecrementTime = curOLevel.Time

End Function

Public Sub SetLevelTime(lTime As Long, Optional ByVal resetMaxTime As Boolean = False)
    curOLevel.SetTime lTime
    If resetMaxTime Then lMaxTime = lTime
End Sub


Async Public Sub LoadLevel(sPath As String, sTilesetPath As String)
    bgX1 = 0
    bgX2 = 0
    curOLevel.LoadFromFile sPath
    LoadLevelTileset curTiles, sTilesetPath
    sEndWheelAngle = 0
    sEndWheelSpeed = 10
    FindLevelPipeLocations
End Sub


Public Sub DrawLevel(ByVal xScreen As Long, ByVal yScreen As Long, Optional ByVal backdropOnly As Boolean = False)
Dim xx As Long
Dim yy As Long
Dim getY As Long
Dim xdraw As Single
Dim yDraw As Single
Dim StartX As Long
Dim StartY As Long
Dim bXScroll As Boolean
Dim bYScroll As Boolean

    curTileFrameCount = curTileFrameCount + 1
    If curTileFrameCount = ANIM_SPEED Then
        curTileFrameCount = 0
        curTileFrame = curTileFrame + 1
        If curTileFrame = 4 Then curTileFrame = 0
    End If
     
    screenLeft = xScreen
    screenTop = yScreen
    
    'some calculations for horizontal scrolling
    bXScroll = False
    If xScreen <= 0 Then
        If xScreen < 0 Then xScreen = 0
        StartX = 0
    ElseIf xScreen >= GetLevelWidth() - 640 Then
        If xScreen > GetLevelWidth() - 640 Then xScreen = GetLevelWidth() - 640
        StartX = (GetLevelWidth() / 32) - 20
    Else
        bgX1 = (xScreen Mod (640 * 2)) / 2
        bgX2 = (xScreen Mod (640 * 4)) / 4
        StartX = __intDiv(xScreen , 32)
        bXScroll = True
    End If

    'some calculations for vertical scrolling
    bYScroll = False
    If yScreen <= 0 Then
        If yScreen < 0 Then yScreen = 0
        StartY = 0
    ElseIf yScreen >= GetLevelHeight() - 480 Then
        If yScreen > GetLevelWidth() - 480 Then yScreen = GetLevelWidth() - 480
        StartY = (GetLevelHeight() / 32) - 15
    Else
        StartY = __intDiv(yScreen , 32)
        bYScroll = True
    End If
    
    'draw the backdrop
    GFX.DrawSurface surfList.Backdrop, CSng(bgX2), 0, 640 - CSng(bgX2), 480, 0, 0
    GFX.DrawSurface surfList.Backdrop, 0, 0, CSng(bgX2), 480, 640 - CSng(bgX2), 0
    GFX.DrawSurface surfList.Backdrop, CSng(bgX1), 480, 640 - CSng(bgX1), 480, 0, 0
    GFX.DrawSurface surfList.Backdrop, 0, 480, CSng(bgX1), 480, 640 - CSng(bgX1), 0
    
If backdropOnly Then Exit Sub
    
    'spin the wheel of exit
    sEndWheelAngle = sEndWheelAngle + sEndWheelSpeed
    If sEndWheelAngle >= 360 Then sEndWheelAngle = sEndWheelAngle - 360
    
Dim tmpType As udeLTileType

    lEndBlockX = -1
    lEndBlockY = -1

    'draw the effin map
    With curOLevel
    For xx = 0 To 20
    
        If bXScroll Then xdraw = (xx * 32) - (xScreen Mod 32) Else xdraw = xx * 32
        
        For yy = 0 To 15
        tmpType = GetLevelTileType(curTiles, .xSrc(xx + StartX, yy + StartY), .ySrc(xx + StartX, yy + StartY))
        
            If bYScroll Then yDraw = (yy * 32) - (yScreen Mod 32) Else yDraw = yy * 32
            
            If .IsTile(xx + StartX, yy + StartY) And .TileEnemy(xx + StartX, yy + StartY) <> OBJHIDETILE Then
                If GetTileAnim(xx + StartX, yy + StartY) Then getY = curTileFrame * 32 Else getY = 0
                GFX.DrawSurface surfList.Tileset, .xSrc(xx + StartX, yy + StartY) * 32, getY + (.ySrc(xx + StartX, yy + StartY) * 32), 32, 32, xdraw, yDraw
                'draw end wheel
                If .TileTag(xx + StartX, yy + StartY) = ENDWHEEL Then
                    lEndBlockX = xdraw
                    lEndBlockY = yDraw
                    GFX.DrawSurface surfList.Objects, 83, 66, 128, 128, xdraw - 48, yDraw - 160, , , sEndWheelAngle
                End If
            End If
                
        Next yy
        
    Next xx
    
    'draw the end wheel if it's on the screen
    If lEndBlockX > 0 And lEndBlockY > 0 Then
        GFX.DrawSurface surfList.Objects, 83, 66, 128, 128, lEndBlockX - 48, lEndBlockY - 160, , , sEndWheelAngle
    End If
    
    End With

End Sub

Public Function GetLevelSideWarp() As Boolean
    GetLevelSideWarp = curOLevel.SideWarp
End Function
Public Function GetLevelStartX() As Long
    GetLevelStartX = (curOLevel.StartX * 32) + 16
End Function
Public Function GetLevelStartY() As Long
    GetLevelStartY = (curOLevel.StartY * 32) + 32
End Function


Public Sub SetTile(X As Long, Y As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional bOffset As Boolean = False)
    If bOffset Then
        xSrc = xSrc + curOLevel.xSrc(__intDiv(X , 32), __intDiv(Y , 32))
        ySrc = ySrc + curOLevel.ySrc(__intDiv(X , 32), __intDiv(Y , 32))
    End If
    curOLevel.SetTile __intDiv(X , 32), __intDiv(Y , 32), xSrc, ySrc
End Sub
Public Sub SetTileTag(X As Long, Y As Long, tag As udeLTileTag)
    curOLevel.SetTileTag __intDiv(X , 32), __intDiv(Y , 32), tag
End Sub
Public Sub SetTileEnemy(X As Long, Y As Long, tEnemy As udeLTileEnemy)
    curOLevel.SetTileEnemy __intDiv(X , 32), __intDiv(Y , 32), tEnemy
End Sub

Public Function GetTileTagAtPoint(X As Long, Y As Long) As udeLTileTag
    GetTileTagAtPoint = curOLevel.TileTag(__intDiv(X , 32), __intDiv(Y , 32))
End Function
Public Function GetTileEnemy(X As Long, Y As Long) As udeLTileEnemy
    GetTileEnemy = curOLevel.TileEnemy(X, Y)
End Function

Public Sub KillTile(X As Long, Y As Long)
    curOLevel.EraseTile __intDiv(X , 32), __intDiv(Y , 32)
End Sub

Public Function GetLevelWidth() As Long
    GetLevelWidth = (curOLevel.width + 1) * 32
End Function
Public Function GetLevelHeight() As Long
    GetLevelHeight = (curOLevel.height + 1) * 32
End Function

Public Function GetTileAnim(X As Long, Y As Long) As Boolean
    If curOLevel.IsTile(X, Y) Then
        GetTileAnim = GetLevelTileAnimated(curTiles, curOLevel.xSrc(X, Y), curOLevel.ySrc(X, Y))
    Else
        GetTileAnim = False
    End If
End Function


Public Function GetTile(X As Long, Y As Long) As udeLTileType
    If curOLevel.IsTile(X, Y) Then
        GetTile = GetLevelTileType(curTiles, curOLevel.xSrc(X, Y), curOLevel.ySrc(X, Y))
    Else
        GetTile = Background
    End If
End Function


Public Function GetTileXSrcPoint(xIn As Long, yIn As Long) As Long
GetTileXSrcPoint = curOLevel.xSrc(__intDiv(xIn , 32), __intDiv(yIn , 32))
End Function
Public Function GetTileYSrcPoint(xIn As Long, yIn As Long) As Long
GetTileYSrcPoint = curOLevel.ySrc(__intDiv(xIn , 32), __intDiv(yIn , 32))
End Function



Public Function GetTileAtPoint(X As Long, Y As Long) As udeLTileType
    GetTileAtPoint = GetTile(__intDiv(X , 32), __intDiv(Y , 32))
End Function
Public Function TileExistsAtPoint(X As Long, Y As Long) As Boolean
    TileExistsAtPoint = curOLevel.IsTile(__intDiv(X , 32), __intDiv(Y , 32))
End Function


Public Sub KillLevel()
    Set curOLevel = Nothing
End Sub





Private Sub FindLevelPipeLocations()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim TMP As udeLTileTag
    
    For i = 1 To 16
        lPipeX(i) = -1
        lPipeY(i) = -1
    Next i
    For X = 0 To curOLevel.width
        For Y = 0 To curOLevel.height
            TMP = curOLevel.TileTag(X, Y)
            If TMP >= 8 And TMP <= 23 Then
                TMP = TMP - 7
                lPipeX(TMP) = X * 32 + 32
                lPipeY(TMP) = Y * 32
            End If
        Next Y
    Next X

End Sub
Public Function GetLevelPipeX(lPipeID As Long) As Long
    If lPipeID < 1 Or lPipeID > 16 Then Exit Function
    GetLevelPipeX = lPipeX(lPipeID)
End Function
Public Function GetLevelPipeY(lPipeID As Long) As Long
    If lPipeID < 1 Or lPipeID > 16 Then Exit Function
    GetLevelPipeY = lPipeY(lPipeID)
End Function






Public Sub HitSpecialBlock(lXPos As Long, lYPos As Long, bHitFromLeft As Boolean, Optional ByVal bNoBreak As Boolean = False)
Dim tType As udeLTileType
Dim tTag As udeLTileTag

        hitManyCoinBlock lXPos, lYPos

        tType = GetTileAtPoint(lXPos, lYPos)
        If tType <> COINBLOCK And tType <> BRICK Then Exit Sub
        tTag = GetTileTagAtPoint(lXPos, lYPos)
        
        If tTag = 0 Then
            If tType = COINBLOCK Then
                SetTile lXPos, lYPos, 1, 0, True
                MakeLittleCoin (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32
                gameCoins = gameCoins + 1
            ElseIf Not bNoBreak Then
                KillTile lXPos, lYPos
                PlaySound Sounds.BreakBrick
                BreakBrick (__intDiv(lXPos , 32) * 32) + 16, (__intDiv(lYPos , 32) * 32) + 16
            End If
        Else
            SetTile lXPos, lYPos, 1, 0, True
            Select Case tTag
                Case ENDWHEEL
                    Mario.bHasWon = True
                Case FLOWER
                    If Mario.mStatus = MarioSmall Then
                        MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, 0
                    Else
                        MakeFlower (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32
                    End If
                Case PowerStar
                    MakeStar (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft
                Case MOONBOOT
                    If Mario.mStatus = MarioSmall Then
                        MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, 0
                    Else
                        MakeMoonboot (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft
                    End If
                Case POWERHAMMER
                    If Mario.mStatus = MarioSmall Then
                        MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, 0
                    Else
                        MakeHammerPickup (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32
                    End If
                Case HASWINGCAP
                    'If Mario.mStatus = MarioSmall Then
                    '    MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, 0, -5
                    'Else
                        MakeBlueStar (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32
                    'End If
                Case BEANSTALK
                    MakeViney (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32
                Case BRICKCOIN
                    MakeLittleCoin (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32
                    gameCoins = gameCoins + 1
                Case HAS1UP
                    MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, 1
                Case HAS1DOWN
                    MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, -1
                Case ONLYSHROOMS
                    MakeMushroom (__intDiv(lXPos , 32) * 32) + 16, __intDiv(lYPos , 32) * 32, bHitFromLeft, 0
            End Select
        End If
    
End Sub



Public Sub LevelBlockSwitch()
Dim X As Long
Dim Y As Long
Dim xTest As Long
Dim yTest As Long


    For X = 0 To (curOLevel.width + 0)
        For Y = 0 To (curOLevel.height + 0)
            If curOLevel.IsTile(X, Y) Then
            
                xTest = curOLevel.xSrc(X, Y)
                yTest = curOLevel.ySrc(X, Y)
                
                If xTest = oCurWorldData.LevelData(curLevel).dfCoin.xSrc And yTest = oCurWorldData.LevelData(curLevel).dfCoin.ySrc Then
                    SetTile X * 32, Y * 32, oCurWorldData.LevelData(curLevel).dfBrick.xSrc, oCurWorldData.LevelData(curLevel).dfBrick.ySrc
                ElseIf xTest = oCurWorldData.LevelData(curLevel).dfBrick.xSrc And yTest = oCurWorldData.LevelData(curLevel).dfBrick.ySrc Then
                    SetTile X * 32, Y * 32, oCurWorldData.LevelData(curLevel).dfCoin.xSrc, oCurWorldData.LevelData(curLevel).dfCoin.ySrc
                End If
                
            End If
        Next Y
    Next X
        
End Sub



Public Function isTileSolid(lX As Long, lY As Long, Optional ByVal testTop As Boolean = False) As Boolean
Dim testTile As udeLTileType
Dim testEnemy As udeLTileEnemy
Dim testTag As udeLTileTag
    isTileSolid = False
    testEnemy = GetTileEnemy(__intDiv(lX , 32), __intDiv(lY , 32))
    If testEnemy = OBJHIDETILE Then Exit Function
    testTile = GetTile(__intDiv(lX , 32), __intDiv(lY , 32))
    isTileSolid = (testTile = SOLID Or testTile = BOUNCY Or testTile = BRICK Or testTile = COINBLOCK Or testTile = DEADLY Or testTile = INJURETILE Or testTile = ICE)
    If Not isTileSolid Then
        isTileSolid = inManyCoinBlock(CSng(lX), CSng(lY))
        If isTileSolid Then Exit Function
        If testTop And (lY Mod 32 < 16) Then
            isTileSolid = (testTile = SOLIDONTOP)
            If isTileSolid Then Exit Function
            isTileSolid = inActiveDonut(CSng(lX), CSng(lY))
            If isTileSolid Then Exit Function
        End If
    End If
End Function

