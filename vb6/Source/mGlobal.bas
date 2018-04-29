Attribute VB_Name = "mGlobal"
Option Explicit

Public Enum udeBossPassList
    BP_NONE = 0
    BP_1 = 1
    BP_2 = 2
    BP_3 = 4
    BP_4 = 8
    BP_5 = 16
    BP_6 = 32
    BP_7 = 64
    BP_8 = 128
End Enum

Public Const ENEMYBOUNCE As Long = 4
Public Const ENEMYBIGBOUNCE As Long = 13
Public Const PI As Double = 3.14159265358979

'((Not supporting "Declare Function" in Rust interpreter. Just commenting it out here.))
'Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


'the main gfx engine
Public GFX As New DXGraphics

Public sMarioSkin As String

'current game data
Public sWorldSetName As String
Public sWorldList() As String
Public sWorldNameList() As String
Public curWorld As Long
Public Type udtLevelPassedList
    bLevelPassed() As Boolean
End Type
Public Type udtWorldPassData
    bWorldPassed() As udtLevelPassedList
End Type
Public oWorldPassData As udtWorldPassData

'saved game data
Public Type udtSaveGameSlot
    worldPassedData As udtWorldPassData
    sSaveString As String
    lastWorldSaved As Long
    lastMapNode As Long
    lastMarioStatus As Long
    lastMarioReserve As Long
    lastMarioLives As Long
    lastMarioCoins As Long
    lastMarioGreens As Long
    lastBossesPassed As Long
End Type
Public Type udtSaveGameData
    gameSlot(1 To 5) As udtSaveGameSlot
End Type
Public oSaveGameData As udtSaveGameData
Public lLevelPassedCount As Long
Public blBossesPassed As udeBossPassList

Public DebugModeEnabled As Boolean


'whenever the music changes temporarily there must be somehting to fall back on
Public sCurMusic As String


'simple type to hold x and y pos
Public Type XYPoint
    X As Single
    Y As Single
End Type

'items you can have in reserve
Public Enum udeMarioReserveItem
    rsvEmpty
    rsvMushroom
    rsvFlower
    rsvHammer
    rsvCarrot
    rsvWingcap
    rsvStar
End Enum

'holds marios status
Public Enum udeMarioStatus
    MarioSmall
    MarioBig
    MarioFlower
    MarioHammer
    MarioMoonboot
End Enum

'holds marios piping status
Public Enum udeMarioPipeStat
    mpNothing = 0
    mpEnterUp = 1
    mpEnterDown = 2
    mpEnterLeft = 3
    mpEnterRight = 4
    mpEnterDoor = 5
    mpExitUp = 6
    mpExitDown = 7
    mpExitLeft = 8
    mpExitRight = 9
    mpExitDoor = 10
End Enum


'list of all the surfaces used in the game
Public Type udtSurfacesEnemies
    Goomba As Long
    PiranaPlants As Long
    BuzzyBeetle As Long
    DumbKoopa As Long
    SmartKoopa As Long
    BumptyPenguin As Long
    Spiney As Long
    LittleBoo As Long
    RotoDisc As Long
    LavaBubble As Long
    Thwomp As Long
    DryBones As Long
    MovingPlatform As Long
    PowButton As Long
    FreeCheepCheep As Long
    BlockCheepCheep As Long
    RisingLava As Long
    MisterPokey As Long
    BouncyKoopa As Long
    SurfBulletBill As Long
    Bobomb As Long
    BOSSGoomboss As Long
    Wiggler As Long
    SavePointStar As Long
End Type
Public Type udtSurfaces
    Backdrop As Long
    Mario As Long
    MarioPipe As Long
    MarioCarry As Long
    MarioWings As Long
    MarioMap As Long
    Tileset As Long
    Objects As Long
    Sprites As Long
    Enemies As Long
    EnemyList As udtSurfacesEnemies
    WorldMap As Long
    FadeMask As Long
    WorldImages As Long
    BarIcons As Long
    NumberFont As Long
    PowerStar As Long
    TitleScreen As Long
    GameOverScreen As Long
    MenuObjects As Long
End Type
Public surfList As udtSurfaces


'game variables
Public gameLives As Long
Public gameCoins As Long
Public gameGreens As Long
Public gameMarioStat As udeMarioStatus


'((Not supporting "Declare Function" in Rust interpreter, or dealing with the mouse at all.))
'show and hide the mouse
Public Sub Mouse_Hide()
    Dim i As Long
End Sub
Public Sub Mouse_Show()
    Dim i As Long
End Sub


Public Sub LoadBlankWorldPassInfo()
Dim i As Long
    ReDim sWorldNameList(UBound(sWorldList))
    For i = 0 To UBound(sWorldList)
        cwdLoadWorldData App.Path & "\Mods\" & sWorldSetName & "\" & sWorldList(i) & "\Worldata.3kw"
        ReDim oWorldPassData.bWorldPassed(i).bLevelPassed(UBound(oCurWorldData.LevelData))
        sWorldNameList(i) = oCurWorldData.WorldName
    Next i
End Sub




'
'  This subs loads all the surfaces for the game from files
'
Public Sub InitSurfaces()
    
    surfList.Mario = GFX.CreateSurface(GetSkinPath("Standard.bmp"), 0, False)
    surfList.MarioWings = GFX.CreateSurface(GetSkinPath("Wings.bmp"), 0, False)
    surfList.MarioPipe = GFX.CreateSurface(GetSkinPath("Piping.bmp"), 0, False)
    surfList.MarioCarry = GFX.CreateSurface(GetSkinPath("Carrying.bmp"), 0, False)
    surfList.MarioMap = GFX.CreateSurface(GetSkinPath("Map.bmp"), 0, False)
    surfList.Objects = GFX.CreateSurface(GetSkinPath("Objects.bmp"), 0, True)
    surfList.Sprites = GFX.CreateSurface(GetSkinPath("Items.bmp"), 0, False)
    surfList.FadeMask = GFX.CreateSurface(App.Path & "\Data\Mask.bmp", 1, True)
    surfList.BarIcons = GFX.CreateSurface(GetSkinPath("Icons.bmp"), 0, False)
    surfList.NumberFont = GFX.CreateSurface(GetSkinPath("Font.bmp"), 0, False)
    'surfList.PowerStar = GFX.CreateSurface(App.Path & "\Data\PowerStar.bmp", 0, False)
    surfList.TitleScreen = GFX.CreateSurface(App.Path & "\Mods\" & sWorldSetName & "\Menu.gif", 1, False, True)
    surfList.GameOverScreen = GFX.CreateSurface(App.Path & "\Mods\" & sWorldSetName & "\GameOver.gif", 1, False, True)
    surfList.MenuObjects = GFX.CreateSurface(App.Path & "\Mods\" & sWorldSetName & "\MenuObjects.bmp", 0, False)
    
End Sub
Private Function GetSkinPath(sBitmapName As String)
    If Not fileExist(App.Path & "\Skins\" & sMarioSkin & "\" & sBitmapName) Then GetSkinPath = App.Path & "\Skins\Default\" & sBitmapName Else GetSkinPath = App.Path & "\Skins\" & sMarioSkin & "\" & sBitmapName
End Function


Public Sub LoadEnemySurfaces()
    With surfList.EnemyList
        If .Goomba <> 0 Then GFX.DestroySurface .Goomba
        .Goomba = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Goomba, 0, False)
        If .PiranaPlants <> 0 Then GFX.DestroySurface .PiranaPlants
        .PiranaPlants = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.PiranaPlants, 0, False)
        If .BuzzyBeetle <> 0 Then GFX.DestroySurface .BuzzyBeetle
        .BuzzyBeetle = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.BuzzyBeetle, 0, False)
        If .DumbKoopa <> 0 Then GFX.DestroySurface .DumbKoopa
        .DumbKoopa = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.DumbKoopa, 0, False)
        If .SmartKoopa <> 0 Then GFX.DestroySurface .SmartKoopa
        .SmartKoopa = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.SmartKoopa, 0, False)
        If .BumptyPenguin <> 0 Then GFX.DestroySurface .BumptyPenguin
        .BumptyPenguin = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.BumptyPenguin, 0, False)
        If .Spiney <> 0 Then GFX.DestroySurface .Spiney
        .Spiney = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Spiney, 0, False)
        If .LittleBoo <> 0 Then GFX.DestroySurface .LittleBoo
        .LittleBoo = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.LittleBoo, 0, False)
        If .RotoDisc <> 0 Then GFX.DestroySurface .RotoDisc
        .RotoDisc = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.RotoDisc, 0, False)
        If .LavaBubble <> 0 Then GFX.DestroySurface .LavaBubble
        .LavaBubble = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.LavaBubble, 0, False)
        If .Thwomp <> 0 Then GFX.DestroySurface .Thwomp
        .Thwomp = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Thwomp, 0, False)
        If .DryBones <> 0 Then GFX.DestroySurface .DryBones
        .DryBones = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.DryBones, 0, False)
        If .MovingPlatform <> 0 Then GFX.DestroySurface .MovingPlatform
        .MovingPlatform = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.MovingPlatform, 0, False)
        If .PowButton <> 0 Then GFX.DestroySurface .PowButton
        .PowButton = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.PowButton, 0, False)
        If .FreeCheepCheep <> 0 Then GFX.DestroySurface .FreeCheepCheep
        .FreeCheepCheep = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.FreeCheepCheep, 0, False)
        If .BlockCheepCheep <> 0 Then GFX.DestroySurface .BlockCheepCheep
        .BlockCheepCheep = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.BlockCheepCheep, 0, False)
        If .RisingLava <> 0 Then GFX.DestroySurface .RisingLava
        .RisingLava = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Layer3Lava, 0, False)
        If .MisterPokey <> 0 Then GFX.DestroySurface .MisterPokey
        .MisterPokey = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.TallCactus, 0, False)
        If .BouncyKoopa <> 0 Then GFX.DestroySurface .BouncyKoopa
        .BouncyKoopa = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Bouncer, 0, False)
        If .SurfBulletBill <> 0 Then GFX.DestroySurface .SurfBulletBill
        .SurfBulletBill = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.BulletBill, 0, False)
        If .Bobomb <> 0 Then GFX.DestroySurface .Bobomb
        .Bobomb = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Bobomb, 0, False)
        If .BOSSGoomboss <> 0 Then GFX.DestroySurface .BOSSGoomboss
        .BOSSGoomboss = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.BOSS_Goomboss, 0, False)
        If .Wiggler <> 0 Then GFX.DestroySurface .Wiggler
        .Wiggler = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.Wiggler, 0, False)
        If .SavePointStar <> 0 Then GFX.DestroySurface .SavePointStar
        .SavePointStar = GFX.CreateSurface(App.Path & "\Data\Enemies\" & oEnemySkin.SavePoint, 0, False)
    End With
End Sub

Public Sub LoadWorldList()
__fileLoader.LoadWorldList sWorldList
'Dim fFile As Long
'fFile = FreeFile
'Dim i As Long
'Dim sTemp As String
'    i = 0
'    ReDim sWorldList(0)
'    Open App.Path & "\Mods\" & sWorldSetName & "\Worlds.inf" For Input As fFile
'    Do
'        Line Input #fFile, sTemp
'        sTemp = Trim$(sTemp)
'        If sTemp = "." Then Exit Do
'        ReDim Preserve sWorldList(i)
'        sWorldList(i) = sTemp
'        i = i + 1
'    Loop While True
'    Close fFile
'    ReDim oWorldPassData.bWorldPassed(UBound(sWorldList))
'    For i = 0 To UBound(oWorldPassData.bWorldPassed)
'        ReDim oWorldPassData.bWorldPassed(i).bLevelPassed(0)
'    Next i
End Sub


Public Function fileExist(sPath As String) As Boolean
On Error GoTo errNoFile
Dim fFile As Long
fFile = FreeFile
fileExist = False
Open sPath For Input As fFile
fileExist = True
errNoFile:
Close fFile
End Function




Public Sub DrawBonusMeter()
Dim i As Long
    For i = 1 To 6
    If i <= gameGreens Then GFX.DrawSurface surfList.BarIcons, 0, 0, 16, 16, 272 + (16 * (i - 1)), 0 Else GFX.DrawSurface surfList.BarIcons, 0, 16, 16, 16, 272 + (16 * (i - 1)), 0
    Next i
    GFX.DrawSurface surfList.BarIcons, 16, 0, 48, 48, 296, 16 'Reserve Box
    Select Case marioReserveItem
        Case udeMarioReserveItem.rsvMushroom
            GFX.DrawSurface surfList.Sprites, 0, 0, 32, 32, 304, 24
        Case udeMarioReserveItem.rsvFlower
            GFX.DrawSurface surfList.Sprites, 32, 0, 32, 32, 304, 24
        Case udeMarioReserveItem.rsvHammer
            GFX.DrawSurface surfList.Sprites, 64, 32, 32, 32, 304, 24
        Case udeMarioReserveItem.rsvCarrot
            GFX.DrawSurface surfList.Sprites, 32, 32, 32, 32, 304, 24
        Case udeMarioReserveItem.rsvStar
            GFX.DrawSurface surfList.Sprites, 64, 0, 32, 32, 304, 24
    End Select
    GFX.DrawSurface surfList.BarIcons, 64, 0, 32, 32, 8, 8 'Mario Face
    GFX.DrawSurface surfList.BarIcons, 0, 32, 16, 16, 44, 16 'Lives X
    DrawBitmapNumber 64, 8, CStr(gameLives)
    GFX.DrawSurface surfList.BarIcons, 96, 0, 32, 32, 600, 8 'Mario Face
    GFX.DrawSurface surfList.BarIcons, 0, 32, 16, 16, 580, 16 'Lives X
    DrawBitmapNumber IIf(Len(CStr(gameCoins)) = 1, 560, 544), 8, CStr(gameCoins)
End Sub



Public Sub DrawBitmapNumber(xPos As Single, yPos As Single, ByVal sNumber As String)
Dim i As Long

    If Strings.Left$(sNumber, 1) = "-" Then sNumber = Strings.Right$(sNumber, Len(sNumber) - 1)
    For i = 1 To Len(sNumber)
        GFX.DrawSurface surfList.NumberFont, CLng(Mid$(sNumber, i, 1)) * 16, 0, 16, 32, xPos + (16 * (i - 1)), yPos
    Next i
    
End Sub


Public Function findDist(ax As Single, ay As Single, bx As Single, by As Single) As Single
Dim aa As Single
Dim bb As Single
    aa = (ax - bx)
    bb = (ay - by)
    findDist = Sqr((aa * aa) + (bb * bb))
End Function


Public Sub LoadSavedGame()
oSaveGameData = __fileLoader.LoadSavedGame()
'
'Dim fFile As Long
'Dim tempSaveGame As udtSaveGameData
'If Not fileExist(App.Path & "\Mods\" & sWorldSetName & "\SavedGames.3ks") Then
'    For fFile = 1 To 5
'        oSaveGameData.gameSlot(fFile).sSaveString = "[EMPTY]"
'    Next fFile
'Exit Sub
'End If
'    fFile = FreeFile
'    Open App.Path & "\Mods\" & sWorldSetName & "\SavedGames.3ks" For Binary Access Read Lock Write As fFile
'    Get fFile, 1, oSaveGameData
'    Close fFile
End Sub

Public Sub SaveGameData()
Dim fFile As Long
fFile = FreeFile
Open App.Path & "\Mods\" & sWorldSetName & "\SavedGames.3ks" For Binary Access Write Lock Read Write As fFile
Put fFile, 1, oSaveGameData
Close fFile
End Sub


Public Sub DetermineLevelPassedCount()
Dim i As Long
Dim u As Long
    lLevelPassedCount = 0
    For i = 0 To UBound(oWorldPassData.bWorldPassed())
        For u = 0 To UBound(oWorldPassData.bWorldPassed(i).bLevelPassed)
            If oWorldPassData.bWorldPassed(i).bLevelPassed(u) Then lLevelPassedCount = lLevelPassedCount + 1
        Next u
    Next i
End Sub



Async Public Sub ShowBlackTextScreen(sText As String)
Dim i As Long
    GFX.SetFont "Courier New", 16, True, False, False, False
    For i = 100 To 0 Step -4
    GFX.BeginScene 25
        GFX.DrawRect 0, 0, 640, 480, 0
        GFX.DrawText sText, 75, 75, vbWhite
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, 0, i, 0, 0, 0
    GFX.EndScene
    Next i
    For i = 0 To 50
    GFX.BeginScene 25
        GFX.DrawRect 0, 0, 640, 480, 0
        GFX.DrawText sText, 75, 75, vbWhite
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, 0, 0, 0, 0, 0
    GFX.EndScene
    Next i
    For i = 0 To 100 Step 4
    GFX.BeginScene 25
        GFX.DrawRect 0, 0, 640, 480, 0
        GFX.DrawText sText, 75, 75, vbWhite
        GFX.DrawSurface surfList.FadeMask, 0, 0, 16, 16, 0, 0, 640, 480, 0, i, 0, 0, 0
    GFX.EndScene
    Next i
End Sub

