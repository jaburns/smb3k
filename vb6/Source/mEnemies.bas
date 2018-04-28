Attribute VB_Name = "mEnemies"
Option Explicit

Public Enum udeKoopaStyle
    KOOPA_GREEN
    KOOPA_RED
    KOOPA_BEETLE
    KOOPA_BOUNCER
End Enum

Private Type objHiddenTileInfo
    xPos As Long
    yPos As Long
    bHit As Boolean
End Type

'all the enemies
Private zoGoomba() As ecGoomba
Private zoPPlant() As ecPiranaPlant
Private zoLavaBall() As ecLavaBall
Private zoGhostDisc() As ecGhostDisc
Private zoGhost() As ecGhost
Private zoSpiney() As ecSpiney
Private zoPlatform() As ecPlatform
Private zoKoopa() As ecKoopa
Private zoPenguin() As ecPenguin
Private zoDonut() As ecDonut
Private zoManyCoinBlock() As ecManyCoinBlock
Private zoDryBones() As ecDryBones
Private zoThwomp() As ecThwomp
Private zoPowBlock() As ecPowBlock
Private zoTallCactus() As ecCactus
Private zoBulletBill() As ecBulletBill
Private zoCheepCheep() As ecCheepCheep
Private zoBobomb() As ecBobomb
Private zoBossGoomboss As eBossGoomboss
Private bGoombossOn As Boolean
Private zoHiddenTile() As objHiddenTileInfo
Private zoWiggler() As ecWiggler


'this also finds the tags and other special things
Public Sub FindAndCreateEnemies()
Dim TMP As Long
Dim X As Long
Dim Y As Long
Dim tType As udeLTileType
Dim tEnemy As udeLTileEnemy
Dim tTag As udeLTileTag

    ReDim zoGoomba(0)
    ReDim zoPPlant(0)
    ReDim zoLavaBall(0)
    ReDim zoGhostDisc(0)
    ReDim zoGhost(0)
    ReDim zoSpiney(0)
    ReDim zoPlatform(0)
    ReDim zoKoopa(0)
    ReDim zoPenguin(0)
    ReDim zoDonut(0)
    ReDim zoDryBones(0)
    ReDim zoThwomp(0)
    ReDim zoPowBlock(0)
    ReDim zoTallCactus(0)
    ReDim zoBulletBill(0)
    ReDim zoCheepCheep(0)
    ReDim zoBobomb(0)
    ReDim zoManyCoinBlock(0)
    ReDim zoHiddenTile(0)
    ReDim zoWiggler(0)
    
    bGoombossOn = False
    
    Mario.SetAlwaysSwimming False, 0
    For X = 0 To __intDiv(GetLevelWidth , 32)
        For Y = 0 To __intDiv(GetLevelHeight , 32)
            tEnemy = GetTileEnemy(X, Y)
            tType = GetTile(X, Y)
            tTag = GetTileTagAtPoint(X * 32, Y * 32)
            PlatformTest tEnemy, tTag, X, Y, TMP
            If tType = DONUTTILE Then
                TMP = UBound(zoDonut) + 1
                ReDim Preserve zoDonut(TMP)
                Set zoDonut(TMP) = New ecDonut
                zoDonut(TMP).CreateAt (X * 32), (Y * 32), GetTileXSrcPoint(X * 32, Y * 32), GetTileYSrcPoint(X * 32, Y * 32)
                KillTile X * 32, Y * 32
            End If
            If tTag = MANYCOINSTAG Then
                TMP = UBound(zoManyCoinBlock) + 1
                ReDim Preserve zoManyCoinBlock(TMP)
                Set zoManyCoinBlock(TMP) = New ecManyCoinBlock
                zoManyCoinBlock(TMP).CreateAt (X * 32), (Y * 32), GetTileXSrcPoint(X * 32, Y * 32), GetTileYSrcPoint(X * 32, Y * 32), GetTileAnim(X, Y)
                KillTile X * 32, Y * 32
            End If
            If tTag = ALLWATERTAG Then Mario.SetAlwaysSwimming True, Y * 32
            Select Case tEnemy
                Case OBJGOOMBA
                    TMP = UBound(zoGoomba) + 1
                    ReDim Preserve zoGoomba(TMP)
                    Set zoGoomba(TMP) = New ecGoomba
                    zoGoomba(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, , , (tTag = ENEMYSMARTTAG)
                Case OBJPPLANT
                    TMP = UBound(zoPPlant) + 1
                    ReDim Preserve zoPPlant(TMP)
                    Set zoPPlant(TMP) = New ecPiranaPlant
                    zoPPlant(TMP).CreateAt (X * 32) + 32, (Y * 32) + 32, tTag = PPLANTMEAN, False
                Case OBJFIREBALL
                    TMP = UBound(zoLavaBall) + 1
                    ReDim Preserve zoLavaBall(TMP)
                    Set zoLavaBall(TMP) = New ecLavaBall
                    zoLavaBall(TMP).CreateAt (X * 32) + 16, (Y * 32) + 64
                Case OBJROTODISC
                    TMP = UBound(zoGhostDisc) + 1
                    ReDim Preserve zoGhostDisc(TMP)
                    Set zoGhostDisc(TMP) = New ecGhostDisc
                    zoGhostDisc(TMP).CreateAt (X * 32) + 16, (Y * 32) + 16
                Case OBJLITTLEBOO
                    TMP = UBound(zoGhost) + 1
                    ReDim Preserve zoGhost(TMP)
                    Set zoGhost(TMP) = New ecGhost
                    zoGhost(TMP).CreateAt (X * 32) + 16, (Y * 32) + 16
                Case OBJSPINEY
                    TMP = UBound(zoSpiney) + 1
                    ReDim Preserve zoSpiney(TMP)
                    Set zoSpiney(TMP) = New ecSpiney
                    zoSpiney(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, (tTag = ENEMYSMARTTAG)
                Case OBJKOOPA
                    TMP = UBound(zoKoopa) + 1
                    ReDim Preserve zoKoopa(TMP)
                    Set zoKoopa(TMP) = New ecKoopa
                    zoKoopa(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, KOOPA_GREEN, tTag = KOOPASHELL, tTag = WINGEDKOOPA, (tTag = ENEMYSMARTTAG)
                Case OBJSMARTKOOPA
                    TMP = UBound(zoKoopa) + 1
                    ReDim Preserve zoKoopa(TMP)
                    Set zoKoopa(TMP) = New ecKoopa
                    zoKoopa(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, KOOPA_RED, tTag = KOOPASHELL, tTag = WINGEDKOOPA, (tTag = ENEMYSMARTTAG)
                Case OBJBEETLE
                    TMP = UBound(zoKoopa) + 1
                    ReDim Preserve zoKoopa(TMP)
                    Set zoKoopa(TMP) = New ecKoopa
                    zoKoopa(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, KOOPA_BEETLE, tTag = KOOPASHELL, tTag = WINGEDKOOPA, (tTag = ENEMYSMARTTAG)
                Case OBJKOOPABOUNCER
                    TMP = UBound(zoKoopa) + 1
                    ReDim Preserve zoKoopa(TMP)
                    Set zoKoopa(TMP) = New ecKoopa
                    zoKoopa(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, KOOPA_BOUNCER, tTag = KOOPASHELL, tTag = WINGEDKOOPA, (tTag = ENEMYSMARTTAG)
                Case OBJBUMPTY
                    TMP = UBound(zoPenguin) + 1
                    ReDim Preserve zoPenguin(TMP)
                    Set zoPenguin(TMP) = New ecPenguin
                    zoPenguin(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, (tTag = ENEMYSMARTTAG)
                Case OBJDRYBONES
                    TMP = UBound(zoDryBones) + 1
                    ReDim Preserve zoDryBones(TMP)
                    Set zoDryBones(TMP) = New ecDryBones
                    zoDryBones(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, (tTag = ENEMYSMARTTAG)
                Case OBJTHWOMP
                    TMP = UBound(zoThwomp) + 1
                    ReDim Preserve zoThwomp(TMP)
                    Set zoThwomp(TMP) = New ecThwomp
                    zoThwomp(TMP).CreateAt (X * 32) + 32, (Y * 32) + 64
                Case OBJPOWBUTTON
                    TMP = UBound(zoPowBlock) + 1
                    ReDim Preserve zoPowBlock(TMP)
                    Set zoPowBlock(TMP) = New ecPowBlock
                    zoPowBlock(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32
                Case OBJRISINGLAVA
                    layer3enabled = True
                Case OBJSINUSOIDLAVA
                    layer3enabled = True
                    layer3sinusoid = True
                    layer3maximum = GetLevelHeight - (Y * 32) - 16
                Case OBJTALLCACTUS
                    TMP = UBound(zoTallCactus) + 1
                    ReDim Preserve zoTallCactus(TMP)
                    Set zoTallCactus(TMP) = New ecCactus
                    zoTallCactus(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, (tTag = ENEMYSMARTTAG)
                Case OBJBULLETBILL
                    TMP = UBound(zoBulletBill) + 1
                    ReDim Preserve zoBulletBill(TMP)
                    Set zoBulletBill(TMP) = New ecBulletBill
                    zoBulletBill(TMP).CreateAt (X * 32) + 16, (Y * 32) + 16, tTag = BULLETSEEKER
                Case OBJTHRUFISH
                    TMP = UBound(zoCheepCheep) + 1
                    ReDim Preserve zoCheepCheep(TMP)
                    Set zoCheepCheep(TMP) = New ecCheepCheep
                    zoCheepCheep(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, False
                Case OBJBLOCKFISH
                    TMP = UBound(zoCheepCheep) + 1
                    ReDim Preserve zoCheepCheep(TMP)
                    Set zoCheepCheep(TMP) = New ecCheepCheep
                    zoCheepCheep(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, True
                Case OBJBOBOMB
                    TMP = UBound(zoBobomb) + 1
                    ReDim Preserve zoBobomb(TMP)
                    Set zoBobomb(TMP) = New ecBobomb
                    zoBobomb(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, (tTag = ENEMYSMARTTAG)
                Case OBJUPSIDEDOWNPLANT
                    TMP = UBound(zoPPlant) + 1
                    ReDim Preserve zoPPlant(TMP)
                    Set zoPPlant(TMP) = New ecPiranaPlant
                    zoPPlant(TMP).CreateAt (X * 32) + 32, (Y * 32) + 32, tTag = PPLANTMEAN, True
                Case OBJBOSSGOOMBOSS
                    If tType = Background Then
                        Set zoBossGoomboss = New eBossGoomboss
                        zoBossGoomboss.CreateAt (X * 32), (Y * 32)
                        bGoombossOn = True
                    Else
                        If tTag = BOSSTILEHIDE Then
                            If (blBossesPassed And BP_1) Then KillTile X * 32, Y * 32
                        ElseIf (blBossesPassed And BP_1) = 0 Then
                            KillTile X * 32, Y * 32
                        End If
                    End If
                Case OBJHIDETILE
                    TMP = UBound(zoHiddenTile) + 1
                    ReDim Preserve zoHiddenTile(TMP)
                    zoHiddenTile(TMP).bHit = False
                    zoHiddenTile(TMP).xPos = X
                    zoHiddenTile(TMP).yPos = Y
                Case OBJWIGGLER
                    TMP = UBound(zoWiggler) + 1
                    ReDim Preserve zoWiggler(TMP)
                    Set zoWiggler(TMP) = New ecWiggler
                    zoWiggler(TMP).CreateAt (X * 32) + 16, (Y * 32) + 32, (tTag = ENEMYSMARTTAG)
            End Select
        Next Y
    Next X

End Sub


Public Sub MakeGoomba(lX As Long, lY As Long, sInitXSpeed As Single, sInitYSpeed As Single)
    ReDim Preserve zoGoomba(UBound(zoGoomba) + 1)
    Set zoGoomba(UBound(zoGoomba)) = New ecGoomba
    zoGoomba(UBound(zoGoomba)).CreateAt lX, lY, sInitXSpeed, sInitYSpeed
End Sub


Public Sub HandleEnemies()
Dim i As Long
    If bGoombossOn Then
        zoBossGoomboss.HandleMe
    End If
    For i = 1 To UBound(zoDonut)
    zoDonut(i).HandleMe
    Next i
    For i = 1 To UBound(zoWiggler)
    zoWiggler(i).HandleMe
    Next i
    For i = 1 To UBound(zoPenguin)
    zoPenguin(i).HandleMe
    Next i
    For i = 1 To UBound(zoGoomba)
    zoGoomba(i).HandleMe
    Next i
    For i = 1 To UBound(zoPPlant)
    zoPPlant(i).HandleMe
    Next i
    For i = 1 To UBound(zoLavaBall)
    zoLavaBall(i).HandleMe
    Next i
    For i = 1 To UBound(zoGhostDisc)
    zoGhostDisc(i).HandleMe
    Next i
    For i = 1 To UBound(zoGhost)
    zoGhost(i).HandleMe
    Next i
    For i = 1 To UBound(zoSpiney)
    zoSpiney(i).HandleMe
    Next i
    For i = 1 To UBound(zoPlatform)
    zoPlatform(i).HandleMe
    Next i
    For i = 1 To UBound(zoKoopa)
    zoKoopa(i).HandleMe
    Next i
    For i = 1 To UBound(zoDryBones)
    zoDryBones(i).HandleMe
    Next i
    For i = 1 To UBound(zoThwomp)
    zoThwomp(i).HandleMe
    Next i
    For i = 1 To UBound(zoPowBlock)
    zoPowBlock(i).HandleMe
    Next i
    For i = 1 To UBound(zoTallCactus)
    zoTallCactus(i).HandleMe
    Next i
    For i = 1 To UBound(zoBulletBill)
    zoBulletBill(i).HandleMe
    Next i
    For i = 1 To UBound(zoCheepCheep)
    zoCheepCheep(i).HandleMe
    Next i
    For i = 1 To UBound(zoBobomb)
    zoBobomb(i).HandleMe
    Next i
    For i = 1 To UBound(zoManyCoinBlock)
    zoManyCoinBlock(i).HandleMe
    Next i
End Sub


Public Sub KillEnemies()
Dim i As Long
    For i = 0 To UBound(zoGoomba)
    Set zoGoomba(i) = Nothing
    Next i
    For i = 0 To UBound(zoPPlant)
    Set zoPPlant(i) = Nothing
    Next i
    For i = 0 To UBound(zoLavaBall)
    Set zoLavaBall(i) = Nothing
    Next i
    For i = 0 To UBound(zoGhostDisc)
    Set zoGhostDisc(i) = Nothing
    Next i
    For i = 0 To UBound(zoGhost)
    Set zoGhost(i) = Nothing
    Next i
    For i = 0 To UBound(zoSpiney)
    Set zoSpiney(i) = Nothing
    Next i
    For i = 0 To UBound(zoPlatform)
    Set zoPlatform(i) = Nothing
    Next i
    For i = 0 To UBound(zoKoopa)
    Set zoKoopa(i) = Nothing
    Next i
    For i = 0 To UBound(zoPenguin)
    Set zoPenguin(i) = Nothing
    Next i
    For i = 0 To UBound(zoDonut)
    Set zoDonut(i) = Nothing
    Next i
    For i = 0 To UBound(zoManyCoinBlock)
    Set zoManyCoinBlock(i) = Nothing
    Next i
    For i = 0 To UBound(zoDryBones)
    Set zoDryBones(i) = Nothing
    Next i
    For i = 0 To UBound(zoThwomp)
    Set zoThwomp(i) = Nothing
    Next i
    For i = 0 To UBound(zoPowBlock)
    Set zoPowBlock(i) = Nothing
    Next i
    For i = 0 To UBound(zoTallCactus)
    Set zoTallCactus(i) = Nothing
    Next i
    For i = 0 To UBound(zoBulletBill)
    Set zoBulletBill(i) = Nothing
    Next i
    For i = 0 To UBound(zoCheepCheep)
    Set zoCheepCheep(i) = Nothing
    Next i
    For i = 1 To UBound(zoBobomb)
    Set zoBobomb(i) = Nothing
    Next i
    For i = 1 To UBound(zoWiggler)
    Set zoWiggler(i) = Nothing
    Next i
    If bGoombossOn Then
        Set zoBossGoomboss = Nothing
    End If

End Sub



Public Sub PlatformTest(zTileEnemy As udeLTileEnemy, zTileTag As udeLTileTag, X As Long, Y As Long, TMP As Long)
Dim sSize As Single

    If zTileEnemy = OBJPLATFORM2 Then
        sSize = 0
    ElseIf zTileEnemy = OBJPLATFORM3 Then
        sSize = 1
    ElseIf zTileEnemy = OBJPLATFORM4 Then
        sSize = 2
    ElseIf zTileEnemy = OBJPLATFORM5 Then
        sSize = 3
    Else
        Exit Sub
    End If
       
    TMP = UBound(zoPlatform) + 1
    ReDim Preserve zoPlatform(TMP)
    Set zoPlatform(TMP) = New ecPlatform
    If zTileTag = PLATFORMSLOW Then
        zoPlatform(TMP).CreateAt (X * 32), (Y * 32), 1, sSize
    ElseIf zTileTag = PLATFORMFAST Then
        zoPlatform(TMP).CreateAt (X * 32), (Y * 32), 4, sSize
    Else
        zoPlatform(TMP).CreateAt (X * 32), (Y * 32), 2, sSize
    End If
    
End Sub




'make shell data public
Public Function GetShellX(lID As Long) As Single
If lID < 1 Or lID > UBound(zoKoopa) Then Exit Function
GetShellX = zoKoopa(lID).xLoc
End Function
Public Function GetShellY(lID As Long) As Single
If lID < 1 Or lID > UBound(zoKoopa) Then Exit Function
GetShellY = zoKoopa(lID).yLoc
End Function
Public Function GetShellActive(lID As Long) As Boolean
If lID < 1 Or lID > UBound(zoKoopa) Then Exit Function
GetShellActive = zoKoopa(lID).isShreddin
End Function
Public Sub DestroyShellIfCarrying(lID As Long)
If lID < 1 Or lID > UBound(zoKoopa) Then Exit Sub
zoKoopa(lID).killShellIfCarrying
End Sub
Public Function GetShellCount() As Long
GetShellCount = UBound(zoKoopa)
End Function
Public Function GetShellisShelled(lID As Long) As Boolean
If lID < 1 Or lID > UBound(zoKoopa) Then Exit Function
GetShellisShelled = zoKoopa(lID).isShelled
End Function



'tests if the point is inside ANY bobomb blast
Public Function inBombBlast(sX As Single, sY As Single) As Boolean
Dim i As Long
    inBombBlast = False
    For i = 1 To UBound(zoBobomb)
        If zoBobomb(i).isInBlastRadius(sX, sY) Then
            inBombBlast = True
            Exit For
        End If
    Next i
End Function


'tests to see of the point is in an active falling donut
Public Function inActiveDonut(sX As Single, sY As Single) As Boolean
Dim i As Long
    inActiveDonut = False
    For i = 1 To UBound(zoDonut)
        If zoDonut(i).activeAtLocation(sX, sY) Then
            inActiveDonut = True
            Exit For
        End If
    Next i
End Function


'unhide any hidden block which may lie under this point and return true if that happened
Public Function hitHiddenBlock(lX As Long, lY As Long) As Boolean
Dim i As Long
    hitHiddenBlock = False
    For i = 1 To UBound(zoHiddenTile)
        If Not zoHiddenTile(i).bHit Then
            If __intDiv(lX , 32) = zoHiddenTile(i).xPos And __intDiv(lY , 32) = zoHiddenTile(i).yPos Then
                SetTileEnemy lX, lY, NO_ENEMY
                zoHiddenTile(i).bHit = True
                hitHiddenBlock = True
                Exit For
            End If
        End If
    Next i
End Function


'tests to see of the point is in a many coin block
Public Function inManyCoinBlock(sX As Single, sY As Single) As Boolean
Dim i As Long
    inManyCoinBlock = False
    For i = 1 To UBound(zoManyCoinBlock)
        If zoManyCoinBlock(i).activeAtLocation(sX, sY) Then
            inManyCoinBlock = True
            Exit For
        End If
    Next i
End Function
'allow mario to hit the many coin block
Public Sub hitManyCoinBlock(lX As Long, lY As Long)
Dim i As Long
    For i = 1 To UBound(zoManyCoinBlock)
        If zoManyCoinBlock(i).hitMe(CSng(lX), CSng(lY)) Then Exit For
    Next i
End Sub
