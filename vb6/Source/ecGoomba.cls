VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecGoomba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single
Private sActualSpeed As Single

Private lFrame As Long
Private bFrameOn As Boolean

Private bAlive As Boolean
Private bDying As Boolean
Private bStomped As Boolean
Private lStompTime As Long

Private bWatchForEdges As Boolean


Public Sub CreateAt(X As Long, Y As Long, Optional ByVal InitXSpeed As Single = -1, Optional ByVal InitYSpeed As Single = 0, Optional ByVal bSmartFall As Boolean = False)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    xSpeed = InitXSpeed
    ySpeed = InitYSpeed
    sActualSpeed = Abs(InitXSpeed)
    bFrameOn = False
    lFrame = 0
    bAlive = True
    bDying = False
    bStomped = False
    lStompTime = 0
    bWatchForEdges = bSmartFall
End Sub



Public Sub HandleMe()
Dim i As Long

    
    If Not bAlive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 _
        Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CSng(xInit), CSng(yInit), , , bWatchForEdges
        Exit Sub
    End If
    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 _
    Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
        If xPos <> xInit And yPos <> yInit Then bAlive = False
        Exit Sub
    End If

    If bDying Then
        If bStomped Then
            lStompTime = lStompTime + 1
            MoveMe
            If lStompTime > 30 Then
                bAlive = False
                bDying = False
                Exit Sub
            End If
            GFX.DrawSurface surfList.EnemyList.Goomba, 64, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
            Exit Sub
        Else
            ySpeed = ySpeed + 0.3
            If ySpeed > 7 Then ySpeed = 7
            yPos = yPos + ySpeed
            If yPos > screenTop + 512 Then
                bAlive = False
                bDying = False
                Exit Sub
            End If
            GFX.DrawSurface surfList.EnemyList.Goomba, 96, 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
            Exit Sub
        End If
    End If
    
    lFrame = lFrame + 1
    If lFrame >= 5 Then
        lFrame = 0
        bFrameOn = Not bFrameOn
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
            If getMarioHammerX(i) > xPos - 16 And getMarioHammerX(i) < xPos + 16 _
            And getMarioHammerY(i) > yPos - 32 And getMarioHammerY(i) < yPos Then KillMe
        End If
        If getMarioFireActive(i) Then
            If getMarioFireX(i) > xPos - 16 And getMarioFireX(i) < xPos + 16 _
            And getMarioFireY(i) > yPos - 32 And getMarioFireY(i) < yPos Then
                killMarioFire i
                KillMe
            End If
        End If
    Next i
    
    For i = 1 To GetShellCount
        If GetShellActive(i) Then
            If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos) < 20 Then
                KillMe
                DestroyShellIfCarrying (i)
            End If
        End If
    Next i
    If inBombBlast(xPos, yPos - 16) Then KillMe

    GFX.DrawSurface surfList.EnemyList.Goomba, IIf(bFrameOn, 32, 0), 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    
End Sub



Private Sub KillMe(Optional ByVal bIsStomped As Boolean = False)
    PlaySound Sounds.BumpOff
    bStomped = bIsStomped
    bDying = True
    ySpeed = IIf(bStomped, 0, -4)
End Sub



Private Sub MoveMe()
Dim tTile As udeLTileType

    If Not bDying Then xPos = xPos + xSpeed
    
    If GetLevelSideWarp Then
        If xPos <= -1 Then xPos = GetLevelWidth
        If xPos >= GetLevelWidth + 1 Then xPos = 0
    Else
        If xPos <= 16 Then xPos = 16: xSpeed = sActualSpeed
        If xPos >= GetLevelWidth - 16 Then xPos = GetLevelWidth - 16: xSpeed = -sActualSpeed
    End If
    
    If isTileSolid(xPos + 16, yPos - 16) Then
        If GetTileAtPoint(xPos + 16, yPos - 16) = DEADLY Then KillMe
        xSpeed = -sActualSpeed
    ElseIf isTileSolid(xPos - 16, yPos - 16) Then
        If GetTileAtPoint(xPos - 16, yPos - 16) = DEADLY Then KillMe
        xSpeed = sActualSpeed
    End If
    
Dim sTestX As Single
Dim sTestY As Single
Dim i As Long
    For i = 1 To GetShellCount
        If GetShellisShelled(i) Then
            sTestX = GetShellX(i) - xPos
            sTestY = Abs(GetShellY(i) - yPos)
            If sTestY <= 32 Then
                If sTestX > 0 And sTestX <= 32 Then xSpeed = -sActualSpeed
                If sTestX < 0 And sTestX >= -32 Then xSpeed = sActualSpeed
            End If
        End If
    Next i
    
Dim maxYSpeed As Long
    ySpeed = ySpeed + 0.3
    If GetTileAtPoint(CLng(xPos), CLng(yPos) - 8) = WATER Then maxYSpeed = 3 Else maxYSpeed = 5
    If ySpeed > maxYSpeed Then ySpeed = maxYSpeed
    yPos = yPos + ySpeed
    
    If isTileSolid(CLng(xPos), CLng(yPos), True) Then
        yPos = ((yPos \ 32) * 32)
        ySpeed = 0
        If GetTileAtPoint(CLng(xPos), CLng(yPos)) = DEADLY Then KillMe
    End If

    If bWatchForEdges Then
        If ySpeed = 0 And xSpeed <> 0 Then
            If Not isTileSolid(CLng(xPos) + IIf(xSpeed > 0, 16, -16), CLng(yPos) + 2, True) Then
                If xSpeed > 0 Then xSpeed = -sActualSpeed Else xSpeed = sActualSpeed
            End If
        End If
    End If
    
End Sub







