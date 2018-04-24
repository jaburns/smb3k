VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecCheepCheep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'blocked means slow in this shit

Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single

Private bBlocked As Boolean

Private lFrame As Long
Private bFrameOn As Boolean

Private bAlive As Boolean
Private bDying As Boolean


Public Sub CreateAt(X As Long, Y As Long, bIsBlocked As Boolean)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    xSpeed = -1
    bFrameOn = False
    lFrame = 0
    bAlive = True
    bDying = False
    bBlocked = bIsBlocked
End Sub



Public Sub HandleMe()
Dim i As Long

    If Not bAlive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CSng(xInit), CSng(yInit), bBlocked
        Exit Sub
    End If
    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
        If xPos <> xInit And yPos <> yInit Then bAlive = False
        Exit Sub
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
        GFX.DrawSurface IIf(bBlocked, surfList.EnemyList.BlockCheepCheep, surfList.EnemyList.FreeCheepCheep), IIf(xSpeed > 0, 128, 160), 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        Exit Sub
    End If
    
    lFrame = lFrame + 1
    If lFrame >= IIf(bBlocked, 5, 3) Then
        lFrame = 0
        bFrameOn = Not bFrameOn
    End If
    
    MoveMe
    
    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 32 And .yPos < yPos - 20 Then
            If Not .isSwimming Then
                If .bInvinsible Then
                    KillMe
                Else
                    .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
                    KillMe
                End If
            Else
                If .bInvinsible Then KillMe Else .Injure
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
                KillMe
            End If
        End If
    Next i
    
    For i = 1 To GetShellCount
        If GetShellActive(i) Then
            If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos) < 32 Then
                KillMe
                DestroyShellIfCarrying (i)
            End If
        End If
    Next i
    If inBombBlast(xPos, yPos - 16) Then KillMe

    GFX.DrawSurface IIf(bBlocked, surfList.EnemyList.BlockCheepCheep, surfList.EnemyList.FreeCheepCheep), IIf(xSpeed > 0, 0, 64) + IIf(bFrameOn, 32, 0), 0, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    
End Sub



Private Sub KillMe()
    PlaySound Sounds.BumpOff
    bDying = True
    ySpeed = -2
End Sub



Private Sub MoveMe()
Dim tTile As udeLTileType

    xPos = xPos + IIf(bBlocked, xSpeed, 2 * xSpeed)
    
    If GetLevelSideWarp Then
        If xPos <= -1 Then xPos = GetLevelWidth
        If xPos >= GetLevelWidth + 1 Then xPos = 0
    Else
        If xPos <= 16 Then 
            xPos = 16
            xSpeed = 1
        End If
        If xPos >= GetLevelWidth - 16 Then 
            xPos = GetLevelWidth - 16
            xSpeed = -1
        End If
    End If
    
    If isTileSolid(xPos + 16, yPos - 16) Then
        If GetTileAtPoint(xPos + 16, yPos - 16) = DEADLY Then KillMe
        xSpeed = -1
    ElseIf isTileSolid(xPos - 16, yPos - 16) Then
        If GetTileAtPoint(xPos - 16, yPos - 16) = DEADLY Then KillMe
        xSpeed = 1
    End If
    
End Sub

