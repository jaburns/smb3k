VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecDryBones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const PLAYDEADTIME As Long = 125 'frames
Private Const SHAKETIME As Long = 15 'frames


Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single

Private lFrame As Long
Private bFrameOn As Boolean

Private bAlive As Boolean
Private bDying As Boolean
Private lBreaking As Long
Private lBrokenTime As Long

Private bWatchForEdges As Boolean


Public Sub CreateAt(X As Long, Y As Long, Optional ByVal bSmartFall As Boolean = False)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    xSpeed = -1
    ySpeed = 0
    bFrameOn = False
    lFrame = 0
    bAlive = True
    bDying = False
    lBreaking = 0
    lBrokenTime = 0
    bWatchForEdges = bSmartFall
End Sub



Public Sub HandleMe()
Dim i As Long

    
    If Not bAlive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CSng(xInit), CSng(yInit), bWatchForEdges
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
        GFX.DrawSurface surfList.EnemyList.DryBones, 128, 0, 64, 32, xPos - screenLeft - 32, yPos - screenTop - 32
        Exit Sub
    End If
    
    If lBreaking Then
        If lBreaking < 5 Then
            GFX.DrawSurface surfList.EnemyList.DryBones, 128, 0, 64, 32, xPos - screenLeft - 32, yPos - screenTop - 32
        ElseIf lBreaking < PLAYDEADTIME Then
            GFX.DrawSurface surfList.EnemyList.DryBones, 128, 32, 64, 32, xPos - screenLeft - 32, yPos - screenTop - 32
        ElseIf lBreaking < PLAYDEADTIME + SHAKETIME Then
            GFX.DrawSurface surfList.EnemyList.DryBones, 128, 32, 64, 32, xPos - screenLeft - 32 + IIf(lBreaking Mod 2, 1, -1), yPos - screenTop - 32
        ElseIf lBreaking < PLAYDEADTIME + SHAKETIME + 4 Then
            GFX.DrawSurface surfList.EnemyList.DryBones, 128, 0, 64, 32, xPos - screenLeft - 32, yPos - screenTop - 32
        Else
            GFX.DrawSurface surfList.EnemyList.DryBones, 128, 0, 64, 32, xPos - screenLeft - 32, yPos - screenTop - 32
            lBreaking = 0
        End If
        If lBreaking > 0 Then lBreaking = lBreaking + 1
        Exit Sub
    End If
    
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
                PlaySound Sounds.BumpOff
                lBreaking = 1
            End If
        ElseIf .yPos >= yPos - 30 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
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
            End If
        End If
    Next i
    
    For i = 1 To GetShellCount()
        If GetShellActive(i) Then
            If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos) < 20 Then
                KillMe
                DestroyShellIfCarrying (i)
            End If
        End If
    Next i

    GFX.DrawSurface surfList.EnemyList.DryBones, IIf(xSpeed > 0, 64, 0) + IIf(bFrameOn, 32, 0), 0, 32, 64, xPos - screenLeft - 16, yPos - screenTop - 64
    
End Sub



Private Sub KillMe(Optional ByVal bIsStomped As Boolean = False)
    PlaySound Sounds.BumpOff
    bDying = True
    ySpeed = -4
End Sub



Private Sub MoveMe()
Dim tTile As udeLTileType

    xPos = xPos + xSpeed
    
    If GetLevelSideWarp() Then
        If xPos <= -1 Then xPos = GetLevelWidth()
        If xPos >= GetLevelWidth() + 1 Then xPos = 0
    Else
        If xPos <= 16 Then 
            xPos = 16
            xSpeed = 1
        End If
        If xPos >= GetLevelWidth() - 16 Then 
            xPos = GetLevelWidth() - 16
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
    
Dim sTestX As Single
Dim sTestY As Single
Dim i As Long
    For i = 1 To GetShellCount()
        If GetShellisShelled(i) Then
            sTestX = GetShellX(i) - xPos
            sTestY = Abs(GetShellY(i) - yPos)
            If sTestY <= 32 Then
                If sTestX > 0 And sTestX <= 32 Then xSpeed = -1
                If sTestX < 0 And sTestX >= -32 Then xSpeed = 1
            End If
        End If
    Next i
    
    ySpeed = ySpeed + 0.3
    If ySpeed > 5 Then ySpeed = 5
    yPos = yPos + ySpeed
    
    If isTileSolid(CLng(xPos), CLng(yPos), True) Then
        yPos = (__intDiv(yPos , 32) * 32)
        ySpeed = 0
        If GetTileAtPoint(CLng(xPos), CLng(yPos)) = DEADLY Then KillMe
    End If
    
    If bWatchForEdges Then
        If ySpeed = 0 And xSpeed <> 0 Then
            If Not isTileSolid(CLng(xPos) + IIf(xSpeed > 0, 16, -16), CLng(yPos) + 2, True) Then
                If xSpeed > 0 Then xSpeed = -1 Else xSpeed = 1
            End If
        End If
    End If
    
End Sub









