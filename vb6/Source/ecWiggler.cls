VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecWiggler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NORM_SPEED As Long = 5
Private Const PISSED_SPEED As Long = 2
Private Const TIME_SPENT_PISSED As Long = 500

Private xInit As Single
Private yInit As Single
Private xPos(4) As Single
Private yPos(4) As Single
Private xSpeed(4) As Single
Private ySpeed(4) As Single

Private lFrame(4) As Long
Private lFrameCount As Long

Private yDeathTrail(20) As Long
Private yDied(4) As Single

Private bAlive As Boolean
Private bDying As Boolean
Private lPissedCount As Long
Private curAnimSpeed As Long
Private curSpeed As Long

Private bWatchForEdges As Boolean


Public Sub CreateAt(X As Long, Y As Long, Optional ByVal bSmartFall As Boolean = False)
Dim i As Long
    xInit = CSng(X)
    yInit = CSng(Y)
    For i = 0 To 4
        xPos(i) = xInit + (20 * i)
        yPos(i) = yInit
        xSpeed(i) = -1
        ySpeed(i) = 0
    Next i
    curAnimSpeed = NORM_SPEED
    curSpeed = 1
    lPissedCount = 0
    lFrame(0) = 0
    lFrame(1) = 1
    lFrame(2) = 2
    lFrame(3) = 3
    lFrame(4) = 0
    lFrameCount = 0
    bAlive = True
    bDying = False
    bWatchForEdges = bSmartFall
End Sub



Public Sub HandleMe()
Dim i As Long
Dim u As Long

    If Not bAlive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CSng(xInit), CSng(yInit), bWatchForEdges
        Exit Sub
    End If
    
    If bDying Then
        ySpeed(0) = ySpeed(0) + 0.3
        If ySpeed(0) > 7 Then ySpeed(0) = 7
        yPos(0) = yPos(0) + ySpeed(0)
        UpdateDeathTrail
        If yPos(4) > screenTop + 512 Then
            bAlive = False
            bDying = False
        End If
        DrawWiggler
        Exit Sub
    End If

    If xPos(0) < screenLeft - 128 Or xPos(0) > screenLeft + 768 Or yPos(0) < screenTop - 128 Or yPos(0) > screenTop + 608 Then
        If xPos(0) <> xInit And yPos(0) <> yInit Then bAlive = False
        Exit Sub
    End If
        
    If lPissedCount > 0 Then
        lPissedCount = lPissedCount - 1
        If lPissedCount = 0 Then
            curAnimSpeed = NORM_SPEED
            SetCurSpeed 1
        End If
    End If
    
    lFrameCount = lFrameCount + 1
        If lFrameCount >= curAnimSpeed Then
        lFrameCount = 0
        For i = 0 To 4
            lFrame(i) = lFrame(i) + 1
            If lFrame(i) = 4 Then lFrame(i) = 0
        Next i
    End If

    
    For u = 0 To 4
        MoveMe u
        With Mario
            If .xPos >= xPos(u) - 28 And .xPos <= xPos(u) + 28 And .bAlive Then
                If .yPos >= yPos(u) - 32 And .yPos < yPos(u) + IIf(Not .isTall, 32, 52) Then
                    If .yPos < yPos(u) - 20 Then
                        If .bInvinsible Then
                            KillMe
                        Else
                            .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
                            KillMe True
                        End If
                    Else
                        If .bInvinsible Then KillMe Else .Injure
                    End If
                End If
            End If
        End With
        For i = 0 To 1
            If getMarioHammerActive(i) And u = 0 Then
                If getMarioHammerX(i) > xPos(u) - 16 And getMarioHammerX(i) < xPos(u) + 16 And getMarioHammerY(i) > yPos(u) - 32 And getMarioHammerY(i) < yPos(u) Then KillMe
            End If
            If getMarioFireActive(i) Then
                If getMarioFireX(i) > xPos(u) - 16 And getMarioFireX(i) < xPos(u) + 16 And getMarioFireY(i) > yPos(u) - 32 And getMarioFireY(i) < yPos(u) Then killMarioFire i
            End If
        Next i
        For i = 1 To GetShellCount()
            If GetShellActive(i) Then
                If Abs(GetShellX(i) - xPos(u)) < 32 And Abs(GetShellY(i) - yPos(u)) < 20 Then
                    KillMe
                    DestroyShellIfCarrying (i)
                End If
            End If
        Next i
    Next u
    
    DrawWiggler

End Sub



Private Function getXSrc(lBall As Long) As Long
    getXSrc = lFrame(lBall) * 32
    If getXSrc = 96 Then getXSrc = 32
    If lBall > 0 Then getXSrc = getXSrc + 96
End Function
Private Function getYSrc(lBall As Long) As Long
    getYSrc = 0
    If lBall = 0 Then
        If xSpeed(lBall) > 0 Then getYSrc = 32
    Else
        If xPos(lBall) < xPos(lBall - 1) Then getYSrc = 32
    End If
    If lPissedCount > 0 Then getYSrc = getYSrc + 64
End Function

Private Sub DrawWiggler()
Dim i As Long
    For i = 4 To 0 Step -1
        GFX.DrawSurface surfList.EnemyList.Wiggler, getXSrc(i), getYSrc(i), 32, 32, xPos(i) - screenLeft - 16, yPos(i) - screenTop - 32
    Next i
    If lPissedCount = 0 Then
        i = lFrame(0)
        If i = 3 Then i = 1
        i = 2 - i
        GFX.DrawSurface surfList.EnemyList.Wiggler, 192, 0, 16, 16, xPos(0) - screenLeft - 8, yPos(0) - screenTop - 44 - (i * 2)
    End If
End Sub



Private Sub SetCurSpeed(lNewXSpeed As Long)
Dim i As Long
    curSpeed = lNewXSpeed
    For i = 0 To 4
        xSpeed(i) = Sgn(xSpeed(i)) * lNewXSpeed
    Next i
End Sub



Private Sub KillMe(Optional ByVal bIsStomped As Boolean = False)
    PlaySound Sounds.BumpOff
    If bIsStomped Then
        lPissedCount = TIME_SPENT_PISSED
        SetCurSpeed 2
        curAnimSpeed = PISSED_SPEED
    Else
        bDying = True
        ySpeed(0) = -3
        StartDeathTrail
    End If
End Sub
Private Sub StartDeathTrail()
Dim i As Long
    For i = 0 To 4
        yDied(i) = yPos(i)
    Next i
    For i = 0 To 20
        yDeathTrail(i) = 0
    Next i
End Sub
Public Sub UpdateDeathTrail()
Dim i As Long
    For i = 20 To 1 Step -1
        yDeathTrail(i) = yDeathTrail(i - 1)
    Next i
    yDeathTrail(0) = yPos(0) - yDied(0)
    For i = 1 To 4
        yPos(i) = yDied(i) + yDeathTrail(i * 5)
    Next i
End Sub




Private Sub MoveMe(lBall As Long)
Dim tTile As udeLTileType

    If Not bDying Then xPos(lBall) = xPos(lBall) + xSpeed(lBall)
    
    If GetLevelSideWarp Then
        If xPos(lBall) <= -1 Then xPos(lBall) = GetLevelWidth()
        If xPos(lBall) >= GetLevelWidth() + 1 Then xPos(lBall) = 0
    Else
        If xPos(lBall) <= 16 Then 
            xPos(lBall) = 16
            xSpeed(lBall) = curSpeed
        End If
        If xPos(lBall) >= GetLevelWidth() - 16 Then 
            xPos(lBall) = GetLevelWidth() - 16
            xSpeed(lBall) = -curSpeed
        End If
    End If
    
    If isTileSolid(xPos(lBall) + 16, yPos(lBall) - 16) Then
        If GetTileAtPoint(xPos(lBall) + 16, yPos(lBall) - 16) = DEADLY Then KillMe
        xSpeed(lBall) = -curSpeed
    ElseIf isTileSolid(xPos(lBall) - 16, yPos(lBall) - 16) Then
        If GetTileAtPoint(xPos(lBall) - 16, yPos(lBall) - 16) = DEADLY Then KillMe
        xSpeed(lBall) = curSpeed
    End If
    
Dim sTestX As Single
Dim sTestY As Single
Dim i As Long
Dim u As Long
    For i = 1 To GetShellCount()
        If GetShellisShelled(i) Then
            sTestX = GetShellX(i) - xPos(lBall)
            sTestY = Abs(GetShellY(i) - yPos(lBall))
            If sTestY <= 32 Then
                If sTestX > 0 And sTestX <= 32 Then xSpeed(lBall) = -curSpeed
                If sTestX < 0 And sTestX >= -32 Then xSpeed(lBall) = curSpeed
            End If
        End If
    Next i
    
    ySpeed(lBall) = ySpeed(lBall) + 0.3
    If ySpeed(lBall) > 5 Then ySpeed(lBall) = 5
    yPos(lBall) = yPos(lBall) + ySpeed(lBall)
    
    If isTileSolid(CLng(xPos(lBall)), CLng(yPos(lBall)), True) Then
        yPos(lBall) = (__intDiv(yPos(lBall) , 32) * 32)
        ySpeed(lBall) = 0
        If GetTileAtPoint(CLng(xPos(lBall)), CLng(yPos(lBall))) = DEADLY Then KillMe
    End If
    
    If bWatchForEdges Then
        If ySpeed(lBall) = 0 And xSpeed(lBall) <> 0 Then
            If Not isTileSolid(CLng(xPos(lBall)) + IIf(xSpeed(lBall) > 0, 16, -16), CLng(yPos(lBall)) + 2, True) Then
                If xSpeed(lBall) > 0 Then xSpeed(lBall) = -curSpeed Else xSpeed(lBall) = curSpeed
            End If
        End If
    End If
    
End Sub
