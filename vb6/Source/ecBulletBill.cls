VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecBulletBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SEEKER_ACCEL As Single = 0.11
Private Const BULLET_SPEED As Single = 5

Private xInit As Single
Private yInit As Single

Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single

Private lCount As Long
Private lNextCount As Long

Private lFrame As Long
Private bFrameOn As Boolean

Private bSeeker As Boolean
Private lSeekDir As Long
Private bHasSeeked As Boolean

Private bActive As Boolean
Private bDying As Boolean

Private lFaceDir As Long


Public Sub CreateAt(X As Long, Y As Long, Optional ByVal bIsSeeker As Boolean = False)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    bSeeker = bIsSeeker
    bActive = False
    bDying = False
    bHasSeeked = False
    lSeekDir = 0
    Randomize
    lCount = Int(Rnd * 50)
    lNextCount = 30 + Int(Rnd * 20)
End Sub



Public Sub HandleMe()
Dim i As Long
    
    If Not bActive Then
    
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then Exit Sub
        
        lCount = lCount + 1
        Randomize
        If lCount >= lNextCount Then
            lCount = 0
            Randomize
            lNextCount = 30 + Int(Rnd * 20)
            bActive = True
            bDying = False
            xPos = xInit + IIf(Mario.xPos > xInit, 16, -16)
            yPos = yInit
            xSpeed = IIf(Mario.xPos > xInit, BULLET_SPEED, -BULLET_SPEED)
            PlaySound Sounds.BulletBillFire
            bHasSeeked = False
            lSeekDir = 0
        End If
        
    Exit Sub
    End If
    
    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then bActive = False

    If bDying Then
        ySpeed = ySpeed + 0.3
        If ySpeed > 7 Then ySpeed = 7
        yPos = yPos + ySpeed
        If yPos > screenTop + 512 Then
            bActive = False
            bDying = False
            Exit Sub
        End If
        GFX.DrawSurface surfList.EnemyList.SurfBulletBill, 32, IIf(xSpeed > 0, 32, 0), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 16
        Exit Sub
    End If
    
    If bSeeker And Not bHasSeeked Then
        Select Case lSeekDir
            Case 0
                If xSpeed = BULLET_SPEED And Mario.xPos < xPos Then
                    lSeekDir = -1
                ElseIf xSpeed = -BULLET_SPEED And Mario.xPos > xPos Then
                    lSeekDir = 1
                End If
            Case -1
                xSpeed = xSpeed - SEEKER_ACCEL
                If xSpeed <= -BULLET_SPEED Then
                    xSpeed = -BULLET_SPEED
                    bHasSeeked = True
                End If
            Case 1
                xSpeed = xSpeed + SEEKER_ACCEL
                If xSpeed >= BULLET_SPEED Then
                    xSpeed = BULLET_SPEED
                    bHasSeeked = True
                End If
        End Select
    End If
    xPos = xPos + xSpeed
           
    With Mario
    If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
        If .yPos >= yPos - 16 And .yPos < yPos - 4 Then
            If .bInvinsible Then
                KillMe
            Else
                .makeJump IIf(GameKeyDown(Jump), ENEMYBIGBOUNCE, ENEMYBOUNCE)
                KillMe
            End If
        ElseIf .yPos >= yPos - 4 And .yPos < yPos + IIf(Not .isTall, 48, 68) Then
            If .bInvinsible Then KillMe Else .Injure
        End If
    End If
    End With
    
    For i = 0 To 1
        If getMarioHammerActive(i) Then
            If getMarioHammerX(i) > xPos - 16 And getMarioHammerX(i) < xPos + 16 And getMarioHammerY(i) > yPos - 16 And getMarioHammerY(i) < yPos + 16 Then KillMe
        End If
        If getMarioFireActive(i) Then
            If getMarioFireX(i) > xPos - 16 And getMarioFireX(i) < xPos + 16 And getMarioFireY(i) > yPos - 16 And getMarioFireY(i) < yPos + 16 Then killMarioFire i
        End If
    Next i

    For i = 1 To GetShellCount
        If GetShellActive(i) Then
            If Abs(GetShellX(i) - xPos) < 32 And Abs(GetShellY(i) - yPos - 16) < 20 Then
                KillMe
                DestroyShellIfCarrying (i)
            End If
        End If
    Next i
    
    If bSeeker Then
        HitSpecialBlock CLng(xPos + IIf(xPos > 0, 16, -16)), CLng(yPos), xSpeed > 0
        If lSeekDir = 0 Then
            lFaceDir = IIf(xSpeed > 0, 32, 0)
        Else
            lFaceDir = IIf(lSeekDir = 1, 32, 0)
        End If
        lFrame = lFrame + 1
        If lFrame >= 2 Then 
            lFrame = 0
            bFrameOn = Not bFrameOn
        End If
        If bFrameOn Then
            GFX.DrawSurface surfList.EnemyList.SurfBulletBill, CSng(lFaceDir), 64, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 16
        Else
            GFX.DrawSurface surfList.EnemyList.SurfBulletBill, 0, CSng(lFaceDir), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 16
        End If
    Else
        lFaceDir = IIf(xSpeed > 0, 32, 0)
        GFX.DrawSurface surfList.EnemyList.SurfBulletBill, 0, CSng(lFaceDir), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 16
    End If
    
End Sub



Private Sub KillMe()
    PlaySound Sounds.BumpOff
    bDying = True
    ySpeed = -2
End Sub








