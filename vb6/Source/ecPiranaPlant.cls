VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecPiranaPlant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private xPos As Single
Private yPos As Single
Private lNextBiteCount As Long
Private lBiteValue As Long
Private lBiteDelay As Long

Private lFrame As Long
Private bFrameOn As Boolean
Private lMaxFrame As Long

Private bAlive As Boolean
Private bMean As Boolean
Private bDown As Boolean

Private xSrc As Single

Public Sub CreateAt(X As Long, Y As Long, bVicious As Boolean, bUpsideDown As Boolean)

    xPos = CSng(X)
    yPos = CSng(Y) - IIf(bUpsideDown, 32, 0)
    
    bFrameOn = False
    lFrame = 0
    bDown = bUpsideDown
    
    bMean = bVicious
    If bMean Then
        xSrc = 64
        lBiteDelay = 40
        lMaxFrame = 3
    Else
        xSrc = 0
        lBiteDelay = 60
        lMaxFrame = 5
    End If
    
    Randomize
    lNextBiteCount = Int(Rnd() * lBiteDelay)
    lBiteValue = 0
    bAlive = True
    
End Sub



Public Sub HandleMe()
Dim lDraw As Long
Dim i As Long

    If Not bAlive Then
        If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
            bAlive = True
            lNextBiteCount = 0
            lBiteValue = 0
        End If
        Exit Sub
    End If

    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then Exit Sub
    
    lFrame = lFrame + 1
    If lFrame >= lMaxFrame Then
        lFrame = 0
        bFrameOn = Not bFrameOn
    End If
    
    If lBiteValue Then
    
        If lBiteValue > 64 Then lDraw = 128 - lBiteValue Else lDraw = lBiteValue
        If bDown Then
            GFX.DrawSurface surfList.EnemyList.PiranaPlants, xSrc + IIf(bFrameOn, 32, 0), 128 - lDraw, 32, lDraw, xPos - screenLeft - 16, yPos - screenTop
        Else
            GFX.DrawSurface surfList.EnemyList.PiranaPlants, xSrc + IIf(bFrameOn, 32, 0), 0, 32, lDraw, xPos - screenLeft - 16, yPos - screenTop - lDraw
        End If
        lBiteValue = lBiteValue + IIf(bMean, 2, 1)
        If lBiteValue >= 128 Then lBiteValue = 0
        
        With Mario
        If Abs(.xPos - xPos) < 28 And .bAlive Then
            If .yPos >= yPos - IIf(bDown, 0, lDraw) And .yPos < yPos + IIf(Not .isTall, 32, 52) + IIf(bDown, lDraw, 0) Then
                If .bInvinsible Then KillMe Else .Injure
            End If
        End If
        End With
    
        For i = 1 To GetShellCount
            If GetShellActive(i) Then
                If Abs(GetShellX(i) - xPos) < 32 Then
                    If GetShellY(i) > yPos - IIf(bDown, 0, lDraw) And GetShellY(i) < yPos + 32 + IIf(bDown, lDraw, 0) Then
                        KillMe
                        DestroyShellIfCarrying (i)
                    End If
                End If
            End If
        Next i
        If inBombBlast(xPos, yPos + IIf(bDown, lDraw, -lDraw)) Then KillMe
    
        For i = 0 To 1
            If getMarioHammerActive(i) Then
                If getMarioHammerX(i) > xPos - 16 And getMarioHammerX(i) < xPos + 16 And getMarioHammerY(i) > yPos - IIf(bDown, 0, lDraw) And getMarioHammerY(i) < yPos + IIf(bDown, lDraw, 0) Then KillMe
            End If
            If getMarioFireActive(i) Then
                If getMarioFireX(i) > xPos - 16 And getMarioFireX(i) < xPos + 16 And getMarioFireY(i) > yPos - IIf(bDown, 0, lDraw) And getMarioFireY(i) < yPos + IIf(bDown, lDraw, 0) Then
                    killMarioFire i
                    KillMe
                End If
            End If
        Next i
        
    ElseIf Abs(Mario.xPos - xPos) > 50 Or bMean Or bDown Then
    
        lNextBiteCount = lNextBiteCount + 1
        If lNextBiteCount >= lBiteDelay Then
            lNextBiteCount = 0
            lBiteValue = 1
        End If
        
    End If
    
End Sub



Private Sub KillMe()
    PlaySound Sounds.BumpOff
    bAlive = False
End Sub

