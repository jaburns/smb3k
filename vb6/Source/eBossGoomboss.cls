VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eBossGoomboss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const MAX_INJURIES As Long = 4

Private Enum BOSS_STAT
    Sitting
    Spitting
    Waiting
    Shelling
End Enum

Private xInit As Single
Private yInit As Single
Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private lSpitCount As Long
Private bSpitLeft As Boolean
Private bShellLeft As Boolean
Private bFrameOn As Boolean
Private lFrameCount As Long
Private lHurtCount As Long
Private curStat As BOSS_STAT
Private lStatCount As Long
Private ownShell As ecKoopa
Private lShellCount As Long
Private lInjuries As Long
Private bDone As Boolean
Private lBlockDestroyCount As Long

Private gCount As Long


Private Sub Class_Initialize()
    Set ownShell = New ecKoopa
    ownShell.setShellSpeed 4
End Sub

Private Sub Class_Terminate()
    Set ownShell = Nothing
End Sub


Public Sub CreateAt(X As Long, Y As Long)
Dim i As Long
    xPos = CSng(X)
    yPos = CSng(Y)
    xInit = xPos
    yInit = yPos
    lSpitCount = 0
    xSpeed = 0
    curStat = Sitting
    lStatCount = 0
    bSpitLeft = False
    lInjuries = 0
    bDone = False
    bShellLeft = False
End Sub


Public Sub HandleMe()
Dim xSrc As Single
Dim i As Long

    If bDone Then Exit Sub
    
    If lHurtCount = 0 And lInjuries = MAX_INJURIES Then
        lStatCount = lStatCount + 1
        If lStatCount >= 5 Then
            lStatCount = 0
            With oCurWorldData.LevelData(curLevel)
                If lBlockDestroyCount >= 8 And lBlockDestroyCount <= 11 Then
                    SetTile 16 + (32 * lBlockDestroyCount), 176, .dfBrick.xSrc, .dfBrick.ySrc
                Else
                    SetTile 16 + (32 * lBlockDestroyCount), 176, .dfCoin.xSrc, .dfCoin.ySrc
                End If
                BreakBrick 16 + (32 * lBlockDestroyCount), 176
                PlaySound Sounds.BreakBrick
            End With
            lBlockDestroyCount = lBlockDestroyCount + 1
            If lBlockDestroyCount = 20 Then bDone = True
        End If
        Exit Sub
    End If
    
    lFrameCount = lFrameCount + 1
    If lFrameCount = 10 Then
        lFrameCount = 0
        bFrameOn = Not bFrameOn
    End If
    
    If lHurtCount = 0 Then
        Select Case curStat
            Case Sitting
                lStatCount = lStatCount + 1
                If lStatCount > 120 Then
                    lStatCount = 0
                    curStat = Spitting
                End If
            Case Spitting
                gCount = gCount + 1
                If gCount >= 70 - (lInjuries * 5) Then
                    gCount = 0
                    MakeGoomba xPos + 64, yPos + 160, (xSpeed - 3.5) * IIf(bSpitLeft, 1, -1), -6
                    lSpitCount = 10
                    bSpitLeft = Not bSpitLeft
                End If
                If xPos >= 256 Then
                    If xSpeed < 3 Then xSpeed = xSpeed + 0.02
                    xPos = xPos + xSpeed
                Else
                    If xSpeed > 0.5 Then xSpeed = xSpeed - 0.01
                    xPos = xPos + xSpeed
                    If xPos > xInit Then
                        xPos = xInit
                        curStat = Shelling
                    End If
                End If
                If xPos > 704 Then xPos = xPos - 832
            Case Shelling
                lStatCount = lStatCount + 1
                If lStatCount > 80 Then
                    lStatCount = 0
                    If Not ownShell.isAlive Then
                        bShellLeft = Not bShellLeft
                        ownShell.createShell xPos + 64, yPos + 160, IIf(bShellLeft, -4, 4), -8
                        lSpitCount = 10
                        lShellCount = 100
                    End If
                    curStat = Sitting
                End If
        End Select
        
        If lShellCount > 0 Then lShellCount = lShellCount - 1
        If lShellCount = 0 Then
            With ownShell
                If .xLoc > xPos And .xLoc < (xPos + 128) Then
                    If .yLoc > yPos And .yLoc < yPos + 160 Then
                        lHurtCount = 50
                        ownShell.KillMe
                        lShellCount = 100
                        lInjuries = lInjuries + 1
                        If lInjuries = MAX_INJURIES Then
                            blBossesPassed = blBossesPassed Or BP_1
                            PlaySound Sounds.BossDie
                            lStatCount = 0
                        Else
                            PlaySound Sounds.BossHit
                        End If
                    End If
                End If
            End With
        End If
    End If
    

    If lSpitCount > 0 Then lSpitCount = lSpitCount - 1
    If lHurtCount > 0 Then lHurtCount = lHurtCount - 1
    
    If Mario.yPos - IIf(Mario.isTall, 52, 32) < 192 Then
        Mario.ResetYPos Mario.yPos + 5, 1
        PlaySound Sounds.Bump
    End If
            
    If lHurtCount Then 
        xSrc = IIf(__intDiv(lHurtCount , 2), 384, 0)
    Else 
        If lSpitCount Then xSrc = 256 Else xSrc = IIf(bFrameOn, 0, 128)
    End If

    GFX.DrawSurface surfList.EnemyList.BOSSGoomboss, xSrc, 0, 128, 160, xPos - screenLeft, yPos - screenTop + 3

    ownShell.HandleMe
    
End Sub


