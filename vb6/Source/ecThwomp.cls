VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecThwomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum udeThwompStat
    thpNothing
    thpFalling
    thpRising
End Enum

Private Enum udeTileReturn
    trNothing
    trSolid
    trDeath
End Enum

Private xPos As Single
Private yInit As Single
Private yPos As Single
Private ySpeed As Single

Private lAnger As Long

Private myStat As udeThwompStat


Public Sub CreateAt(X As Long, Y As Long)
    xPos = CSng(X)
    yInit = CSng(Y)
    yPos = yInit
    lAnger = 0
    myStat = thpNothing
End Sub


Public Sub HandleMe()

    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
        
        Exit Sub
    End If

    If Abs(Mario.xPos - xPos) <= 32 Then
        lAnger = 2
    ElseIf Abs(Mario.xPos - xPos) < 64 Then
        lAnger = 1
    Else
        lAnger = 0
    End If
    
    If (Not myStat = thpNothing) Or lAnger = 2 Then
        If myStat = thpNothing Then myStat = thpFalling
        If myStat = thpFalling Then
            lAnger = 2
            ySpeed = ySpeed + 0.45
            yPos = yPos + ySpeed
            If getTileSolid(CLng(xPos - 16), CLng(yPos)) = trSolid Or getTileSolid(CLng(xPos + 16), CLng(yPos)) = trSolid Then
                yPos = ((yPos \ 32) * 32)
                ySpeed = 0
                myStat = thpRising
                PlaySound Sounds.Thwomp
            End If
        Else
            If lAnger = 2 Then lAnger = 1
            yPos = yPos - 2
            If yPos <= yInit Then
                yPos = yInit
                myStat = thpNothing
            End If
        End If
    End If
    
    With Mario
    If Abs(.xPos - xPos) < 40 And .bAlive Then
        If .yPos >= yPos - 64 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
            If Not .bInvinsible Then .Injure
        End If
    End If
    End With
    
    GFX.DrawSurface surfList.EnemyList.Thwomp, lAnger * 64, 0, 64, 64, xPos - screenLeft - 32, yPos - screenTop - 64
    
End Sub



Private Function getTileSolid(xTile As Long, yTile As Long) As udeTileReturn
    getTileSolid = IIf(isTileSolid(xTile, yTile, True), trSolid, trNothing)
End Function
