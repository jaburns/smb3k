VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecPlatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private xInit As Single
Private xPos As Single
Private xSpeed As Single
Private yInit As Single
Private yPos As Single
Private ySpeed As Single
Private lDefaultSpeed As Long

Private bActive As Boolean
Private bFalling As Boolean
Private bStopped As Boolean

Private sSize As Single
Private sWidth As Single


Public Sub CreateAt(X As Long, Y As Long, inSpeed As Single, inSize As Single)
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    lDefaultSpeed = CLng(inSpeed)
    xSpeed = 0
    ySpeed = 0
    sSize = inSize
    sWidth = 32 * (sSize + 2)
    bActive = True
    bFalling = False
    bStopped = False
End Sub



Public Sub HandleMe()
Dim tType As udeLTileEnemy
Dim i As Long
       
    If Not bActive Then
        If xInit < screenLeft - 128 - sWidth Or xInit > screenLeft + 768 _
        Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CLng(xInit), CLng(yInit), CSng(lDefaultSpeed), sSize
        Exit Sub
    End If
    
    If xPos < screenLeft - 256 - sWidth Or xPos > screenLeft + 896 _
    Or yPos < screenTop - 256 Or yPos > screenTop + 736 Then
        bActive = False
        Exit Sub
    End If
        
    If bFalling Then ySpeed = ySpeed + 0.4
    
    AdjustPosition
    xPos = xPos + xSpeed
    yPos = yPos + ySpeed
        
    With Mario
    If (Not bFalling) And .xPos >= xPos - 10 And .xPos <= xPos + sWidth + 10 And .bAlive Then
        If .yPos >= yPos - lDefaultSpeed And .yPos < yPos + 16 And .ySpeed >= 0 Then
        
            If xSpeed = 0 And ySpeed = 0 And bStopped = False Then
                tType = GetTileEnemy((xPos + 48) \ 32, (yPos + 16) \ 32)
                If tType = OBJPLATFORMLEFT Then
                    xSpeed = -lDefaultSpeed
                ElseIf tType = OBJPLATFORMUP Then
                    ySpeed = -lDefaultSpeed
                ElseIf tType = OBJPLATFORMDOWN Then
                    ySpeed = lDefaultSpeed
                Else
                    xSpeed = lDefaultSpeed
                End If
            End If

            .makeStand .xPos + xSpeed, yPos
            If .yPos = yPos + ySpeed And ySpeed < 0 Then
                If isTileSolid(.xPos, .yPos - IIf(Mario.mStatus = MarioSmall, 30, 50)) Then .makeDie
            End If
            
        End If
    End If
    End With
        
    GFX.DrawSurface surfList.EnemyList.MovingPlatform, 0, 0, 32, 32, xPos - screenLeft, yPos - screenTop
    For i = 1 To sSize
        GFX.DrawSurface surfList.EnemyList.MovingPlatform, 32, 0, 32, 32, xPos - screenLeft + (i * 32), yPos - screenTop
    Next i
    GFX.DrawSurface surfList.EnemyList.MovingPlatform, 64, 0, 32, 32, xPos - screenLeft + ((sSize + 1) * 32), yPos - screenTop
        
End Sub



Private Sub AdjustPosition()
Dim i As Long
Dim sOff As Single
Dim tTile As udeLTileEnemy

    If xSpeed Then
        
        If xSpeed > 0 Then sOff = sWidth + lDefaultSpeed Else sOff = -lDefaultSpeed
        tTile = GetTileEnemy((xPos + sOff) \ 32, (yPos + 16) \ 32)
        Select Case tTile
            Case OBJPLATFORMRIGHT
                xSpeed = lDefaultSpeed
            Case OBJPLATFORMLEFT
                xSpeed = -lDefaultSpeed
            Case OBJPLATFORMUP
                xSpeed = 0
                ySpeed = -lDefaultSpeed
            Case OBJPLATFORMDOWN
                xSpeed = 0
                ySpeed = lDefaultSpeed
            Case OBJPLATFORMSTOP
                xSpeed = 0
                ySpeed = 0
                bStopped = True
            Case OBJPLATFORMDROP
                ySpeed = 2
                bFalling = True
        End Select
    
    ElseIf ySpeed Then
    
        If ySpeed > 0 Then sOff = 32 + lDefaultSpeed Else sOff = -lDefaultSpeed
        For i = 0 To (sSize + 1)
            tTile = GetTileEnemy((xPos + 16 + (i * 32)) \ 32, (yPos + sOff) \ 32)
            Select Case tTile
                Case OBJPLATFORMRIGHT
                    xSpeed = lDefaultSpeed
                    ySpeed = 0
                Case OBJPLATFORMLEFT
                    xSpeed = -lDefaultSpeed
                    ySpeed = 0
                Case OBJPLATFORMUP
                    ySpeed = -lDefaultSpeed
                Case OBJPLATFORMDOWN
                    ySpeed = lDefaultSpeed
                Case OBJPLATFORMSTOP
                    xSpeed = 0
                    ySpeed = 0
                    bStopped = True
                Case OBJPLATFORMDROP
                    ySpeed = 2
                    bFalling = True
            End Select
        Next i
    
    End If
    
End Sub


