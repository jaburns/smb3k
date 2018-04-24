VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecLavaBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const FIRE_DELAY As Long = 50
Private Const FIRE_POWER As Long = 15


Private xPos As Single
Private yPos As Single
Private yInit As Single
Private ySpeed As Single
Private lNextFireCount As Long

Private lFrame As Long
Private bFrameOn As Boolean
Private bAlive As Boolean
Private bJumping As Boolean


Public Sub CreateAt(x As Long, y As Long)
    xPos = CSng(x)
    yInit = CSng(y)
    yPos = yInit
    Randomize
    lNextFireCount = Int(Rnd * FIRE_DELAY)
    ySpeed = 0
    lFrame = 0
    bFrameOn = False
End Sub



Public Sub HandleMe()
Dim i As Long

    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then
        bJumping = False
        yPos = yInit
        lNextFireCount = 0
        Exit Sub
    End If
    
    lFrame = lFrame + 1
    If lFrame >= 2 Then
        lFrame = 0
        bFrameOn = Not bFrameOn
    End If
    
    If bJumping Then
    
        ySpeed = ySpeed + 0.3
        yPos = yPos + ySpeed
        
        If yPos >= yInit Then
            bJumping = False
            Exit Sub
        End If
        
        With Mario
        If .xPos >= xPos - 28 And .xPos <= xPos + 28 And .bAlive Then
            If .yPos >= yPos - 32 And .yPos < yPos + IIf(Not .isTall, 32, 52) Then
                If Not .bInvinsible Then .Injure
            End If
        End If
        End With
        
        GFX.DrawSurface surfList.EnemyList.LavaBubble, IIf(bFrameOn, 0, 32), IIf(ySpeed > 0, 32, 0), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
        
    Else
    
        lNextFireCount = lNextFireCount + 1
        If lNextFireCount >= FIRE_DELAY Then
            lNextFireCount = 0
            ySpeed = CSng(-FIRE_POWER)
            bJumping = True
        End If
        
    End If
    
End Sub




