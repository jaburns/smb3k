VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecGhostDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const DISC_SPEED As Long = 3
Private Const DISC_RADIUS As Long = 128


Private xPos As Single
Private yPos As Single
Private sAngle As Single
Private lFrameCount As Long
Private lFrame As Long
Private bDual As Boolean


Public Sub CreateAt(X As Long, Y As Long)
    xPos = CSng(X)
    yPos = CSng(Y)
    sAngle = 0
    lFrameCount = 0
    lFrame = 0
End Sub



Public Sub HandleMe()
Dim DX As Single
Dim dY As Single
Dim i As Long

    If xPos < screenLeft - 128 - DISC_RADIUS Or xPos > screenLeft + 768 + DISC_RADIUS _
    Or yPos < screenTop - 128 - DISC_RADIUS Or yPos > screenTop + 608 + DISC_RADIUS Then Exit Sub

    lFrameCount = lFrameCount + 1
    If lFrameCount >= 2 Then
        lFrameCount = 0
        lFrame = lFrame + 1
        If lFrame > 3 Then lFrame = 0
    End If
    
    sAngle = sAngle + DISC_SPEED
    DX = xPos + (Cos(sAngle * 3.141592658979 / 180) * DISC_RADIUS)
    dY = yPos + (Sin(sAngle * 3.141592658979 / 180) * DISC_RADIUS)
    
    
    With Mario
    If .xPos >= DX - 27 And .xPos <= DX + 28 And .bAlive Then
        If .yPos >= dY - 14 And .yPos < dY + IIf(Not .isTall, 32, 52) + 14 Then
            If Not .bInvinsible Then .Injure
        End If
    End If
    End With
        
    GFX.DrawSurface surfList.EnemyList.RotoDisc, lFrame * 32, 0, 32, 32, DX - screenLeft - 16, dY - screenTop - 16
   
    
End Sub

