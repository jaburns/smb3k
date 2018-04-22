VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocFlower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos As Single
Private yPos As Single
Private growVal As Long
Private bActive As Boolean
Private bHammer As Boolean
Private isReserved As Boolean

Public Sub CreateAt(x As Single, y As Single, bIsHammer As Boolean, Optional ByVal bReserved As Boolean = False)
    bActive = True
    xPos = x
    growVal = 1
    isReserved = bReserved
    If bReserved Then
        growVal = 33
        PlaySound Sounds.ReleaseItem
    Else
        PlaySound Sounds.Sprout
    End If
    yPos = y
    bHammer = bIsHammer
End Sub

Public Sub HandleFlower()

    If Not bActive Then Exit Sub
    
    If growVal < 32 Then
        growVal = growVal + 1
        GFX.DrawSurface surfList.Sprites, IIf(bHammer, 64, 32), IIf(bHammer, 32, 0), 32, growVal, (xPos - screenLeft) - 16, ((yPos - screenTop) - growVal)
        Exit Sub
    End If

    If isReserved Then yPos = yPos + 2
    
    If Abs(xPos - Mario.xPos) < 32 And Abs(yPos - Mario.yPos) < 32 And Mario.bAlive Then
        bActive = False
        ReserveMariosItem
        If bHammer Then Mario.mStatus = MarioHammer Else Mario.mStatus = MarioFlower
        PlaySound Sounds.Powerup
        If Not isReserved Then Mario.addShield 4, False
        Exit Sub
    End If
    
    GFX.DrawSurface surfList.Sprites, IIf(bHammer, 64, 32), IIf(bHammer, 32, 0), 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32

End Sub
