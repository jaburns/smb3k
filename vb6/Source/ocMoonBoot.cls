VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocMoonBoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos As Single
Private yPos As Single
Private growVal As Long
Private bActive As Boolean
Private isReserved As Boolean


Public Sub CreateAt(x As Single, y As Single, bFaceRight As Boolean, Optional ByVal bReserved As Boolean = False)
    xPos = x
    yPos = y
    growVal = 1
    bActive = True
    isReserved = bReserved
    If bReserved Then
        growVal = 33
        PlaySound Sounds.ReleaseItem
    Else
        PlaySound Sounds.Sprout
    End If
End Sub

Public Sub HandleBoot()
Dim tTile As udeLTileType

    If Not bActive Then Exit Sub
    
    If growVal < 32 Then
        growVal = growVal + 1
        GFX.DrawSurface surfList.Sprites, 32, 32, 32, growVal, (xPos - screenLeft) - 16, ((yPos - screenTop) - growVal)
        Exit Sub
    End If

    If isReserved Then yPos = yPos + 2
    
    If Abs(xPos - Mario.xPos) < 32 And Abs(yPos - Mario.yPos) < 32 And Mario.bAlive Then
        bActive = False
        ReserveMariosItem
        Mario.mStatus = MarioMoonboot
        PlaySound Sounds.Powerup
        If Not isReserved Then Mario.addShield 4, False
        Exit Sub
    End If
    
    GFX.DrawSurface surfList.Sprites, 32, 32, 32, 32, (xPos - screenLeft) - 16, (yPos - screenTop) - 32

End Sub

