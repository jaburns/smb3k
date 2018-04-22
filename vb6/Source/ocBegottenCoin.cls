VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocBegottenCoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos As Single
Private yPos As Single
Private sAccel As Single
Private bActive As Boolean
Private sAlpha As Single

Public Sub CreateAt(x As Single, y As Single)
bActive = True
xPos = x
sAlpha = 100
sAccel = 2
yPos = y
PlaySound Sounds.Coin
End Sub

Public Sub HandleMiniCoin()

    If Not bActive Then Exit Sub
    sAccel = sAccel - 0.1
    yPos = yPos - sAccel
    If yPos <= 0 Then bActive = False
    sAlpha = sAlpha - 5
    If sAlpha <= 0 Then bActive = False
    If Not bActive Then Exit Sub
            
    GFX.DrawSurface surfList.Objects, 0, 167, 16, 16, xPos - screenLeft - 8, yPos - screenTop - 16, , , , sAlpha

End Sub

