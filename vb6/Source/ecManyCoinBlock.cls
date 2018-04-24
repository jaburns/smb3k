VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecManyCoinBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAXCOINTIME As Long = 250

Private xSrc As Long
Private ySrc As Long
Private xPos As Single
Private yPos As Single
Private bIsAnimated As Boolean
Private lTimeCount As Long
Private lBounceCount As Long
Private bBounceUp As Boolean
Private bDead As Boolean



Public Sub CreateAt(X As Long, Y As Long, xTile As Long, yTile As Long, bAnimated As Boolean)
    xSrc = xTile
    ySrc = yTile
    xPos = CSng(X)
    yPos = CSng(Y)
    bIsAnimated = bAnimated
    lTimeCount = 0
    bDead = False
End Sub


Public Function activeAtLocation(sXPos As Single, sYPos As Single) As Boolean
    activeAtLocation = (32 * (sXPos \ 32) = xPos And 32 * (sYPos \ 32) = yPos)
End Function


'returns true if its being hit
Public Function hitMe(lXPos As Long, lYPos As Long) As Boolean
    hitMe = False
    If (32 * (lXPos \ 32) = xPos And 32 * (lYPos \ 32) = yPos) Then
        hitMe = True
        If bDead Then Exit Function
        If lTimeCount = 0 Then lTimeCount = 1
        MakeLittleCoin ((lXPos \ 32) * 32) + 16, (lYPos \ 32) * 32
        PlaySound Sounds.Coin
        lBounceCount = 1
        bBounceUp = True
        gameCoins = gameCoins + 1
        If lTimeCount > MAXCOINTIME Then
            bIsAnimated = False
            bDead = True
            lTimeCount = 0
            lBounceCount = 0
            xSrc = xSrc + 1
        End If
    End If
End Function



Public Sub HandleMe()
Dim tType As udeLTileType
Dim ySource As Single
Dim i As Long
    
    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then Exit Sub
    
    If lTimeCount > 0 Then lTimeCount = lTimeCount + 1
    If lBounceCount > 0 Then
        If bBounceUp Then
            lBounceCount = lBounceCount + 1
            If lBounceCount > 4 Then bBounceUp = False
        Else
            lBounceCount = lBounceCount - 1
        End If
    End If
    
    If bIsAnimated Then ySource = (ySrc + curTileFrame) * 32 Else ySource = ySrc * 32
    GFX.DrawSurface surfList.Tileset, xSrc * 32, ySource, 32, 32, xPos - screenLeft, yPos - screenTop - lBounceCount
    'GFX.DrawText CStr(lTimeCount), xPos - screenLeft, yPos - screenTop - lBounceCount, 0
        
End Sub




