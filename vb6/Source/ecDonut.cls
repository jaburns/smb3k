VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ecDonut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private xSrc As Long
Private ySrc As Long
Private xInit As Single
Private xPos As Single
Private yInit As Single
Private yPos As Single
Private ySpeed As Single

Private bActive As Boolean
Private bFalling As Boolean
Private lShake As Long
Private lShakeTime As Long


Public Sub CreateAt(X As Long, Y As Long, xTile As Long, yTile As Long)
    xSrc = xTile
    ySrc = yTile
    xInit = CSng(X)
    yInit = CSng(Y)
    xPos = xInit
    yPos = yInit
    ySpeed = 0
    bActive = True
    bFalling = False
    lShake = 0
End Sub


Public Function activeAtLocation(sXPos As Single, sYPos As Single) As Boolean
    activeAtLocation = (32 * (sXPos \ 32) = xInit And 32 * (sYPos \ 32) = yInit And bActive And (Not bFalling))
End Function



Public Sub HandleMe()
Dim tType As udeLTileType
Dim yStand As Single
Dim i As Long
       
    If Not bActive Then
        If xInit < screenLeft - 128 Or xInit > screenLeft + 768 _
        Or yInit < screenTop - 128 Or yInit > screenTop + 608 Then CreateAt CLng(xInit), CLng(yInit), xSrc, ySrc
        Exit Sub
    End If
    
    If xPos < screenLeft - 128 Or xPos > screenLeft + 768 _
    Or yPos < screenTop - 128 Or yPos > screenTop + 608 Then Exit Sub
        
    If bFalling Then
        ySpeed = ySpeed + 0.5
        yPos = yPos + ySpeed
        If yPos > screenTop + 480 Then bActive = False
    End If
        
    With Mario
    If (Not bFalling) And .xPos >= xPos - 10 And .xPos <= xPos + 42 And .bAlive Then
        If .yPos >= yPos - 2 And .yPos < yPos + 16 And .ySpeed >= 0 Then
            .makeStand .xPos, yPos
            If lShake >= 0 Then lShake = -1 Else lShake = 1
            lShakeTime = lShakeTime + 1
            If lShakeTime >= 40 Then bFalling = True
        Else
        lShake = 0
        lShakeTime = 0
        End If
    Else
    lShake = 0
    lShakeTime = 0
    End If
    End With
        
    GFX.DrawSurface surfList.Tileset, xSrc * 32, ySrc * 32, 32, 32, xPos - screenLeft + lShake, yPos - screenTop
        
End Sub


