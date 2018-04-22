VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocVineMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private xSrc As Long
Private ySrc As Long
Private xPos As Single
Private yPos As Single
Private yDiff As Long
Private growVal As Long
Private bActive As Boolean
Private curAFrame As Single
Private curAFrameCount As Long
Private bLastRun As Boolean

Public Sub CreateAt(X As Single, Y As Single)
    bActive = True
    xPos = X
    yPos = Y
    xSrc = oCurWorldData.LevelData(curLevel).dfVine.xSrc
    ySrc = oCurWorldData.LevelData(curLevel).dfVine.ySrc
    growVal = 0
    PlaySound Sounds.Sprout
    curAFrame = 32
    curAFrameCount = 0
    bLastRun = False
End Sub

Public Sub HandleViney()
Dim tType As udeLTileType

    If yPos = 0 Then bActive = False
    If Not bActive Then Exit Sub
    
    curAFrameCount = curAFrameCount + 1
    If curAFrameCount >= 10 Then
        curAFrameCount = 0
        If curAFrame = 32 Then curAFrame = 64 Else curAFrame = 32
    End If
    
    If growVal < 32 Then
        growVal = growVal + 2
        GFX.DrawSurface surfList.Sprites, curAFrame, 64, 32, growVal, (xPos - screenLeft) - 16, ((yPos - screenTop) - growVal)
        Exit Sub
    End If
    
    yPos = yPos - 2
    yDiff = yDiff + 2
    
    If yDiff >= 32 Then
           
        SetTile CLng(xPos), CLng(yPos) + 2, xSrc, ySrc
        
        If bLastRun Then
            bActive = False
            yDiff = 0
            Exit Sub
        End If
        
        tType = GetTileAtPoint(CLng(xPos), CLng(yPos) - 34)
        If isTileSolid(CLng(xPos), CLng(yPos) - 34) Then
            bLastRun = True
            yDiff = 0
            Exit Sub
        End If
        
    End If
    
    If bLastRun Then
        GFX.DrawSurface surfList.Sprites, curAFrame, 64 + yDiff, 32, 32 - yDiff, xPos - screenLeft - 16, yPos - screenTop - 32 + yDiff
    Else
        GFX.DrawSurface surfList.Sprites, curAFrame, 64, 32, 32, xPos - screenLeft - 16, yPos - screenTop - 32
    End If
    GFX.DrawSurface surfList.Tileset, xSrc * 32, ((ySrc + 1) * 32) - yDiff, 32, yDiff, xPos - screenLeft - 16, yPos - screenTop
    If yDiff >= 32 Then yDiff = 0

End Sub


