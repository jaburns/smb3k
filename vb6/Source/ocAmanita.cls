VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ocAmanita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xPos As Single
Private yPos As Single
Private xSpeed As Single
Private ySpeed As Single
Private growVal As Long
Private bActive As Boolean
Private imStyle As Integer
Private isReserved As Boolean


Public Sub CreateAt(X As Single, Y As Single, bFaceRight As Boolean, mStyle As Integer, Optional ByVal inYSpeed As Single = -1, Optional ByVal bReserved As Boolean = False)
    xPos = X
    yPos = Y
    imStyle = mStyle
    If imStyle > 1 Then imStyle = 1
    If imStyle < -1 Then imStyle = -1
    If bFaceRight Then xSpeed = 2 Else xSpeed = -2
    isReserved = bReserved
    If inYSpeed = -1 And Not bReserved Then
        ySpeed = 0.01
        growVal = 1
        PlaySound Sounds.Sprout
    Else
        If bReserved Then ySpeed = 2 Else ySpeed = inYSpeed
        growVal = 33
        PlaySound Sounds.ReleaseItem
    End If
    bActive = True
End Sub

Public Sub HandleMushroom()
Dim tTile As udeLTileType
Dim yDraw As Single

    If Not bActive Then Exit Sub
    
    Select Case imStyle
        Case -1
            yDraw = 64
        Case 0
            yDraw = 0
        Case 1
            yDraw = 32
    End Select
    
    If growVal < 32 Then
        growVal = growVal + 1
        GFX.DrawSurface surfList.Sprites, 0, yDraw, 32, growVal, (xPos - screenLeft) - 16, ((yPos - screenTop) - growVal)
        Exit Sub
    End If

    If Not isReserved Then
    
        If GetLevelSideWarp Then
            If xPos <= -1 Then xPos = GetLevelWidth
            If xPos >= GetLevelWidth + 1 Then xPos = 0
        Else
            If xPos <= 16 Then 
                xPos = 16
                xSpeed = 2
            End If
            If xPos >= GetLevelWidth - 16 Then 
                xPos = GetLevelWidth - 16
                xSpeed = -2
            End If
        End If
        
        xPos = xPos + xSpeed
        If isTileSolid(xPos + 16, yPos - 16) Then xSpeed = -2
        If isTileSolid(xPos - 16, yPos - 16) Then xSpeed = 2
    
        ySpeed = ySpeed + 0.2
        If ySpeed > 5 Then ySpeed = 5
        
    End If
        
    yPos = yPos + ySpeed
    
    If Not isReserved Then
        If isTileSolid(CLng(xPos), CLng(yPos), True) Then
            yPos = ((yPos \ 32) * 32)
            ySpeed = 0
        End If
    End If
    
    If Abs(xPos - Mario.xPos) < 32 And Abs(yPos - Mario.yPos) < 32 And Mario.bAlive Then
        bActive = False
        Select Case imStyle
            Case -1
                Mario.Injure
            Case 0
                If Mario.mStatus = MarioSmall Then
                    Mario.mStatus = MarioBig
                Else
                    marioReserveItem = rsvMushroom
                End If
                PlaySound Sounds.Powerup
                If Not isReserved Then Mario.addShield 4, False
            Case 1
                PlaySound Sounds.OneUp
                gameLives = gameLives + 1
        End Select
        Exit Sub
    End If
    
    GFX.DrawSurface surfList.Sprites, 0, yDraw, 32, 32, (xPos - screenLeft) - 16, (yPos - screenTop) - 32
    
End Sub
