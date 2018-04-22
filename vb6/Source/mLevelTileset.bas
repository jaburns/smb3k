Attribute VB_Name = "mLevelTileset"
Option Explicit

Public Enum udeLTileTag
    NO_ITEM = 0
    FLOWER = 1
    PowerStar = 2
    ENDWHEEL = 3
    MOONBOOT = 4
    POWERHAMMER = 5
    BEANSTALK = 6
    BRICKCOIN = 7
    PIPE1TAG = 8
    PIPE2TAG = 9
    PIPE3TAG = 10
    PIPE4TAG = 11
    PIPE5TAG = 12
    PIPE6TAG = 13
    PIPE7TAG = 14
    PIPE8TAG = 15
    PIPE9TAG = 16
    PIPE10TAG = 17
    PIPE11TAG = 18
    PIPE12TAG = 19
    PIPE13TAG = 20
    PIPE14TAG = 21
    PIPE15TAG = 22
    PIPE16TAG = 23
    HAS1UP = 24
    HAS1DOWN = 25
    PLATFORMSLOW = 26
    PLATFORMFAST = 27
    PPLANTMEAN = 28
    KOOPASHELL = 29
    HASWINGCAP = 30
    ONLYSHROOMS = 31
    WINGEDKOOPA = 32
    BULLETSEEKER = 33
    ALLWATERTAG = 34
    MANYCOINSTAG = 35
    BOSSTILEHIDE = 36
    ENEMYSMARTTAG = 37
End Enum

Public Enum udeLTileType
    Background = 0
    SOLID = 1
    DEADLY = 2
    WATER = 3
    ICE = 4
    Coin = 5
    BRICK = 6
    COINBLOCK = 7
    VINE = 8
    INJURETILE = 9
    SOLIDONTOP = 10
    BOUNCY = 11
    DONUTTILE = 12
    SLANTLEFT = 13
    SLANTRIGHT = 14
End Enum

Public Enum udeLTileEnemy
    NO_ENEMY = 0
    OBJGOOMBA = 1
    OBJPPLANT = 2
    OBJBUMPTY = 3
    OBJFIREBALL = 4
    OBJSPINEY = 5
    OBJLITTLEBOO = 6
    OBJKOOPA = 7
    OBJBEETLE = 8
    OBJSMARTKOOPA = 9
    OBJDRYBONES = 10
    OBJTHWOMP = 11
    OBJPOWBUTTON = 12
    OBJPLATFORM2 = 13
    OBJPLATFORM3 = 14
    OBJPLATFORM4 = 15
    OBJPLATFORM5 = 16
    OBJPLATFORMUP = 17
    OBJPLATFORMDOWN = 18
    OBJPLATFORMLEFT = 19
    OBJPLATFORMRIGHT = 20
    OBJPLATFORMSTOP = 21
    OBJPLATFORMDROP = 22
    OBJROTODISC = 23
    OBJTHRUFISH = 24
    OBJBLOCKFISH = 25
    OBJRISINGLAVA = 26
    OBJSINUSOIDLAVA = 27
    OBJTALLCACTUS = 28
    OBJKOOPABOUNCER = 29
    OBJBULLETBILL = 30
    OBJBOBOMB = 31
    OBJUPSIDEDOWNPLANT = 32
    OBJBOSSGOOMBOSS = 33
    OBJHIDETILE = 34
    OBJWIGGLER = 35
    OBJENDFLYING = 36
    OBJFIRETILE = 37
    OBJSAVEPOINT = 38
End Enum

Public Type udtLTileSetDataColumn
    Row() As udeLTileType
End Type
Public Type udtLTileSetData
    bWidth As Byte
    bHeight As Byte
    Column() As udtLTileSetDataColumn
End Type


Public Sub LoadLevelTileset(oTileset As udtLTileSetData, sPath As String)
Dim lFile As Long: lFile = FreeFile
Dim xx As Long
Dim yy As Long
Dim byt_ As Byte
    
    Open sPath For Input As lFile
    Close lFile
    Open sPath For Binary Access Read Lock Write As lFile
    With oTileset
    
        Get #lFile, 1, byt_
        .bWidth = byt_
        Get #lFile, , byt_
        .bHeight = byt_
        
        ReDim .Column(.bWidth)
        
        For xx = 0 To .bWidth
            ReDim Preserve .Column(xx).Row(.bHeight)
            For yy = 0 To .bHeight
                Get #lFile, , byt_
                .Column(xx).Row(yy) = CLng(byt_)
            Next yy
        Next xx
    
    End With
    
    Close lFile
    
End Sub

Public Sub SaveLevelTileset(oTileset As udtLTileSetData, sPath As String)
Dim lFile As Long: lFile = FreeFile
Dim xx As Long
Dim yy As Long
Dim byt_ As Byte
    
    Open sPath For Binary Access Write Lock Read Write As lFile
    With oTileset
        
        Put #lFile, 1, .bWidth
        Put #lFile, , .bHeight
        For xx = 0 To .bWidth
            For yy = 0 To .bHeight
                byt_ = CByte(.Column(xx).Row(yy))
                Put #lFile, , byt_
            Next yy
        Next xx
        
    End With
    Close lFile
    
End Sub



Public Function GetLevelTileType(oTileset As udtLTileSetData, xSrc As Long, ySrc As Long) As udeLTileType
On Error Resume Next
    GetLevelTileType = CLng(oTileset.Column(xSrc).Row(ySrc))
    If GetLevelTileType > 127 Then GetLevelTileType = GetLevelTileType - 128
End Function

Public Function GetLevelTileAnimated(oTileset As udtLTileSetData, xSrc As Long, ySrc As Long) As Boolean
On Error Resume Next
    GetLevelTileAnimated = oTileset.Column(xSrc).Row(ySrc) > 127
End Function
