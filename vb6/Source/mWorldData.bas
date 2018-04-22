Attribute VB_Name = "mWorldData"
Option Explicit

'Definitions for the world data object: oCurWorldData
Public Enum udeWorldDataPipeExit
    epLeft = 0
    epRight = 1
    epUp = 2
    epDown = 3
    epDoor = 4
End Enum
Public Type udeWorldDataXY
    xSrc As Integer
    ySrc As Integer
End Type
Public Type udtWorldDataPipe
    destLevel As Integer
    destTag As Byte
    destDir As Byte
End Type
Public Type udtWorldDataLevel
    LevelName As String
    Filename As String
    Background As String
    EnemySkin As String
    MusicFile As String
    TileFile As String
    TimeGiven As Integer
    iScrollStyle As Byte
    iScrollSpeed As Integer
    dfCoin As udeWorldDataXY
    dfBrick As udeWorldDataXY
    dfVine As udeWorldDataXY
    PipeDest(1 To 16) As udtWorldDataPipe
End Type
Public Type udtWorldData
    WorldName As String
    LevelData() As udtWorldDataLevel
End Type


Public oCurWorldData As udtWorldData
Public curLevel As Long


Public Sub cwdResetWorldData()
Dim i As Long
    With oCurWorldData
        .WorldName = "New World"
        ReDim .LevelData(0)
    End With
    With oCurWorldData.LevelData(0)
        .Background = ""
        .Filename = ""
        .LevelName = "[NEW]"
        .MusicFile = ""
        .EnemySkin = ""
        .TimeGiven = 200
        For i = 1 To 16
            .PipeDest(i).destLevel = 0
            .PipeDest(i).destTag = 0
        Next i
    End With
End Sub


Public Sub cwdAddLevel(sLevelName As String)
Dim i As Long
    With oCurWorldData
        i = UBound(.LevelData) + 1
        ReDim Preserve .LevelData(i)
    End With
    With oCurWorldData.LevelData(i)
        .iScrollSpeed = 0
        .iScrollStyle = 0
        .LevelName = sLevelName
        .Filename = ""
        .Background = ""
        .MusicFile = ""
        .EnemySkin = ""
        .TileFile = ""
        .TimeGiven = 200
    End With
End Sub


Public Sub cwdRemoveLevel()
    If UBound(oCurWorldData.LevelData) = 0 Then Exit Sub
    ReDim Preserve oCurWorldData.LevelData(UBound(oCurWorldData.LevelData) - 1)
End Sub


Public Function cwdSaveWorldData(sPath As String) As Boolean
On Error GoTo errOut
Dim fFile As Long: fFile = FreeFile
    cwdSaveWorldData = False
    Open sPath For Binary Access Write Lock Read Write As fFile
    Put #fFile, 1, oCurWorldData
    cwdSaveWorldData = True
errOut:
    Close fFile
End Function


Public Function cwdLoadWorldData(sPath As String) As Boolean
On Error GoTo errOut
'Dim oTMP as udtWorldData, i As Long, u As Long
Dim fFile As Long: fFile = FreeFile
    cwdLoadWorldData = False
    Open sPath For Input As fFile: Close fFile
    Open sPath For Binary Access Read Lock Write As fFile
    Get #fFile, 1, oCurWorldData
'        oCurWorldData.WorldName = oTMP.WorldName
'        ReDim oCurWorldData.LevelData(UBound(oTMP.LevelData))
'        For i = 0 To UBound(oCurWorldData.LevelData)
'            oCurWorldData.LevelData(i).Background = oTMP.LevelData(i).Background
'            oCurWorldData.LevelData(i).dfBrick = oTMP.LevelData(i).dfBrick
'            oCurWorldData.LevelData(i).dfCoin = oTMP.LevelData(i).dfCoin
'            oCurWorldData.LevelData(i).dfVine = oTMP.LevelData(i).dfVine
'            oCurWorldData.LevelData(i).EnemySkin = oTMP.LevelData(i).EnemySkin
'            oCurWorldData.LevelData(i).Filename = oTMP.LevelData(i).Filename
'            oCurWorldData.LevelData(i).LevelName = oTMP.LevelData(i).LevelName
'            oCurWorldData.LevelData(i).MusicFile = oTMP.LevelData(i).MusicFile
'            oCurWorldData.LevelData(i).TileFile = oTMP.LevelData(i).TileFile
'            oCurWorldData.LevelData(i).TimeGiven = oTMP.LevelData(i).TimeGiven
'            For u = 1 To 16
'                oCurWorldData.LevelData(i).PipeDest(u) = oTMP.LevelData(i).PipeDest(u)
'            Next u
'            oCurWorldData.LevelData(i).iScrollStyle = 0
'            oCurWorldData.LevelData(i).iScrollSpeed = 0
'        Next i
    cwdLoadWorldData = True
errOut:
    Close fFile
End Function
