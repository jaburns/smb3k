VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type zudtLevelTile
    notEmpty As Boolean
    xSrc As Byte
    ySrc As Byte
    tag As udeLTileTag
    TileEnemy As udeLTileEnemy
End Type

Private Type zudtLevelColumn
    Row() As zudtLevelTile
End Type

Private Type zudtLevel
    iStartX As Integer
    iStartY As Integer
    bSideWarp As Boolean
    iWidth As Integer
    iHeight As Integer
    iTime As Integer
    Column() As zudtLevelColumn
End Type

Private zoLevel As zudtLevel



Public Sub SetTile(X As Long, Y As Long, xSrc As Long, ySrc As Long)
On Error GoTo errOut
With zoLevel.Column(X).Row(Y)
    .xSrc = CByte(xSrc)
    .ySrc = CByte(ySrc)
    .notEmpty = True
End With
errOut:
End Sub

Public Sub EraseTile(X As Long, Y As Long)
On Error GoTo errOut
With zoLevel.Column(X).Row(Y)
    .xSrc = 0
    .ySrc = 0
    .notEmpty = False
End With
errOut:
End Sub

Public Sub SetStartPos(X As Integer, Y As Integer)
    zoLevel.iStartX = X
    zoLevel.iStartY = Y
End Sub

Public Function IsTile(X As Long, Y As Long) As Boolean
On Error Resume Next
IsTile = zoLevel.Column(X).Row(Y).notEmpty
End Function

Public Property Get SideWarp() As Boolean
    SideWarp = zoLevel.bSideWarp
End Property
Public Sub SetSideWarp(bYes As Boolean)
    zoLevel.bSideWarp = bYes
End Sub


Public Property Get StartX() As Long
    StartX = CLng(zoLevel.iStartX)
End Property

Public Property Get StartY() As Long
    StartY = CLng(zoLevel.iStartY)
End Property

Public Property Get xSrc(X As Long, Y As Long) As Long
On Error Resume Next
xSrc = CLng(zoLevel.Column(X).Row(Y).xSrc)
End Property

Public Property Get ySrc(X As Long, Y As Long) As Long
On Error Resume Next
ySrc = CLng(zoLevel.Column(X).Row(Y).ySrc)
End Property

Public Property Get TileTag(X As Long, Y As Long) As udeLTileTag
On Error Resume Next
    TileTag = zoLevel.Column(X).Row(Y).tag
End Property

Public Property Get TileEnemy(X As Long, Y As Long) As udeLTileEnemy
On Error Resume Next
    TileEnemy = zoLevel.Column(X).Row(Y).TileEnemy
End Property

Public Sub SetTileTag(X As Long, Y As Long, tag As udeLTileTag)
    zoLevel.Column(X).Row(Y).tag = tag
End Sub

Public Sub SetTileEnemy(X As Long, Y As Long, newEnemy As udeLTileEnemy)
    zoLevel.Column(X).Row(Y).TileEnemy = newEnemy
End Sub

Public Property Get width() As Long
On Error Resume Next
width = CLng(zoLevel.iWidth)
End Property

Public Property Get height() As Long
On Error Resume Next
height = CLng(zoLevel.iHeight)
End Property

Public Property Get Time() As Long
On Error Resume Next
Time = CLng(zoLevel.iTime)
End Property
Public Sub SetTime(lTime As Long)
On Error Resume Next
zoLevel.iTime = CInt(lTime)
End Sub



Public Sub SaveToFile(sPath As String)
Dim lFile As Long
lFile = FreeFile
Dim xx As Long
Dim yy As Long
Dim byt_ As Byte
Dim int_ As Integer
    
    If fileExist(sPath) Then Kill sPath
    Open sPath For Binary Access Write Lock Read Write As lFile
    With zoLevel
        
        If .bSideWarp Then byt_ = 255 Else byt_ = 0
        Put #lFile, 1, byt_
        Put #lFile, , .iStartX
        Put #lFile, , .iStartY
        Put #lFile, , .iWidth
        Put #lFile, , .iHeight
        Put #lFile, , .iTime
        
        For xx = 0 To zoLevel.iWidth
            For yy = 0 To zoLevel.iHeight
                byt_ = IIf(.Column(xx).Row(yy).notEmpty, 255, 0)
                Put #lFile, , byt_
                byt_ = CByte(.Column(xx).Row(yy).xSrc)
                Put #lFile, , byt_
                byt_ = CByte(.Column(xx).Row(yy).ySrc)
                Put #lFile, , byt_
                byt_ = CByte(.Column(xx).Row(yy).tag)
                Put #lFile, , byt_
                byt_ = CByte(.Column(xx).Row(yy).TileEnemy)
                Put #lFile, , byt_
            Next yy
        Next xx
    
    End With
    Close lFile
    
End Sub


Public Sub LoadFromFile(sPath As String)
Dim lFile As Long
lFile = FreeFile
Dim xx As Long
Dim yy As Long
Dim byt_ As Byte
Dim int_ As Integer
    
    If Not fileExist(sPath) Then Exit Sub
    Open sPath For Binary Access Read Lock Write As lFile
    With zoLevel
        
        Get #lFile, 1, byt_
        .bSideWarp = byt_ = 255
        Get #lFile, , int_
        .iStartX = int_
        Get #lFile, , int_
        .iStartY = int_
        Get #lFile, , int_
        .iWidth = int_
        Get #lFile, , int_
        .iHeight = int_
        Get #lFile, , int_
        .iTime = int_
        
        ReDim .Column(.iWidth)
        
        For xx = 0 To zoLevel.iWidth
            ReDim .Column(xx).Row(.iHeight)
            For yy = 0 To zoLevel.iHeight
                Get #lFile, , byt_
                .Column(xx).Row(yy).notEmpty = IIf(byt_ = 255, True, False)
                Get #lFile, , byt_
                .Column(xx).Row(yy).xSrc = byt_
                Get #lFile, , byt_
                .Column(xx).Row(yy).ySrc = byt_
                Get #lFile, , byt_
                .Column(xx).Row(yy).tag = CLng(byt_)
                Get #lFile, , byt_
                .Column(xx).Row(yy).TileEnemy = CLng(byt_)
            Next yy
        Next xx
    
    End With
    Close lFile

End Sub

Public Sub Resize(lWidth As Long, lHeight As Long)
Dim xx As Long

    If lWidth <= 0 Or lHeight <= 0 Then Exit Sub
    With zoLevel
        .iWidth = lWidth - 1
        .iHeight = lHeight - 1
        ReDim Preserve .Column(.iWidth)
        For xx = 0 To .iWidth
            ReDim Preserve .Column(xx).Row(.iHeight)
        Next xx
    End With
    
End Sub

'
' function used to determine if a file exists
'
Private Function fileExist(sPath As String) As Boolean
On Error Resume Next
Dim lFile As Long
lFile = FreeFile
Open sPath For Input As lFile
Close lFile
fileExist = (Err.Number = 0)
Err.Clear
End Function

