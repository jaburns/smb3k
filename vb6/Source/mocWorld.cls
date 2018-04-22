VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oWorldMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mapNode
    zxPos As Integer
    zyPos As Integer
    zpassThrough As Boolean
    zleftNode As Byte
    zrightNode As Byte
    zupNode As Byte
    zdownNode As Byte
    znodeTag As Byte
    znodeImage As Byte
    zentryPoint As Byte
    zexitWorld As Byte
End Type
Private oNodes() As mapNode



Public Property Get nodeCount() As Long
    nodeCount = UBound(oNodes)
End Property
Public Property Get xPos(ID As Long) As Integer
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    xPos = oNodes(ID).zxPos
End Property
Public Property Get yPos(ID As Long) As Integer
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    yPos = oNodes(ID).zyPos
End Property
Public Sub moveNode(ID As Long, iX As Integer, iY As Integer)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zxPos = iX
    oNodes(ID).zyPos = iY
End Sub

Public Property Get PassThru(ID As Long) As Boolean
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    PassThru = oNodes(ID).zpassThrough
End Property
Public Property Get LeftNode(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    LeftNode = oNodes(ID).zleftNode
End Property
Public Property Get RightNode(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    RightNode = oNodes(ID).zrightNode
End Property
Public Property Get UpNode(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    UpNode = oNodes(ID).zupNode
End Property
Public Property Get DownNode(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    DownNode = oNodes(ID).zdownNode
End Property
Public Property Get NodeTag(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    NodeTag = oNodes(ID).znodeTag
End Property
Public Property Get NodeImage(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    NodeImage = oNodes(ID).znodeImage
End Property
Public Property Get EntryMode(ID As Long) As Integer  '0-4
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    EntryMode = (oNodes(ID).zentryPoint - 1) Mod 5
End Property
Public Property Get EntryTag(ID As Long) As Integer '0-4
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    If oNodes(ID).zentryPoint = 0 Then Exit Property
    EntryTag = ((oNodes(ID).zentryPoint - 1) \ 5) + 1
End Property
Public Property Get warpExitWorld(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    warpExitWorld = oNodes(ID).zexitWorld
End Property
Public Property Get EntryPoint(ID As Long) As Byte
    If ID < 1 Or ID > UBound(oNodes) Then Exit Property
    EntryPoint = oNodes(ID).zentryPoint
End Property



Public Sub setPassThru(ID As Long, bVal As Boolean)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zpassThrough = bVal
End Sub
Public Sub setLeftNode(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zleftNode = bVal
End Sub
Public Sub setRightNode(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zrightNode = bVal
End Sub
Public Sub setUpNode(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zupNode = bVal
End Sub
Public Sub setDownNode(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zdownNode = bVal
End Sub
Public Sub setNodeTag(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).znodeTag = bVal
End Sub
Public Sub warpSetExitWorld(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).zexitWorld = bVal
End Sub
Public Sub setNodeImage(ID As Long, bVal As Byte)
    If ID < 1 Or ID > UBound(oNodes) Then Exit Sub
    oNodes(ID).znodeImage = bVal
End Sub
Public Sub setEntryPoint(ID As Long, bMode As Byte, bTag As Byte) 'mode=0-4   tag=1-16
    
End Sub
Public Sub setEntryPointValue(ID As Long, bVal As Byte)
    If bVal < 0 Or bVal > 80 Then Exit Sub
    oNodes(ID).zentryPoint = bVal
End Sub




Private Sub Class_Initialize()
    ReDim oNodes(1 To 1)
    oNodes(1) = buildNode(100, 100, False, 0, 0, 0, 0, 0)
End Sub



Public Sub AddNode(iX As Integer, iY As Integer, bpassthrough As Boolean)
    ReDim Preserve oNodes(1 To UBound(oNodes) + 1)
    oNodes(UBound(oNodes)) = buildNode(iX, iY, False, 0, 0, 0, 0, 0)
End Sub
Public Sub removeNode()
    If UBound(oNodes) > 1 Then ReDim Preserve oNodes(1 To UBound(oNodes) - 1)
End Sub




Public Sub saveMap(sPath As String)
Dim fFile As Long: fFile = FreeFile
Dim tByte As Byte
Dim tInt As Integer
Dim i As Long

    Open sPath For Binary Access Write Lock Read Write As fFile
        tByte = UBound(oNodes)
        Put fFile, 1, tByte
        For i = 1 To UBound(oNodes)
        With oNodes(i)
            tInt = .zxPos
            Put fFile, , tInt
            tInt = .zyPos
            Put fFile, , tInt
            tByte = .zupNode
            Put fFile, , tByte
            tByte = .zdownNode
            Put fFile, , tByte
            tByte = .zleftNode
            Put fFile, , tByte
            tByte = .zrightNode
            Put fFile, , tByte
            tByte = getTag(.znodeTag, .zpassThrough)
            Put fFile, , tByte
            tByte = .znodeImage
            Put fFile, , tByte
            tByte = .zentryPoint
            Put fFile, , tByte
            tByte = .zexitWorld
            Put fFile, , tByte
        End With
        Next i
    Close fFile

errOut:
Close fFile
End Sub
Private Function getTag(ByVal bNodeTag As Byte, bpassthrough As Boolean) As Byte
    If bNodeTag = 0 Then getTag = 0 Else getTag = bNodeTag + 1
    If bpassthrough Then getTag = getTag + 128
End Function



Public Sub loadMap(sPath As String)
Dim fFile As Long: fFile = FreeFile
Dim tByte As Byte
Dim tInt As Integer
Dim i As Long

    Open sPath For Input As fFile
    Close fFile
    Open sPath For Binary Access Read Lock Write As fFile
        Get fFile, 1, tByte
        ReDim oNodes(1 To tByte)
        For i = 1 To UBound(oNodes)
        With oNodes(i)
            Get fFile, , tInt
            .zxPos = tInt
            Get fFile, , tInt
            .zyPos = tInt
            Get fFile, , tByte
            .zupNode = tByte
            Get fFile, , tByte
            .zdownNode = tByte
            Get fFile, , tByte
            .zleftNode = tByte
            Get fFile, , tByte
            .zrightNode = tByte
            Get fFile, , tByte
            If tByte >= 128 Then
                tByte = tByte - 128
                .zpassThrough = True
            End If
            If tByte > 0 Then
                tByte = tByte - 1
                .znodeTag = tByte
            End If
            Get fFile, , tByte
            .znodeImage = tByte
            Get fFile, , tByte
            .zentryPoint = tByte
            Get fFile, , tByte
            .zexitWorld = tByte
        End With
        Next i
    Close fFile

errOut:
Close fFile
End Sub




Private Function buildNode(iX As Integer, iY As Integer, bPass As Boolean, nLeft As Byte, nRight As Byte, nUp As Byte, nDown As Byte, nTag As Byte, Optional nWorld As Byte = 0) As mapNode
With buildNode
.zxPos = iX
.zyPos = iY
.zleftNode = nLeft
.zrightNode = nRight
.zupNode = nUp
.zdownNode = nDown
.znodeTag = nTag
.zpassThrough = bPass
.zexitWorld = nWorld
End With
End Function
