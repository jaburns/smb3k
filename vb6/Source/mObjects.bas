Attribute VB_Name = "mObjects"
Option Explicit

Private zoBrickBreak(4) As ocBrickBreak, curBrick As Long
'Private zoBrickBreak(4) As ocBrickBreak, curBrick As Long
Private zoGotCoin(4) As ocBegottenCoin, curGotCoin As Long
Private zoCactusBall(4) As ocCactusBall, curCactusBall As Long
Private zoMushroom As ocAmanita
Private zoUpMushroom(4) As ocAmanita, curMushroom As Long
Private zoDownMushroom As ocAmanita
Private zoFlower As ocFlower
Private zoGetHammer As ocFlower
Private zoStar As ocPowerStar
Private zoBlueStar As ocBlueStar
Private zoMoonboot As ocMoonBoot
Private zoVineMaker As ocVineMaker
Private zoMariosFire(1) As ocFireball
Private zoMariosHam(1) As ocThrowHammer


Public Sub InitObjects()
Dim i As Long

    KillObjects
    For i = 0 To 4
        Set zoBrickBreak(i) = New ocBrickBreak
        Set zoGotCoin(i) = New ocBegottenCoin
        Set zoCactusBall(i) = New ocCactusBall
        Set zoUpMushroom(i) = New ocAmanita
        If i <= 1 Then
            Set zoMariosFire(i) = New ocFireball
            Set zoMariosHam(i) = New ocThrowHammer
        End If
    Next i
    Set zoMushroom = New ocAmanita
    Set zoDownMushroom = New ocAmanita
    Set zoFlower = New ocFlower
    Set zoGetHammer = New ocFlower
    Set zoStar = New ocPowerStar
    Set zoMoonboot = New ocMoonBoot
    Set zoBlueStar = New ocBlueStar
    Set zoVineMaker = New ocVineMaker

End Sub

Public Sub HandleObjects()
Dim i As Long

    zoMushroom.HandleMushroom
    zoDownMushroom.HandleMushroom
    zoFlower.HandleFlower
    zoGetHammer.HandleFlower
    zoStar.HandleStar
    zoMoonboot.HandleBoot
    zoBlueStar.HandleStar
    zoVineMaker.HandleViney
    For i = 0 To 4
        zoBrickBreak(i).HandleBrick
        zoGotCoin(i).HandleMiniCoin
        zoCactusBall(i).HandleBall
        zoUpMushroom(i).HandleMushroom
        If i <= 1 Then
            zoMariosFire(i).HandleBall
            zoMariosHam(i).HandleBall
        End If
    Next i
    
End Sub

Public Sub KillObjects()
Dim i As Long

    For i = 0 To 4
        Set zoBrickBreak(i) = Nothing
        Set zoGotCoin(i) = Nothing
        Set zoCactusBall(i) = Nothing
        Set zoUpMushroom(i) = Nothing
        If i <= 1 Then
            Set zoMariosFire(i) = Nothing
            Set zoMariosHam(i) = Nothing
        End If
    Next i
    Set zoMushroom = Nothing
    Set zoDownMushroom = Nothing
    Set zoFlower = Nothing
    Set zoGetHammer = Nothing
    Set zoStar = Nothing
    Set zoMoonboot = Nothing
    Set zoBlueStar = Nothing
    Set zoVineMaker = Nothing

End Sub



Public Sub BreakBrick(X As Single, Y As Single)
    zoBrickBreak(curBrick).CreateAt X, Y
    curBrick = curBrick + 1
    If curBrick > 4 Then curBrick = 0
End Sub

Public Sub MakeMushroom(X As Single, Y As Single, bFaceRight As Boolean, mStyle As Integer, Optional ByVal sySpeed As Single = -1, Optional ByVal bReserveItem As Boolean = False)
    If mStyle <= -1 Then zoDownMushroom.CreateAt X, Y, bFaceRight, -1, sySpeed, bReserveItem
    If mStyle = 0 Then zoMushroom.CreateAt X, Y, bFaceRight, 0, sySpeed, bReserveItem
    If mStyle >= 1 Then
        zoUpMushroom(curMushroom).CreateAt X, Y, bFaceRight, 1, sySpeed, bReserveItem
        curMushroom = curMushroom + 1
        If curMushroom > 4 Then curMushroom = 0
    End If
End Sub

Public Sub MakeStar(X As Single, Y As Single, bFaceRight As Boolean)
    zoStar.CreateAt X, Y, bFaceRight
End Sub

Public Sub MakeBlueStar(X As Single, Y As Single)
    zoBlueStar.CreateAt X, Y
End Sub

Public Sub MakeFlower(X As Single, Y As Single, Optional ByVal bReserveItem As Boolean = False)
    zoFlower.CreateAt X, Y, False, bReserveItem
End Sub

Public Sub MakeMoonboot(X As Single, Y As Single, bFaceRight As Boolean, Optional ByVal bReserveItem As Boolean = False)
    zoMoonboot.CreateAt X, Y, bFaceRight, bReserveItem
End Sub

Public Sub MakeHammerPickup(X As Single, Y As Single, Optional ByVal bReserveItem As Boolean = False)
    zoGetHammer.CreateAt X, Y, True, bReserveItem
End Sub

Public Sub MakeViney(X As Single, Y As Single)
    zoVineMaker.CreateAt X, Y
End Sub

Public Sub MakeLittleCoin(X As Single, Y As Single)
    zoGotCoin(curGotCoin).CreateAt X, Y
    curGotCoin = curGotCoin + 1
    If curGotCoin > 4 Then curGotCoin = 0
End Sub

Public Sub MakeCactusBall(X As Single, Y As Single, xVelocity As Single)
    zoCactusBall(curCactusBall).CreateAt X, Y, xVelocity
    curCactusBall = curCactusBall + 1
    If curCactusBall > 4 Then curCactusBall = 0
End Sub

Public Function MarioThrowFire(X As Single, Y As Single, bFaceRight As Boolean) As Boolean
    MarioThrowFire = True
    If Not zoMariosFire(0).isActive Then
        zoMariosFire(0).CreateAt X, Y, bFaceRight
    ElseIf Not zoMariosFire(1).isActive Then
        zoMariosFire(1).CreateAt X, Y, bFaceRight
    Else
        MarioThrowFire = False
    End If
End Function

Public Function MarioThrowHammer(X As Single, Y As Single, sXSpeed As Single) As Boolean
    MarioThrowHammer = True
    If Not zoMariosHam(0).isActive Then
        zoMariosHam(0).CreateAt X, Y, sXSpeed
    ElseIf Not zoMariosHam(1).isActive Then
        zoMariosHam(1).CreateAt X, Y, sXSpeed
    Else
        MarioThrowHammer = False
    End If
End Function




Public Sub killMarioFire(lID As Long)
zoMariosFire(lID).DestroyMe
End Sub
Public Function getMarioFireActive(lID As Long) As Boolean
getMarioFireActive = zoMariosFire(lID).isActive
End Function
Public Function getMarioFireX(lID As Long) As Single
getMarioFireX = zoMariosFire(lID).getXPos
End Function
Public Function getMarioFireY(lID As Long) As Single
getMarioFireY = zoMariosFire(lID).getYPos
End Function



Public Sub killMarioHammer(lID As Long)
zoMariosHam(lID).DestroyMe
End Sub
Public Function getMarioHammerActive(lID As Long) As Boolean
getMarioHammerActive = zoMariosHam(lID).isActive
End Function
Public Function getMarioHammerX(lID As Long) As Single
getMarioHammerX = zoMariosHam(lID).getXPos
End Function
Public Function getMarioHammerY(lID As Long) As Single
getMarioHammerY = zoMariosHam(lID).getYPos
End Function
