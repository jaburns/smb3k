Attribute VB_Name = "mEnemySkinFile"
Option Explicit

Public Type udtEnemySkin
    Goomba As String
    PiranaPlants As String
    BuzzyBeetle As String
    DumbKoopa As String
    SmartKoopa As String
    BumptyPenguin As String
    Spiney As String
    LittleBoo As String
    RotoDisc As String
    LavaBubble As String
    Thwomp As String
    DryBones As String
    MovingPlatform As String
    PowButton As String
    FreeCheepCheep As String
    BlockCheepCheep As String
    Layer3Lava As String
    TallCactus As String
    EnemyWings As String
    Bouncer As String
    BulletBill As String
    Bobomb As String
    BOSS_Goomboss As String
    Wiggler As String
    PowerBlock As String
    SavePoint As String
End Type
Public oEnemySkin As udtEnemySkin


Public Sub SaveEnemySkinFile(sPath As String)
On Error GoTo errOut:
Dim fFile As Long, i As Long: fFile = FreeFile
Open sPath For Binary Access Write Lock Read Write As fFile
Put fFile, 1, oEnemySkin
errOut:
Close fFile
End Sub


Public Sub LoadEnemySkinFile(sPath As String)
On Error GoTo errOut:
Dim fFile As Long, i As Long: fFile = FreeFile
Open sPath For Input As fFile: Close fFile
Open sPath For Binary Access Read Lock Write As fFile
Get fFile, 1, oEnemySkin
errOut:
Close fFile
End Sub
