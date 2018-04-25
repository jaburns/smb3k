Attribute VB_Name = "mSound"
Option Explicit

Private zoSound As New DXSound
Public Music As New DXMusic

Public Type udtSounds
    Jump As Long
    Swim As Long
    Coin As Long
    Bump As Long
    BreakBrick As Long
    Boing As Long
    Death As Long
    Powerup As Long
    Powerdown As Long
    Sprout As Long
    Star As Long
    Win As Long
    Fireball As Long
    Fly As Long
    OneUp As Long
    BumpOff As Long
    Pipe As Long
    Door As Long
    PowLoop As Long
    Thwomp As Long
    HardShell As Long
    Kick As Long
    EnterLevel As Long
    Floating As Long
    GetWings As Long
    ReleaseItem As Long
    QuickExitNoise As Long
    BulletBillFire As Long
    BobombExplode As Long
    GameOverSound As Long
    BossHit As Long
    BossDie As Long
    SaveSound As Long
End Type
Public Sounds As udtSounds

Public bMusicEnabled As Boolean
Public bSoundEnabled As Boolean


Public Sub LoadSounds()
With zoSound
    .Initialize frmMain_hWnd
    Sounds.Jump = .LoadSound(App.Path & "\Data\Sound\Jump.wav")
    Sounds.Swim = .LoadSound(App.Path & "\Data\Sound\Swim.wav")
    Sounds.Coin = .LoadSound(App.Path & "\Data\Sound\Coin.wav")
    Sounds.Bump = .LoadSound(App.Path & "\Data\Sound\Hit.wav")
    Sounds.BreakBrick = .LoadSound(App.Path & "\Data\Sound\Brick.wav")
    Sounds.Boing = .LoadSound(App.Path & "\Data\Sound\Boing.wav")
    Sounds.Death = .LoadSound(App.Path & "\Data\Sound\Dead.wav")
    Sounds.Powerup = .LoadSound(App.Path & "\Data\Sound\Powerup.wav")
    Sounds.Powerdown = .LoadSound(App.Path & "\Data\Sound\Powerdown.wav")
    Sounds.Sprout = .LoadSound(App.Path & "\Data\Sound\Sprout.wav")
    Sounds.Win = .LoadSound(App.Path & "\Data\Sound\Win.wav")
    Sounds.Fireball = .LoadSound(App.Path & "\Data\Sound\Fireball.wav")
    Sounds.Fly = .LoadSound(App.Path & "\Data\Sound\Fly.wav")
    Sounds.OneUp = .LoadSound(App.Path & "\Data\Sound\1up.wav")
    Sounds.BumpOff = .LoadSound(App.Path & "\Data\Sound\BumpOff.wav")
    Sounds.Pipe = .LoadSound(App.Path & "\Data\Sound\Pipe.wav")
    Sounds.Door = .LoadSound(App.Path & "\Data\Sound\Door.wav")
    Sounds.PowLoop = .LoadSound(App.Path & "\Data\Sound\PowLoop.wav")
    Sounds.Thwomp = .LoadSound(App.Path & "\Data\Sound\Thwomp.wav")
    Sounds.HardShell = .LoadSound(App.Path & "\Data\Sound\HardShell.wav")
    Sounds.Kick = .LoadSound(App.Path & "\Data\Sound\Kick.wav")
    Sounds.EnterLevel = .LoadSound(App.Path & "\Data\Sound\EnterLevel.wav")
    Sounds.Floating = .LoadSound(App.Path & "\Data\Sound\Float.wav")
    Sounds.GetWings = .LoadSound(App.Path & "\Data\Sound\GetWings.wav")
    Sounds.ReleaseItem = .LoadSound(App.Path & "\Data\Sound\ItemRelease.wav")
    Sounds.QuickExitNoise = .LoadSound(App.Path & "\Data\Sound\QuickExitSound.wav")
    Sounds.BulletBillFire = .LoadSound(App.Path & "\Data\Sound\BulletBill.wav")
    Sounds.BobombExplode = .LoadSound(App.Path & "\Data\Sound\Bobomb.wav")
    Sounds.GameOverSound = .LoadSound(App.Path & "\Data\Sound\GameOver.wav")
    Sounds.BossHit = .LoadSound(App.Path & "\Data\Sound\BossHit.wav")
    Sounds.BossDie = .LoadSound(App.Path & "\Data\Sound\BossDie.wav")
    Sounds.SaveSound = .LoadSound(App.Path & "\Data\Sound\SavePoint.wav")
End With
End Sub

Public Function SoundPlaying(ID As Long) As Boolean
    SoundPlaying = zoSound.StillPlaying(ID)
End Function

Public Sub PlaySound(ID As Long, Optional ByVal bLoopIt As Boolean = False)
If Not bSoundEnabled Then Exit Sub
With zoSound
    If .StillPlaying(ID) Then .StopSound ID
    .PlaySound ID, bLoopIt
End With
End Sub

Public Sub EndSound(ID As Long)
On Error Resume Next
    zoSound.StopSound ID
End Sub

Public Sub KillSounds()
    Set zoSound = Nothing
End Sub



Public Sub MusicPlayMusic(Optional ByVal lRepeats As Long = -1)
If Not bMusicEnabled Then Exit Sub
    Music.PlayMusic lRepeats
End Sub
Public Sub MusicStopMusic()
    Music.StopMusic
End Sub
Public Sub MusicLoadFile(sPath As String)
    Music.LoadFile sPath
End Sub
