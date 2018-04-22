Attribute VB_Name = "mPhysicsDat"
Option Explicit

Public Type udtPhysics
    Gravity As Single
    JumpPower As Single
    PlayerAccel As Single
    WalkSpeed As Single
    RunSpeed As Single
    FlySpeed As Single
    FlyPower As Single
    Friction As Single
End Type
Public Physics As udtPhysics
'
'
'
'  this function reads the physics data from a file
'  and returns true if the operation was successfull
'
Public Function loadPhysicsFile(sPath As String) As Boolean
On Error GoTo errout

    'make sure the file exists
    Open sPath For Input As #1
    Close #1
    
    'open the file and read the physics data
    Open sPath For Binary Access Read Lock Write As #1
        Get #1, 1, Physics
    Close #1
    
loadPhysicsFile = True
errout:
Exit Function
End Function
'
'
'
'  this sub saves the data in the Physics object to a file
'
Public Sub savePhysicsFile(sPath As String)
On Error GoTo errout
    
    'open the file and put the physics data
    Open sPath For Binary Access Write Lock Read Write As #1
        Put #1, 1, Physics
    Close #1
    
errout:
End Sub
