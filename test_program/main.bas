Attribute VB_Name = "main"
Option Explicit

Private someNumber As Long
Public someGlobal As Single

Public Sub __main__()
    someNumber = 2
    CallOtherSub
    someNumber = someNumber + 10
    someGlobal = someNumber
    CallOtherSub
End Sub