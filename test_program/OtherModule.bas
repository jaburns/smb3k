Attribute VB_Name = "OtherModule"
Option Explicit

Private classInstance As New SomeClass

Public Const SOME_CONST As Long = 97

Public Sub CallOtherSub()
    classInstance.publicField = 99
    classInstance.CallMethod 1, someGlobal
    classInstance.SetValueTo classInstance.ReadValue + 10000
    classInstance.CallMethod 0, 0
End Sub