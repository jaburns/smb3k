Attribute VB_Name = "OtherModule"
Option Explicit

Private classInstance As New SomeClass

Public Const SOME_CONST As Long = 97

Public Sub CallOtherSub()
    classInstance.CallMethod 1, someGlobal
End Sub