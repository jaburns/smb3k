VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomeClass"
Option Explicit

Private localField As Long

Public Property Get ReadValue() As Long
    ReadValue = localField
End Property

Private Sub Class_Initialize()
    localField = 123
End Sub

Public Sub SetValueTo(newValue As Long)
    localField = newValue
End Sub

Public Sub CallMethod(a As Single, b As Single)
    Console_Log a, b, SOME_CONST, localField
End Sub