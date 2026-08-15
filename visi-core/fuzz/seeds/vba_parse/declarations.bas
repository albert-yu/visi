Attribute VB_Name = "Module1"
Option Explicit
Private Const MAX As Long = &HFF
Public x As String, y(1 To 5) As Double, z()
Dim WithEvents app As Application
Private Type Point
    X As Long
End Type
Public Enum Color
    Red = 1
End Enum
Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As Long) As Long
Public Function F(ByVal a As Long, Optional b As String = "d") As Long
    F = -2 ^ 2 + a Mod 3 \ 2 & b
End Function
