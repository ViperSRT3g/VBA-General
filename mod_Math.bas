Attribute VB_Name = "mod_Math"
Option Explicit

Public Function DecPrecision(ByVal Number As Variant) As Integer
    If IsNumeric(Number) Then Number = CDec(Number)
    Do While Number <> Round(Number, DecPrecision): DecPrecision = DecPrecision + 1: Loop
End Function
