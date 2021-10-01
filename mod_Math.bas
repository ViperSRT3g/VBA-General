Attribute VB_Name = "mod_Math"
Option Explicit

'Returns the number of decimals for any given non-whole integer
Public Function DecPrecision(ByVal Number As Variant) As Integer
    Select Case VarType(Number) 'If the passed variable is any of the following data variable types (Numeric) perform the operation
        Case 2, 3, 4, 5, 6, 7, 14, 17, 20: Do While CDec(Number) <> Round(CDec(Number), DecPrecision): DecPrecision = DecPrecision + 1: Loop
    End Select
End Function

'Returns the rounded up value with the given amount of decimal digits
Public Function RoundUp(Value As Double, Digits As Integer) As Double
    RoundUp = Application.WorksheetFunction.RoundUp(Value, Digits)
End Function

'Returns the rounded down value with the given amount of decimal digits
Public Function RoundDown(Value As Double, Digits As Integer) As Double
    RoundDown = Application.WorksheetFunction.RoundDown(Value, Digits)
End Function
