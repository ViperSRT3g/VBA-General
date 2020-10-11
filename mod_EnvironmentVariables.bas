Attribute VB_Name = "mod_EnvironmentVariables"
Option Explicit

Sub EnumSEVars()
    Dim strVar As String
    Dim i As Long
    For i = 1 To 255
        strVar = Environ$(i)
        If LenB(strVar) = 0& Then Exit For
        Debug.Print strVar
    Next
End Sub
