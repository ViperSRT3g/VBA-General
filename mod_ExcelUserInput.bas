Attribute VB_Name = "mod_ExcelUserInput"
Option Explicit

Public Function ReturnData(ByVal Message As String) As Range
    On Error Resume Next
    Set ReturnData = Application.InputBox(Prompt:=Message, Title:="Data Selection", Type:=8)
End Function
