Attribute VB_Name = "mod_CustomAutoFilter"
Option Explicit

Public Sub CustomFilter(ByRef Filter() As String, Optional ByRef TargetColumn As Range)
    'Accepts an array of filter items, and an optional target to apply it to (If none, it asks for one)
    Dim Counter As Long
    Dim FilterColumn As Range
    If TargetColumn Is Nothing Then
        Set FilterColumn = ReturnData("Select column to apply custom filter to")
    Else
        Set FilterColumn = TargetColumn
    End If
    
    'Validate selections are made
    If FilterColumn Is Nothing Then
        Set FilterColumn = Nothing
        Exit Sub
    End If
    
    Call LudicrousMode(True)
    FilterColumn.Worksheet.ShowAllData
    'Apply filter array to specified column
    FilterColumn.AutoFilter Field:=TargetColumn.Column, Criteria1:=Filter, Operator:=xlFilterValues
    Call LudicrousMode(False)

    Set FilterColumn = Nothing
End Sub

Private Function NextIndex(ByRef Arr As Variant) As Long
    For NextIndex = LBound(Arr) To UBound(Arr)
        If Len(Arr(NextIndex)) = 0 Then Exit Function
    Next NextIndex
End Function

Public Sub LudicrousMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub

Public Function ReturnData(ByVal Message As String) As Range
    On Error Resume Next
    Set ReturnData = Application.InputBox(Prompt:=Message, Title:="Data Selection", Type:=8)
End Function
