Attribute VB_Name = "mod_WordVarStorage"
Option Explicit

Public ListBoxItems() As String

Private Sub Example()

    Call PokeVar("ListBoxItems", "Item1,Item2,Item3,Item4")
    ListBoxItems = Split(PeekVar("ListBoxItems"), ",")
    MsgBox ListBoxItems(2)
    
End Sub

Private Function PokeVar(ByVal VariableName As String, ByVal VariableData As String) As Variable
    Dim VariableObject As Variable
    
    For Each VariableObject In ActiveDocument.Variables
        If VariableObject.Name = VariableName Then
            VariableObject.Value = VariableData
            PokeVar = VariableObject.Index
            Exit Function
        End If
    Next VariableObject
    
    ActiveDocument.Variables.Add Name:=VariableName, Value:=VariableData
End Function

Private Function PeekVar(ByVal VariableName As String) As String
    Dim VariableObject As Variable
    
    For Each VariableObject In ActiveDocument.Variables
        If VariableObject.Name = VariableName Then
            PeekVar = VariableObject.Value
        End If
    Next VariableObject
End Function

Private Function DelVar(ByVal VariableName As String) As Boolean
    Dim VariableObject As Variable
    
    DelVar = False
    For Each VariableObject In ActiveDocument.Variables
        If VariableObject.Name = VariableName Then
            VariableObject.Delete
            DelVar = True
        End If
    Next VariableObject
End Function
