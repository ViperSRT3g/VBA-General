Attribute VB_Name = "mod_CustomDocumentProperties"
Option Explicit

Public Function GetCPropV(ByVal PropertyName As String, Optional ByRef TargetWorkbook As Workbook) As Variant
    'Returns the value of the CProp (If it exists)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ThisWorkbook
    If CPropExists(PropertyName, TargetWorkbook) Then GetCPropV = TargetWorkbook.CustomDocumentProperties.Item(PropertyName).Value
End Function

Public Function AddCProp(ByVal PropertyName As String, _
                            ByVal LinkToContent As Boolean, _
                            Optional ByVal PropertyType As MsoDocProperties, _
                            Optional ByVal PropertyValue As Variant, _
                            Optional ByRef TargetWorkbook As Workbook, _
                            Optional ByVal LinkSource As Variant) As Boolean
    'Returns TRUE if CProp is added, FALSE if not added (It may already exist)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ThisWorkbook
    If Not CPropExists(PropertyName, TargetWorkbook) Then
        TargetWorkbook.CustomDocumentProperties.Add PropertyName, LinkToContent, PropertyType, PropertyValue, LinkSource
        AddCProp = CPropExists(PropertyName, TargetWorkbook)
    End If
End Function

Public Function DeleteCProp(ByVal PropertyName As String, Optional ByRef TargetWorkbook As Workbook) As Boolean
    'Returns TRUE if CProp is successfully deleted
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ThisWorkbook
    If CPropExists(PropertyName, TargetWorkbook) Then TargetWorkbook.CustomDocumentProperties.Item(PropertyName).Delete
    DeleteCProp = Not CPropExists(PropertyName, TargetWorkbook)
End Function

Public Function CPropExists(ByVal PropertyName As String, Optional ByRef TargetWorkbook As Workbook) As Boolean
    'Returns TRUE if the CProp exists
    On Error Resume Next
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ThisWorkbook
    Dim ObjProperty As DocumentProperty: Set ObjProperty = TargetWorkbook.CustomDocumentProperties.Item(PropertyName)
    CPropExists = Not ObjProperty Is Nothing
    Set ObjProperty = Nothing
End Function
