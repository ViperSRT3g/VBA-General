Attribute VB_Name = "mod_CustomDocumentProperties"
Option Explicit

Public Enum MsoDocProperties
    msoPropertyTypeNumber = 1
    msoPropertyTypeBoolean = 2
    msoPropertyTypeDate = 3
    msoPropertyTypeString = 4
    msoPropertyTypeFloat = 5
End Enum

Public Function AddCustomDocumentProperty(ByVal PropertyName As String, _
                            ByVal LinkToContent As Boolean, _
                            Optional ByVal PropertyType As MsoDocProperties, _
                            Optional ByVal PropertyValue As Variant, _
                            Optional ByRef Targetworkbook As Workbook, _
                            Optional ByVal LinkSource As Variant) As DocumentProperty
    If Targetworkbook Is Nothing Then Set Targetworkbook = ThisWorkbook
    Set AddCustomDocumentProperty = Targetworkbook.CustomDocumentProperties.Add(PropertyName, LinkToContent, PropertyType, PropertyValue, LinkSource)
End Function

Public Function DeleteCDocProp(ByRef Targetworkbook As Workbook, ByVal PropertyName As String) As Boolean
    Targetworkbook.CustomDocumentProperties.Item(PropertyName).Delete
End Function

Public Function CustomPropertyExists(ByVal PropertyName As String, Optional ByRef Targetworkbook As Workbook) As Boolean
    On Error Resume Next
    If Targetworkbook Is Nothing Then Set Targetworkbook = ThisWorkbook
    Dim ObjProperty As Object: Set ObjProperty = Targetworkbook.CustomDocumentProperties.Item(PropertyName)
    CustomPropertyExists = (Err = 0)
    Set ObjProperty = Nothing
    On Error GoTo 0
End Function
