VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Builder_Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DataTypeEnum
    Boolean_
    Byte_
    Collection_
    Currency_
    Date_
    Decimal_
    Dictionary_
    Double_
    Integer_
    Long_
    LongLong_
    LongPtr_
    Object_
    Single_
    String_
    Variant_
    [_First] = Boolean_
    [_Last] = Variant_
End Enum

Public Enum DataTypeInfoEnum
    name_
    ProperName_
    DefaultValue_
    IsObject_
    [_First] = name_
    [_Last] = IsObject_
End Enum

Private Enum Prop
    PropName_
    DataType_
    Initialization_
    PropGet_
    PropLetSet_
    PString_
    PCase_
    [_First] = PropName_
    [_Last] = PCase_
End Enum

Private Properties As Object


Private Sub Class_Initialize()
    Set Properties = CreateObject("Scripting.Dictionary")
    Properties.Add P(PropName_), VBA.VBNullString
    Properties.Add P(DataType_), -1
End Sub
Private Sub Class_Terminate()
    If Not Properties Is Nothing Then Properties.RemoveAll
    
    Set Properties = Nothing
End Sub


Property Get Name() As String: Name = Properties(P(PropName_)): End Property
Property Let Name(Value As String)
    Properties(P(PropName_)) = Value
End Property

Property Get DataType() As DataTypeEnum: DataType = Properties(P(DataType_)): End Property
Property Let DataType(Value As DataTypeEnum)
    Properties(P(DataType_)) = Value
End Property

Property Get EnumName() As String: EnumName = Properties(P(PropName_)) & "_": End Property
Property Get Initialization() As String: Initialization = GetString(Initialization_): End Property
Property Get PropertyGet() As String: PropertyGet = GetString(PropGet_): End Property
Property Get PropertyLetSet() As String: PropertyLetSet = GetString(PropLetSet_): End Property
Property Get PString() As String: PString = GetString(PString_): End Property
Property Get ParseCase() As String: ParseCase = GetString(PCase_): End Property
Property Get HasObject() As Boolean
    HasObject = CBool(GetDataTypeInfo(Properties(P(DataType_)), IsObject_))
End Property

Public Function GetDataTypeEnum(DataTypeName As String) As DataTypeEnum
    Select Case DataTypeName
        Case "Boolean": GetDataTypeEnum = DataTypeEnum.Boolean_
        Case "Byte": GetDataTypeEnum = DataTypeEnum.Byte_
        Case "Collection": GetDataTypeEnum = DataTypeEnum.Collection_
        Case "Currency": GetDataTypeEnum = DataTypeEnum.Currency_
        Case "Date": GetDataTypeEnum = DataTypeEnum.Date_
        Case "Decimal": GetDataTypeEnum = DataTypeEnum.Decimal_
        Case "Dictionary": GetDataTypeEnum = DataTypeEnum.Dictionary_
        Case "Double": GetDataTypeEnum = DataTypeEnum.Double_
        Case "Integer": GetDataTypeEnum = DataTypeEnum.Integer_
        Case "Long": GetDataTypeEnum = DataTypeEnum.Long_
        Case "LongLong": GetDataTypeEnum = DataTypeEnum.LongLong_
        Case "LongPtr": GetDataTypeEnum = DataTypeEnum.LongPtr_
        Case "Object": GetDataTypeEnum = DataTypeEnum.Object_
        Case "Single": GetDataTypeEnum = DataTypeEnum.Single_
        Case "String": GetDataTypeEnum = DataTypeEnum.String_
        Case "Variant": GetDataTypeEnum = DataTypeEnum.Variant_
    End Select
End Function


Private Function GetDataTypeInfo(DType As DataTypeEnum, ReturnType As DataTypeInfoEnum) As String
    Const DTYPE_BOOL    As String = "Boolean|Boolean|FALSE|FALSE"
    Const DTYPE_BYTE    As String = "Byte|Byte|0|FALSE"
    Const DTYPE_COLL    As String = "Collection|VBA.Collection|New VBA.Collection|TRUE"
    Const DTYPE_CURR    As String = "Currency|Double|0#|FALSE"
    Const DTYPE_DATE    As String = "Date|Date|0#|FALSE"
    Const DTYPE_DEC     As String = "Decimal|Double|0#|FALSE"
    Const DTYPE_DICT    As String = "Dictionary|Scripting.Dictionary|New Scripting.Dictionary|TRUE"
    Const DTYPE_DBL     As String = "Double|Double|0#|FALSE"
    Const DTYPE_INT     As String = "Integer|Integer|0|FALSE"
    Const DTYPE_LONG    As String = "Long|Long|0|FALSE"
    Const DTYPE_LLNG    As String = "LongLong|LongLong|0|FALSE"
    Const DTYPE_LNGP    As String = "LongPtr|LongPtr|0|FALSE"
    Const DTYPE_OBJ     As String = "Object|Object|Nothing|TRUE"
    Const DTYPE_SNG     As String = "Single|Single|0!|FALSE"
    Const DTYPE_STR     As String = "String|String|VBA.VBNullString|FALSE"
    Const DTYPE_VAR     As String = "Variant|Variant|Empty|FALSE"
    
    Dim EvalStr As String
    Dim Data() As String
    
    Select Case DType
        Case DataTypeEnum.Boolean_: EvalStr = DTYPE_BOOL
        Case DataTypeEnum.Byte_: EvalStr = DTYPE_BYTE
        Case DataTypeEnum.Collection_: EvalStr = DTYPE_COLL
        Case DataTypeEnum.Currency_: EvalStr = DTYPE_CURR
        Case DataTypeEnum.Date_: EvalStr = DTYPE_DATE
        Case DataTypeEnum.Decimal_: EvalStr = DTYPE_DEC
        Case DataTypeEnum.Dictionary_: EvalStr = DTYPE_DICT
        Case DataTypeEnum.Double_: EvalStr = DTYPE_DBL
        Case DataTypeEnum.Integer_: EvalStr = DTYPE_INT
        Case DataTypeEnum.Long_: EvalStr = DTYPE_LONG
        Case DataTypeEnum.LongLong_: EvalStr = DTYPE_LLNG
        Case DataTypeEnum.LongPtr_: EvalStr = DTYPE_LNGP
        Case DataTypeEnum.Object_: EvalStr = DTYPE_OBJ
        Case DataTypeEnum.Single_: EvalStr = DTYPE_SNG
        Case DataTypeEnum.String_: EvalStr = DTYPE_STR
        Case DataTypeEnum.Variant_: EvalStr = DTYPE_VAR
    End Select
    
    Data = Split(EvalStr, "|")
    Select Case ReturnType
        Case DataTypeInfoEnum.name_: GetDataTypeInfo = Data(0)
        Case DataTypeInfoEnum.ProperName_: GetDataTypeInfo = Data(1)
        Case DataTypeInfoEnum.DefaultValue_: GetDataTypeInfo = Data(2)
        Case DataTypeInfoEnum.IsObject_: GetDataTypeInfo = Data(3)
    End Select
End Function

'Returns a STRING containing the formatted code for the selected return type
Private Function GetString(ReturnType As Prop) As String
    If Len(Properties(P(PropName_))) = 0 And Properties(P(DataType_)) = -1 Then Exit Function
    
    Dim DType As DataTypeEnum: DType = Properties(P(DataType_))
    
    Select Case ReturnType
        Case Prop.Initialization_
            GetString = "Properties.Add P(" & Me.EnumName & "), " & GetDataTypeInfo(DType, DefaultValue_)
            
        Case Prop.PropGet_
            GetString = "Property Get " & Me.Name & "() As " & GetDataTypeInfo(DType, name_) & _
                        ": " & IIf(CBool(GetDataTypeInfo(DType, IsObject_)), "Set ", VBA.VBNullString) & Me.Name & " = Properties(P(" & Me.EnumName & ")): End Property"
            
        Case Prop.PropLetSet_
            If CBool(GetDataTypeInfo(DType, IsObject_)) Then
                GetString = "Property Set " & Me.Name & "(Target As " & GetDataTypeInfo(DType, name_) & "):Set Properties(P(" & Me.EnumName & ")) = Target:End Property"
            Else
                GetString = "Property Let " & Me.Name & "(Value As " & GetDataTypeInfo(DType, name_) & "):Properties(P(" & Me.EnumName & ")) = Value:End Property"
            End If

        Case Prop.PString_
            GetString = "Case Prop." & Me.EnumName & ": P = """ & Me.Name & """"
            
        Case Prop.PCase_
            If CBool(GetDataTypeInfo(DType, IsObject_)) Then
                GetString = "Case P(" & Me.EnumName & "):Set Properties(P(" & Me.EnumName & ")) = JSON.Items()(Index)"
            Else
                GetString = "Case P(" & Me.EnumName & "):Properties(P(" & Me.EnumName & ")) = JSON.Items()(Index)"
            End If
    End Select
End Function

Private Function P(Value As Prop) As String
    Select Case Value
        Case Prop.PropName_: P = "PropName"
        Case Prop.DataType_: P = "DataType"
    End Select
End Function
