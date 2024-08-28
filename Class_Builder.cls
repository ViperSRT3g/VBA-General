VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Builder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PRIVATE CONSTS
Private Const INDENT As String = "    "

Private Const HEAD_VER          As String = "VERSION 1.0 CLASS"
Private Const HEAD_BEGIN        As String = "BEGIN"
Private Const HEAD_MULTI        As String = "  MultiUse = -1  'True"
Private Const HEAD_END          As String = "END"
Private Const HEAD_NAME         As String = "Attribute VB_Name = ""Class_Template"""
Private Const HEAD_GLOBALNS     As String = "Attribute VB_GlobalNameSpace = False"
Private Const HEAD_CREATABLE    As String = "Attribute VB_Creatable = False"
Private Const HEAD_PREDECLARED  As String = "Attribute VB_PredeclaredId = False"
Private Const HEAD_EXPOSED      As String = "Attribute VB_Exposed = False"
Private Const HEAD_FLAG_EXP     As String = "Option Explicit"

Private Const SECT_DESC As String = "'MODULE DESCRIPTION"
Private Const SECT_MDEP As String = "'MODULE DEPENDENCIES"
Private Const SECT_PUE  As String = "'PUBLIC ENUMS"
Private Const SECT_PRE  As String = "'PRIVATE ENUMS"
Private Const SECT_PUC  As String = "'PUBLIC CONSTS"
Private Const SECT_PRC  As String = "'PRIVATE CONSTS"
Private Const SECT_PRD  As String = "'PRIVATE DECLARATIONS"
Private Const SECT_PRV  As String = "'PRIVATE VARIABLES"
Private Const SECT_PUV  As String = "'PUBLIC VARIABLES"
Private Const SECT_PROC As String = "'CLASS PROCEDURES"
Private Const SECT_PROP As String = "'CLASS PROPERTIES"
Private Const SECT_PUF  As String = "'PUBLIC FUNCTIONS"
Private Const SECT_PRF  As String = "'PRIVATE FUNCTIONS"

Private Const CPROC_INIT1   As String = "Private Sub Class_Initialize()"
Private Const CPROC_INIT2   As String = INDENT & "Set Properties = CreateObject(""Scripting.Dictionary"")"
Private Const CPROC_SEND    As String = "End Sub"

Private Const CPROC_TERM1   As String = "Private Sub Class_Terminate()"
Private Const CPROC_TERM2   As String = INDENT & "If Not Properties Is Nothing Then Properties.RemoveAll"
Private Const CPROC_TERM3   As String = INDENT & "Set Properties = Nothing"

Private Const VAR_PROP  As String = "Public Properties As Object"

'PRIVATE VARIABLES
Private Properties As Object


'CLASS PROCEDURES
Private Sub Class_Initialize()
    Set Properties = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    If Not Properties Is Nothing Then Properties.RemoveAll
    
    Set Properties = Nothing
End Sub


'PUBLIC FUNCTIONS
'Returns TRUE if a new Class_Builder_Property was successfully created
Public Function AddProperty(PropName As String, DataTypeName As DataTypeEnum) As Boolean
    Dim Init As Long: Init = Properties.Count
    
    If Not Properties.Exists(PropName) Then
        Dim NewProp As New Class_Builder_Property
        NewProp.Name = PropName
        NewProp.DataType = DataTypeName
        Properties.Add PropName, NewProp
        Set NewProp = Nothing
    End If
    
    AddProperty = Not (Init = Properties.Count)
End Function

'Returns a STRING containing the entire Class Module code generated from the given Class_Builder_Property objects
Public Function GenClass(Optional SingleLine As Boolean = True, Optional GetOnly As Boolean, Optional LetSetOnly As Boolean) As String
    Dim Output  As String
    
    'Header Attributes and Flags of the Class Module
    Output = Output & Join(Array(HEAD_VER, HEAD_BEGIN, HEAD_MULTI, HEAD_END, HEAD_NAME, HEAD_GLOBALNS, _
                        HEAD_CREATABLE, HEAD_PREDECLARED, HEAD_EXPOSED, HEAD_FLAG_EXP), VBA.vbCrLf) & VBA.vbCrLf
    Output = Output & Join(Array(SECT_DESC, SECT_MDEP, SECT_PUE, SECT_PRE), VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf) & VBA.vbCrLf
    Output = Output & GenEnumSection & VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf
    Output = Output & Join(Array(SECT_PUC, SECT_PRC, SECT_PRD, SECT_PUV), VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf) & VBA.vbCrLf
    Output = Output & Join(Array(VAR_PROP, SECT_PRV, SECT_PROC), VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf) & VBA.vbCrLf
    Output = Output & Join(Array(CPROC_INIT1, CPROC_INIT2, GenInitSection, CPROC_SEND, CPROC_TERM1, CPROC_TERM2, CPROC_TERM3, CPROC_SEND), VBA.vbCrLf) & VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf
    Output = Output & SECT_PROP & VBA.vbCrLf & GenPropSection(SingleLine, GetOnly, LetSetOnly) & VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf
    Output = Output & Join(Array(SECT_PUF & VBA.vbCrLf & GenLoadData, SECT_PRF & VBA.vbCrLf & GenPFunction), VBA.vbCrLf & VBA.vbCrLf & VBA.vbCrLf) & VBA.vbCrLf
    GenClass = Output
End Function

Public Function GenInitSection() As String
    If Properties.Count = 0 Then Exit Function
    
    Dim Prop    As Class_Builder_Property
    Dim Output  As String
    Dim Index   As Long
    
    For Index = 0 To Properties.Count - 1
        Set Prop = Properties.Items()(Index)
        Output = Output & INDENT & Prop.Initialization & VBA.vbCrLf
    Next Index
    Set Prop = Nothing
    'Remove final VBA.vbCrLf
    Output = Left(Output, Len(Output) - 2)
    GenInitSection = Output
End Function

'Returns a STRING containing the code for the Class properties section
Public Function GenPropSection(Optional SingleLine As Boolean = True, Optional GetOnly As Boolean, Optional LetSetOnly As Boolean) As String
    If Properties.Count = 0 Then Exit Function
    
    Const ToJSON As String = "Property Get ToJSON() As String|    On Error Resume Next|    ToJSON = ConvertToJson(Properties)|    On Error GoTo 0|End Property"
    
    Dim Prop    As Class_Builder_Property
    Dim Output  As String
    Dim Data()  As String
    Dim Item    As Variant
    Dim Index   As Long
    
    For Index = 0 To Properties.Count - 1
        Set Prop = Properties.Items()(Index)
        If SingleLine Then
            If GetOnly And Not LetSetOnly Then
                Output = Output & Prop.PropertyGet & VBA.vbCrLf
                
            ElseIf Not GetOnly And LetSetOnly Then
                Output = Output & Prop.PropertyLetSet & VBA.vbCrLf
            Else
                Output = Output & Prop.PropertyGet & VBA.vbCrLf
                Output = Output & Prop.PropertyLetSet & VBA.vbCrLf
            End If
        Else
            If GetOnly And Not LetSetOnly Then
                Data = Split(Prop.PropertyGet, ":")
                Output = Output & Data(0) & VBA.vbCrLf
                Output = Output & INDENT & Data(1) & VBA.vbCrLf
                Output = Output & Data(2) & VBA.vbCrLf
                
            ElseIf Not GetOnly And LetSetOnly Then
                Data = Split(Prop.PropertyLetSet, ":")
                Output = Output & Data(0) & VBA.vbCrLf
                Output = Output & INDENT & Data(1) & VBA.vbCrLf
                Output = Output & Data(2) & VBA.vbCrLf
            Else
                Data = Split(Prop.PropertyGet, ":")
                Output = Output & Data(0) & VBA.vbCrLf
                Output = Output & INDENT & Data(1) & VBA.vbCrLf
                Output = Output & Data(2) & VBA.vbCrLf
                
                Data = Split(Prop.PropertyLetSet, ":")
                Output = Output & Data(0) & VBA.vbCrLf
                Output = Output & INDENT & Data(1) & VBA.vbCrLf
                Output = Output & Data(2) & VBA.vbCrLf
            End If
        End If
    Next Index
    Set Prop = Nothing
    
    Output = Output & VBA.vbCrLf & Join(Split(ToJSON, "|"), VBA.vbCrLf)
    
    GenPropSection = Output
End Function

'Returns a STRING containing the code for the Private Enum Prop of the Class_Template
Public Function GenEnumSection() As String
    If Properties.Count = 0 Then Exit Function
    
    Const DEF_START As String = "Private Enum Prop"
    Const DEF_END   As String = "End Enum"
    Const DEF_FIRST As String = "[_First] = "
    Const DEF_LAST  As String = "[_Last] = "
    
    Dim Prop    As Class_Builder_Property
    Dim Output  As String
    Dim PropStr As String
    Dim First   As String
    Dim Last    As String
    Dim Index   As Long
    
    For Index = 0 To Properties.Count - 1
        Set Prop = Properties.Items()(Index)
        PropStr = PropStr & INDENT & Prop.EnumName & VBA.vbCrLf
    Next Index
    Set Prop = Nothing
    
    First = Properties.Items()(0).EnumName
    Last = Properties.Items()(Properties.Count - 1).EnumName
    
    Output = DEF_START & VBA.vbCrLf
    Output = Output & PropStr
    Output = Output & INDENT & DEF_FIRST & First & VBA.vbCrLf
    Output = Output & INDENT & DEF_LAST & Last & VBA.vbCrLf
    Output = Output & DEF_END
    GenEnumSection = Output
End Function

Public Function GenLoadData() As String
    If Properties.Count = 0 Then Exit Function
    
    Const FUNC_NORMAL As String = "Public Function LoadData(ByVal JSON As Object) As Boolean|    Dim Init    As String: Init = Me.ToJSON|    Dim Index   As Long|    |    For Index = 0 To JSON.Count - 1|        Select Case JSON.Keys()(Index)|{CASE}|        End Select|    Next Index|    |    LoadData = Not (Init = Me.ToJSON)|End Function"
    Const FUNC_SIMPLE As String = "Public Function LoadData(ByVal JSON As Object) As Boolean|    Dim Init As String: Init = Me.ToJSON|    Set Properties = JSON|    LoadData = Not (Init = Me.ToJSON)|End Function"
    
    Dim Prop    As Class_Builder_Property
    Dim Output  As String
    Dim PropStr As String
    Dim Index   As Long
    
    'If any properties are objects, then output individual properties
    If HasObjects Then
        For Index = 0 To Properties.Count - 1
            Set Prop = Properties.Items()(Index)
            PropStr = PropStr & INDENT & INDENT & INDENT & Prop.ParseCase & VBA.vbCrLf
        Next Index
        Set Prop = Nothing
        
        If Len(PropStr) > 2 Then PropStr = Left(PropStr, Len(PropStr) - 2)
        Output = Replace(FUNC_NORMAL, "{CASE}", PropStr)
    Else 'If no objects, then use the simple method
        Output = FUNC_SIMPLE
    End If
    
    Output = Join(Split(Output, "|"), VBA.vbCrLf)
    GenLoadData = Output
End Function

'Returns a STRING containing the code for the P function that translates Enums to their String equivalents
Public Function GenPFunction() As String
    If Properties.Count = 0 Then Exit Function
    
    Const DEF_FUNC_HEAD As String = "Private Function P(Value As Prop) As String"
    Const DEF_CASE      As String = "Select Case Value"
    Const DEF_CASE_END  As String = "End Select"
    Const DEF_FUNC_END  As String = "End Function"
    
    Dim Prop    As Class_Builder_Property
    Dim Output  As String
    Dim PropStr As String
    Dim Index   As Long
    
    For Index = 0 To Properties.Count - 1
        Set Prop = Properties.Items()(Index)
        PropStr = PropStr & INDENT & INDENT & Prop.PString & VBA.vbCrLf
    Next Index
    Set Prop = Nothing
    
    Output = DEF_FUNC_HEAD & VBA.vbCrLf
    Output = Output & INDENT & DEF_CASE & VBA.vbCrLf
    Output = Output & PropStr
    Output = Output & INDENT & DEF_CASE_END & VBA.vbCrLf
    Output = Output & DEF_FUNC_END
    GenPFunction = Output
End Function

'Returns TRUE if an existing Class_Builder_Property was removed
Public Function RemoveProperty(PropName As String) As Boolean
    Dim Init As Long: Init = Properties.Count
    If Properties.Exists(PropName) Then Properties.Remove PropName
    RemoveProperty = Not (Init = Properties.Count)
End Function


'PRIVATE FUNCTIONS
Private Function HasObjects() As Boolean
    If Properties.Count = 0 Then Exit Function
    
    Dim Prop    As Class_Builder_Property
    Dim Index   As Long
    
    For Index = 0 To Properties.Count - 1
        Set Prop = Properties.Items()(Index)
        If Prop.HasObject Then
            HasObjects = True
            Set Prop = Nothing
            Exit Function
        End If
    Next Index
    Set Prop = Nothing
End Function
