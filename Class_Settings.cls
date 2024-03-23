VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'DEPENDENCIES
'JsonConverter.bas - https://github.com/VBA-tools/VBA-JSON

Private Enum Prop
    DefaultPath_
    SettingObject_
    SettingsJSON_
    [_First] = DefaultPath_
    [_Last] = SettingsJSON_
End Enum

Private Const DEFAULT_ENVIRON As String = "AppData"
Private Const DEFAULT_FILEPATH As String = "VBASettings"

Private Properties As Object


Private Sub Class_Initialize()
    Set Properties = CreateObject("Scripting.Dictionary")
    Properties.Add Prop.DefaultPath_, VBA.vbNullString
    Properties.Add Prop.SettingObject_, CreateObject("Scripting.Dictionary")
    Properties.Add Prop.SettingsJSON_, VBA.vbNullString
    
    If Not FolderExists(BuildPath(Environ(DEFAULT_ENVIRON), DEFAULT_FILEPATH)) Then Call MakeFolder(Environ(DEFAULT_ENVIRON), DEFAULT_FILEPATH)
    If FolderExists(BuildPath(Environ(DEFAULT_ENVIRON), DEFAULT_FILEPATH)) Then Properties(Prop.DefaultPath_) = BuildPath(Environ(DEFAULT_ENVIRON), DEFAULT_FILEPATH)
End Sub
Private Sub Class_Terminate()
    If Not Properties Is Nothing Then Properties.RemoveAll
    Set Properties = Nothing
End Sub

Public Property Get DefaultPath() As String
    DefaultPath = Properties(Prop.DefaultPath_)
End Property

Public Property Get JSON() As String
    JSON = ConvertToJson(Properties(Prop.SettingObject_))
End Property
Public Property Let JSON(Value As String)
    On Error GoTo ErrorHandler
    Set Properties(Prop.SettingObject_) = ParseJson(Value)
    On Error GoTo 0
    Exit Property
ErrorHandler:
    Set Properties(Prop.SettingObject_) = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
End Property


'Returns TRUE if a JSON file is successfully deleted
Public Function FileDelete(FilePath As String) As Boolean
    If Len(FilePath) = 0 Then Exit Function
    
    Dim FName As String: FName = FileName(FilePath)
    Dim PFolder As String: PFolder = IIf(Len(ParentFolder(FilePath)) > 0, ParentFolder(FilePath), Me.DefaultPath)
    If FileExists(BuildPath(PFolder, FName)) Then FileDelete = DeleteTarget(BuildPath(PFolder, FName))
End Function

'Returns TRUE if a JSON file is successfully saved
Public Function FileSave(FilePath As String) As Boolean
    If Len(FilePath) = 0 Then Exit Function
    
    Dim Init As String
    Dim FName As String: FName = FileName(FilePath)
    Dim PFolder As String: PFolder = IIf(Len(ParentFolder(FilePath)) > 0, ParentFolder(FilePath), Me.DefaultPath)
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(BuildPath(PFolder, FName)) Then Init = .OpenTextFile(BuildPath(PFolder, FName), 1).ReadAll
        If .FileExists(BuildPath(PFolder, FName)) Then
            Init = .OpenTextFile(BuildPath(PFolder, FName), 1).ReadAll
            .OpenTextFile(BuildPath(PFolder, FName), 2).Write Me.JSON
        Else
            .CreateTextFile(BuildPath(PFolder, FName)).Write Me.JSON
        End If
        
        Dim Result As String: Result = .OpenTextFile(BuildPath(PFolder, FName), 1).ReadAll
        FileSave = (Not Init = Result) Or (Me.JSON = Result)
    End With
End Function

'Returns TRUE if a JSON file is successfully loaded
Public Function FileLoad(FilePath As String) As Boolean
    If Len(FilePath) = 0 Then Exit Function
    
    Dim JSON As String
    Dim Init As String: Init = Me.JSON
    Dim FName As String: FName = FileName(FilePath)
    Dim PFolder As String: PFolder = IIf(Len(ParentFolder(FilePath)) > 0, ParentFolder(FilePath), Me.DefaultPath)
    If FileExists(BuildPath(PFolder, FName)) Then JSON = CreateObject("Scripting.FileSystemObject").OpenTextFile(BuildPath(PFolder, FName), 1).ReadAll
    
    If Len(JSON) > 0 Then Me.JSON = JSON
    FileLoad = (Not Init = Me.JSON) Or (Me.JSON = JSON)
End Function


'Returns TRUE if a Setting of a given name and value are successfully added
Public Function Add(SettingName As String, Value As Variant) As Boolean
    If Len(SettingName) = 0 Then Exit Function
    
    Dim Init As Boolean: Init = Properties(Prop.SettingObject_).Exists(SettingName)
    If Not Properties(Prop.SettingObject_).Exists(SettingName) Then Properties(Prop.SettingObject_).Add SettingName, Value
    Add = Not (Init = Properties(Prop.SettingObject_).Exists(SettingName))
End Function

'Returns TRUE if a Setting of a given name and value are successfully set
Public Function Edit(SettingName As String, Value As Variant, Optional AutoCreate As Boolean = False) As Boolean
    If Len(SettingName) = 0 Then Exit Function
    
    Dim Init As String: Init = Me.JSON
    If Properties(Prop.SettingObject_).Exists(SettingName) Then
        Properties(Prop.SettingObject_).Item(SettingName) = Value
    ElseIf Not Properties(Prop.SettingObject_).Exists(SettingName) And AutoCreate Then
        Properties(Prop.SettingObject_).Add SettingName, Value
    End If
    Edit = (Not Init = Me.JSON) Or (Me.JSON = ConvertToJson(Properties(Prop.SettingObject_)))
End Function

'Returns TRUE if a setting of a given name exists
Public Function Exists(SettingName As String) As Boolean
    If Len(SettingName) = 0 Then Exit Function
    
    Exists = Properties(Prop.SettingObject_).Exists(SettingName)
End Function

'Returns a VARIANT of a setting if it exists
Public Function Value(SettingName As String) As Variant
    If Len(SettingName) = 0 Then Exit Function
    
    If Properties(Prop.SettingObject_).Exists(SettingName) Then Value = Properties(Prop.SettingObject_).Item(SettingName)
End Function

'Returns TRUE if a setting of a given name is successfully removed
Public Function Remove(SettingName As String) As Boolean
    If Len(SettingName) = 0 Then Exit Function
    
    Dim Init As Boolean: Init = Properties(Prop.SettingObject_).Exists(SettingName)
    If Properties(Prop.SettingObject_).Exists(SettingName) Then Properties(Prop.SettingObject_).Remove SettingName
    Remove = Not (Init = Properties(Prop.SettingObject_).Exists(SettingName))
End Function



Private Function BuildPath(ByVal Directory As String, ByVal AdditionalPath As String) As String
    BuildPath = CreateObject("Scripting.FileSystemObject").BuildPath(Directory, AdditionalPath)
End Function

Private Function DeleteTarget(ByVal Path As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(Path) Then .GetFile(Path).Delete
        DeleteTarget = Not .FileExists(Path)
    End With
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(FilePath)
End Function

Private Function FileName(ByVal FilePath As String) As String
    FileName = CreateObject("Scripting.FileSystemObject").GetFileName(FilePath)
End Function

Private Function FolderExists(ByVal FilePath As String) As Boolean
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(FilePath)
End Function

Private Function MakeFolder(ByVal Directory As String, ByVal NewFolder As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        Dim FolderObject As Object, SubFolderObject As Object, NewFolderObject As Object
        
        If .FolderExists(BuildPath(Directory, NewFolder)) = False Then
            Set FolderObject = .GetFolder(Directory)
            Set SubFolderObject = FolderObject.SubFolders
            Set NewFolderObject = SubFolderObject.Add(NewFolder)
        End If
        MakeFolder = .FolderExists(BuildPath(Directory, NewFolder))
        
        Set FolderObject = Nothing
        Set SubFolderObject = Nothing
        Set NewFolderObject = Nothing
    End With
End Function

Private Function ParentFolder(ByVal FilePath As String) As String
    ParentFolder = Left(FilePath, InStrRev(Left(FilePath, Len(FilePath) - 1), "\"))
End Function
