Attribute VB_Name = "mod_FSO"
Option Explicit

Public Enum FileAttributes
    Normal = 0
    ReadOnly = 1
    Hidden = 2
    System = 4
    Volume = 8
    Directory = 16
    Archive = 32
    Alias = 64
    Compressed = 128
End Enum

Public Enum FileIO
    FileRead = 1
    FileWrite = 2
End Enum

Public Function FileIO(ByVal IO As FileIO, ByVal FilePath As String, Optional ByVal FileData As String) As String
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim FileObject As Object
    
    If IO = FileRead Then
        If FileExists(FilePath) = True Then
            Set FileObject = FSO.OpenTextFile(FilePath, IO)
            FileIO = FileObject.ReadAll
        End If
    ElseIf IO = FileWrite Then
        If FileExists(FilePath) = True Then
            Set FileObject = FSO.OpenTextFile(FilePath, IO)
            FileObject.Write FileData
        Else
            Set FileObject = FSO.CreateTextFile(FilePath)
            FileObject.Write FileData
        End If
    End If
    
    FileObject.Close
    Set FSO = Nothing
    Set FileObject = Nothing
End Function

Public Function FileProperties(ByVal FilePath As String, ByVal FileProperty As FileAttributes) As Boolean
    If Not FileExists(FilePath) Then Exit Function
    With CreateObject("Scripting.FileSystemObject")
        Dim File As Object: Set File = .GetFile(FilePath)
        File.Attributes = FileProperty
        FileProperties = (File.Attributes = FileProperty)
        Set File = Nothing
    End With
End Function

Public Function FileHasAttribute(ByVal FileAttribute As Byte, ByVal FileProperty As FileAttributes) As Boolean
    FileHasAttribute = (FileAttribute And FileProperty) = FileProperty
End Function

Public Function GetFileProperties(ByVal FilePath As String) As Byte
    If Not FileExists(FilePath) Then Exit Function
    With CreateObject("Scripting.FileSystemObject")
        Dim File As Object: Set File = .GetFile(FilePath)
        GetFileProperties = File.Attributes
        Set File = Nothing
    End With
End Function

Public Function FileExists(ByVal FilePath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): FileExists = .FileExists(FilePath): End With
End Function

Public Function FolderExists(ByVal FilePath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): FolderExists = .FolderExists(FilePath): End With
End Function

Public Function MakeFolder(ByVal Directory As String, ByVal NewFolder As String) As Boolean
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

Public Function BuildPath(ByVal Directory As String, ByVal AdditionalPath As String) As String
    With CreateObject("Scripting.FileSystemObject"): BuildPath = .BuildPath(Directory, AdditionalPath): End With
End Function

Public Function CopyFile(ByVal FilePath As String, ByVal FileDestination As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(FilePath) Then
            Dim FileObject As Object: Set FileObject = .GetFile(FilePath)
            FileObject.Copy (FileDestination)
            CopyFile = .FileExists(FileDestination)
            Set FileObject = Nothing
        End If
    End With
End Function

Public Function CopyFolder(ByVal FolderPath As String, ByVal FolderDestination As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(FolderPath) Then
            .CopyFolder FolderPath, FolderDestination
            CopyFolder = FolderExists(FolderDestination)
        End If
    End With
End Function

Public Function DeleteTarget(ByVal Path As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(Path) Then
            Dim Target As Object: Set Target = .GetFile(Path)
            Target.Delete
            DeleteTarget = Not FileExists(Path)
            Set Target = Nothing
        End If
    End With
End Function

Public Function DriveExists(ByVal DriveLetter As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): DriveExists = .DriveExists(DriveLetter): End With
End Function

Public Function AbsolutePath(Optional ByVal Path As String) As String
    With CreateObject("Scripting.FileSystemObject"): AbsolutePath = .GetAbsolutePathName(Path): End With
End Function

Public Function BaseName(ByVal Path As String) As String
    With CreateObject("Scripting.FileSystemObject"): BaseName = IIf(Len(Path) > 0, .GetBaseName(Path), ""): End With
End Function

Public Function FileName(ByVal Path As String) As String
    With CreateObject("Scripting.FileSystemObject"): FileName = .GetFileName(Path): End With
End Function

Public Function FileExt(ByVal Path As String) As String
    With CreateObject("Scripting.FileSystemObject"): FileExt = .GetExtensionName(Path): End With
End Function

Public Function ParentFolder(ByVal FilePath As String) As String
    ParentFolder = Left(FilePath, InStrRev(Left(FilePath, Len(FilePath) - 1), "\"))
End Function

Public Function RenameFile(ByVal OldName As String, ByVal NewName As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If Not .FileExists(OldName) Then Exit Function
        .MoveFile OldName, NewName
        RenameFile = .FileExists(NewName)
    End With
End Function

Public Function FiletoArray(ByVal FilePath As String) As Byte()
    If Len(FilePath) = 0 Or Not FileExists(FilePath) Then Exit Function
    ReDim FiletoArray(FileLen(FilePath)) As Byte
    Dim FileNo As Integer: FileNo = FreeFile
    Open FilePath For Binary Access Read As FileNo: Get FileNo, , FiletoArray: Close FileNo
End Function

Public Function ArrayToFile(ByVal FilePath As String, ByRef Data() As Byte) As Boolean
    If Len(FilePath) = 0 Then Exit Function
    Dim FileNo As Integer: FileNo = FreeFile
    Open FilePath For Binary Lock Read Write As FileNo: Put FileNo, , Data: Close FileNo
    With CreateObject("Scripting.FileSystemObject"): ArrayToFile = .FileExists(FilePath): End With
End Function

Public Sub SubFolderLoop(ByVal TargetFolder As String)
    Dim Folder As Object, SubFolder As Object, File As Object
    Dim FQueue As New Collection
    
    With CreateObject("Scripting.FileSystemObject")
        FQueue.Add .GetFolder(TargetFolder)
        Do While FQueue.Count > 0
            Set Folder = FQueue(1)
            FQueue.Remove 1
            'Code for individual folder
            For Each SubFolder In Folder.SubFolders
                FQueue.Add SubFolder
                'Code for individual subfolders
            Next SubFolder
            For Each File In Folder.Files
                'Code for individual files
            Next File
        Loop
    End With
End Sub
