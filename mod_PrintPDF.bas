Attribute VB_Name = "mod_PrintPDF"
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" ( _
    ByVal lpFile As String, _
    ByVal lpDirectory As String, _
    ByVal lpResult As String) As Long

Public Sub PrintPDF(ByVal FilePath As String)
    Call Shell("""" & Trim(DefaultApp(FilePath)) & """ /p """ & FilePath & """")
    Call Application.Wait(Now() + TimeValue("0:00:03"))
    Call Application.SendKeys("{RETURN}", True)
    Call Application.Wait(Now() + TimeValue("0:00:05"))
    Call CloseAdobeReader(FileName(DefaultApp(FilePath)))
End Sub

Private Sub CloseAdobeReader(ByVal AdobeProcess As String)
    Call Shell("taskkill /F /IM " & AdobeProcess)
End Sub

Private Function DefaultPrinter() As String
    DefaultPrinter = Trim(Left(Application.ActivePrinter, Len(Application.ActivePrinter) - (Len(Application.ActivePrinter) - InStr(1, Application.ActivePrinter, " on Ne") - 3)))
End Function

Private Function DefaultApp(ByVal FilePath As String) As String
    Dim ExecutablePath As String: ExecutablePath = String(260, " ")
    Call FindExecutable(FileName(FilePath), BasePath(FilePath), ExecutablePath)
    DefaultApp = RegExer(ExecutablePath, "([a-zA-Z0-9\s\(\)\\\:\.\$]+)")
End Function

Private Function BasePath(ByVal Path As String) As String
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    If Len(Path) > 0 Then BasePath = Left(FSO.GetAbsolutePathName(Path), Len(FSO.GetAbsolutePathName(Path)) - Len(FileName(Path)))
    Set FSO = Nothing
End Function

Private Function BaseName(ByVal Path As String) As String
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    If Len(Path) > 0 Then BaseName = FSO.GetBaseName(Path)
    Set FSO = Nothing
End Function

Private Function FileName(ByVal Path As String) As String
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    If Len(Path) > 0 Then FileName = FSO.GetFileName(Path)
    Set FSO = Nothing
End Function

Private Function RegExer(ByVal RawData As String, ByVal RegExPattern As String) As String
    Dim RegEx As New RegExp
    Dim Matches As Object
    Dim Result As String
    With RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = RegExPattern
    End With
    If RegEx.test(RawData) Then
        Set Matches = RegEx.Execute(RawData)
        RegExer = Matches(0)
        Set Matches = Nothing
    Else
        RegExer = ""
    End If
End Function
