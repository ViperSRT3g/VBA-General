Attribute VB_Name = "mod_FileDialog"
Option Explicit

Public Function FileDialog(ByVal DialogType As MsoFileDialogType, _
                           Optional ByVal DialogTitle As String, _
                           Optional MultiSelect As Boolean, _
                           Optional Initial As String) As String
    'If MultiSelect then outputs files in the following format: "File1;File2;File3"
    Dim SelectedFile As Variant
    With Application.FileDialog(DialogType)
        If Len(DialogTitle) > 0 Then .Title = DialogTitle
        If Len(Initial) > 0 Then .InitialFileName = Initial & "\"
        .AllowMultiSelect = MultiSelect
        .Filters.Clear
        .Show
        For Each SelectedFile In .SelectedItems
            FileDialog = FileDialog & CStr(SelectedFile) & ";"
        Next SelectedFile
        If Len(FileDialog) > 0 Then FileDialog = Left(FileDialog, Len(FileDialog) - 1)
    End With
End Function
