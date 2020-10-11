Attribute VB_Name = "mod_RegEx"
Option Explicit

Public Function RegExer(ByVal RawData As String, ByVal RegExPattern As String) As String()
    'Outputs an array of strings for each matching expression
    Dim Matches As Object, Match As Variant, Output() As String, Index As Long
    With CreateObject("VBScript.RegExp")
        .Global = True: .MultiLine = True: .IgnoreCase = True: .Pattern = RegExPattern
        If .test(RawData) Then
            Set Matches = .Execute(RawData): ReDim Output(Matches.Count - 1) As String
            For Each Match In Matches: Output(Index) = Matches(Index): Index = Index + 1: Next Match
        Else
            ReDim Output(0) As String: RegExer = Output
        End If
        RegExer = Output
    End With
End Function

Public Function SRegEx(ByVal Pattern As String, ByVal SearchText As String, Optional MatchIndex As Integer = 1, Optional IgnoreCase As Boolean = True) As String
    'Outputs a single string match
    Dim Matches As Object, Match As Variant, Index As Long
    With CreateObject("VBScript.RegExp")
        .Global = True: .MultiLine = True: .IgnoreCase = IgnoreCase: .Pattern = Pattern
        If .test(SearchText) Then
            Set Matches = .Execute(SearchText)
            For Each Match In Matches
                Index = Index + 1
                If Index = MatchIndex Then
                    SRegEx = Match
                    Exit Function
                End If
            Next Match
        End If
    End With
End Function
